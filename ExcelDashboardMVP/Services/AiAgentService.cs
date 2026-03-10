using System.Text;
using System.Text.Json;
using ExcelDashboardMVP.Models;

namespace ExcelDashboardMVP.Services
{
    /// <summary>
    /// Singleton service that drives "Dash", the AI assistant.
    /// Calls Gemini 2.0 Flash with full app context and current data summary.
    /// Provides a command queue so actions triggered via chat work even when
    /// navigating to a new page mid-conversation.
    /// </summary>
    public class AiAgentService
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly ExcelDataService _dataService;
        private readonly IConfiguration _config;
        private readonly ILogger<AiAgentService> _logger;

        // Conversation history kept per-singleton (MVP: single-user)
        private readonly List<GeminiHistoryItem> _history = new();
        private readonly object _historyLock = new();

        // Pending command: stored when data table not yet mounted
        private DataTableCommand? _pendingCommand;

        /// <summary>Fired when an AI action targets the data table and it is currently mounted.</summary>
        public event Action<DataTableCommand>? OnDataTableCommand;

        public AiAgentService(
            IHttpClientFactory httpClientFactory,
            ExcelDataService dataService,
            IConfiguration config,
            ILogger<AiAgentService> logger)
        {
            _httpClientFactory = httpClientFactory;
            _dataService = dataService;
            _config = config;
            _logger = logger;
        }

        // ── Public API ─────────────────────────────────────────────────────

        public async Task<AiChatResponse> SendMessageAsync(string userMessage)
        {
            var apiKey = _config["Gemini:ApiKey"] ?? "";
            if (string.IsNullOrWhiteSpace(apiKey) || apiKey.StartsWith("YOUR_"))
            {
                return new AiChatResponse
                {
                    Message = "⚙️ **Setup required:** Add your Gemini API key to appsettings.json under `\"Gemini\": { \"ApiKey\": \"...\" }`. " +
                              "Get a free key at [aistudio.google.com](https://aistudio.google.com)."
                };
            }

            lock (_historyLock)
                _history.Add(new GeminiHistoryItem { Role = "user", Content = userMessage });

            try
            {
                List<GeminiHistoryItem> snapshot;
                lock (_historyLock)
                    snapshot = _history.TakeLast(20).ToList();

                var response = await CallGeminiAsync(apiKey, snapshot);

                lock (_historyLock)
                    _history.Add(new GeminiHistoryItem { Role = "model", Content = response.Message });

                return response;
            }
            catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                return new AiChatResponse { Message = "❌ Invalid Gemini API key. Please check your appsettings.json." };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gemini API call failed");
                return new AiChatResponse { Message = $"❌ Something went wrong: {ex.Message}" };
            }
        }

        /// <summary>
        /// Called by AiChatBubble when user clicks an action button.
        /// If data table is mounted it fires the event; otherwise queues the command.
        /// </summary>
        public void DispatchDataTableCommand(DataTableCommand cmd)
        {
            if (OnDataTableCommand != null)
                OnDataTableCommand.Invoke(cmd);
            else
                _pendingCommand = cmd;
        }

        /// <summary>
        /// Called by ExcelDataManager on initialisation to pick up any pending command.
        /// </summary>
        public DataTableCommand? ConsumePendingCommand()
        {
            var cmd = _pendingCommand;
            _pendingCommand = null;
            return cmd;
        }

        public void ClearHistory()
        {
            lock (_historyLock)
                _history.Clear();
        }

        // ── Gemini ─────────────────────────────────────────────────────────

        private async Task<AiChatResponse> CallGeminiAsync(string apiKey, List<GeminiHistoryItem> history)
        {
            var systemPrompt = BuildSystemPrompt();

            // Gemini requires alternating user/model turns; ensure we start with user
            var contents = history.Select(m => new
            {
                role  = m.Role,
                parts = new[] { new { text = m.Content } }
            }).ToArray();

            var requestBody = new
            {
                system_instruction = new { parts = new[] { new { text = systemPrompt } } },
                contents,
                generationConfig = new
                {
                    responseMimeType = "application/json",
                    temperature      = 0.65,
                    maxOutputTokens  = 800
                }
            };

            var url    = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={apiKey}";
            var json   = JsonSerializer.Serialize(requestBody);
            var client = _httpClientFactory.CreateClient("Gemini");

            // ── Retry up to 3 times on 429 (free-tier rate limit) ─────────────
            HttpResponseMessage httpResponse = null!;
            int[] retryDelaysMs = { 5000, 15000, 30000 };

            for (int attempt = 0; attempt <= retryDelaysMs.Length; attempt++)
            {
                // Re-create content each attempt (StringContent is single-use)
                var reqContent = new StringContent(json, Encoding.UTF8, "application/json");
                httpResponse = await client.PostAsync(url, reqContent);

                if ((int)httpResponse.StatusCode != 429)
                    break;

                if (attempt < retryDelaysMs.Length)
                {
                    _logger.LogWarning("Gemini 429 rate-limited. Retrying in {D}ms (attempt {A}/{Max})...",
                        retryDelaysMs[attempt], attempt + 1, retryDelaysMs.Length);
                    await Task.Delay(retryDelaysMs[attempt]);
                }
            }

            if ((int)httpResponse.StatusCode == 429)
                return new AiChatResponse
                {
                    Message = "Dash is busy right now (rate limit reached). Please wait a few seconds and try again."
                };

            httpResponse.EnsureSuccessStatusCode();

            var responseJson = await httpResponse.Content.ReadAsStringAsync();
            var doc          = JsonDocument.Parse(responseJson);

            var rawText = doc.RootElement
                .GetProperty("candidates")[0]
                .GetProperty("content")
                .GetProperty("parts")[0]
                .GetProperty("text")
                .GetString() ?? "{}";

            _logger.LogDebug("Gemini raw response: {R}", rawText);

            // Strip possible markdown fences
            rawText = rawText.Trim();
            if (rawText.StartsWith("```")) rawText = rawText.Split('\n', 2).LastOrDefault()?.Trim() ?? rawText;
            if (rawText.EndsWith("```"))   rawText = rawText[..^3].Trim();

            try
            {
                var opts   = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var parsed = JsonSerializer.Deserialize<AiChatResponse>(rawText, opts);
                return parsed ?? new AiChatResponse { Message = rawText };
            }
            catch
            {
                return new AiChatResponse { Message = rawText };
            }
        }

        // ── System prompt (rebuilt per-request with live data) ─────────────

        private string BuildSystemPrompt()
        {
            var records = _dataService.GetRecords();
            int total   = records.Count;

            var male    = records.Count(r => r.Sex.Equals("Male",   StringComparison.OrdinalIgnoreCase));
            var female  = records.Count(r => r.Sex.Equals("Female", StringComparison.OrdinalIgnoreCase));
            var disab   = records.Count(r => r.HasDisability);

            var topCompanies = records
                .GroupBy(r => r.HostCompany)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .OrderByDescending(g => g.Count()).Take(8)
                .Select(g => $"{g.Key} ({g.Count()})");

            var topJobs = records
                .GroupBy(r => r.JobType)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .OrderByDescending(g => g.Count()).Take(6)
                .Select(g => $"{g.Key} ({g.Count()})");

            var municipalities = records
                .Select(r => r.LocalMunicipality)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct().Take(6);

            var demographics = records
                .GroupBy(r => r.DemographicGroup)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .OrderByDescending(g => g.Count()).Take(6)
                .Select(g => $"{g.Key} ({g.Count()})");

            var topLeadCos = records
                .GroupBy(r => r.LeadCompany)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .OrderByDescending(g => g.Count()).Take(4)
                .Select(g => $"{g.Key} ({g.Count()})");

            // ⚠️ Use $$""" so that literal JSON braces { } need no escaping,
            //    and C# interpolations use {{ }} instead.
            return $$"""
            You are "Dash", a smart, friendly AI assistant embedded in the ExcelDashboard MVP application.
            This is a Blazor Server web app used by programme coordinators to manage South African employment-programme participant data (Project Phoenix Cohorts 1-4 and Project Lotus Cohort 5).

            ## Application Pages
            | Route | Description |
            |---|---|
            | /dashboard/projects | Projects dashboard — charts, participant table, team, sponsors |
            | /dashboard/jobs | Jobs dashboard — employment status donut, gender ratio, demographic breakdown |
            | /dashboard/courses | Courses dashboard — skill development programme |
            | /dashboard/projects-dark | Dark-mode version of the Projects dashboard |
            | /upload | Upload an Excel file to import participant data |
            | /data-table | Full data table — add/edit/delete rows, filter, sort, export, custom columns |

            ## Live Data Context (as of now)
            - Total participants: {{total}}
            - Male: {{male}} | Female: {{female}} | Other/unspecified: {{total - male - female}}
            - Persons with disability: {{disab}}
            - Top host companies: {{string.Join(", ", topCompanies.DefaultIfEmpty("none yet"))}}
            - Top job types: {{string.Join(", ", topJobs.DefaultIfEmpty("none yet"))}}
            - Lead companies: {{string.Join(", ", topLeadCos.DefaultIfEmpty("none yet"))}}
            - Municipalities: {{string.Join(", ", municipalities.DefaultIfEmpty("none yet"))}}
            - Demographic groups: {{string.Join(", ", demographics.DefaultIfEmpty("none yet"))}}

            ## Action System
            You can suggest actions the user can click. Include them in the "actions" array:

            ### Navigate
            {"type":"navigate","label":"View Jobs Dashboard","route":"/dashboard/jobs"}

            ### Filter data table by a column (navigates + applies filter)
            {"type":"filter","label":"Show Old Mutual staff","column":"HostCompany","value":"Old Mutual"}
            Valid columns: Name, Surname, Identifier, EmailAddress, LocalMunicipality, HostCompany, LeadCompany, JobType, DemographicGroup, Sex, ContactDetails, EmploymentStatus, PersonDisability

            ### Export filtered records (applies filter then downloads .xlsx)
            {"type":"export","label":"Export Checkers staff","column":"HostCompany","value":"Checkers"}

            ### Clear all filters
            {"type":"clear_filters","label":"Clear all filters"}

            ## Response Rules
            - ALWAYS respond with VALID JSON only — no markdown, no prose outside JSON.
            - Format EXACTLY:
              {"message":"Your helpful text here","actions":[...]}
            - Keep messages concise (2-4 sentences max). Use plain text only inside "message" — no markdown.
            - Suggest 1-3 relevant actions when they'd help the user.
            - If no actions needed, use empty array: "actions":[]
            - When the user asks to filter/export data, ALWAYS include the corresponding action button.
            - Use accurate numbers from the Live Data Context above when answering data questions.
            """;
        }
    }
}