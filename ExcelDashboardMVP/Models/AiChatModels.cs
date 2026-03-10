namespace ExcelDashboardMVP.Models
{
    public class ChatMessage
    {
        public string Role { get; set; } = "user"; // "user" or "assistant"
        public string Content { get; set; } = "";
        public List<AiAction> Actions { get; set; } = new();
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public bool IsLoading { get; set; }
    }

    public class AiChatResponse
    {
        public string Message { get; set; } = "";
        public List<AiAction> Actions { get; set; } = new();
    }

    public class AiAction
    {
        public string Type { get; set; } = "";   // navigate | filter | export | clear_filters
        public string Label { get; set; } = "";
        public string? Route { get; set; }
        public string? Column { get; set; }
        public string? Value { get; set; }
    }

    public class GeminiHistoryItem
    {
        public string Role { get; set; } = "user";   // "user" or "model"
        public string Content { get; set; } = "";
    }

    public class DataTableCommand
    {
        public string Type { get; set; } = "";        // filter | export | clear_filters
        public Dictionary<string, string> Parameters { get; set; } = new();
    }
}