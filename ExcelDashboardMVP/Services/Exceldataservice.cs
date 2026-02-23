using OfficeOpenXml;
using OfficeOpenXml.Style;
using ExcelDashboardMVP.Models;
using System.Drawing;

namespace ExcelDashboardMVP.Services
{
    /// <summary>
    /// Singleton in-memory store for PersonRecord data.
    /// Imports via HEADER-NAME mapping so column order in Excel doesn't matter.
    /// Notifies all subscribers (dashboards, data table) when data changes.
    /// </summary>
    public class ExcelDataService
    {
        private readonly ILogger<ExcelDataService> _logger;
        private readonly object _lock = new();
        private List<PersonRecord> _records = new();
        private int _nextId = 1;

        // Canonical column names and their accepted aliases
        private static readonly Dictionary<string, string[]> ColumnAliases = new()
        {
            ["Name"]              = new[] { "Name", "First Name", "FirstName" },
            ["Surname"]           = new[] { "Surname", "Last Name", "LastName", "Family Name" },
            ["Identifier"]        = new[] { "Identifier", "ID", "ID Number", "IDNumber" },
            ["EmailAddress"]      = new[] { "EmailAddress", "Email Address", "Email" },
            ["LocalMunicipality"] = new[] { "LocalMunicipality", "Local Municipality", "Municipality" },
            ["HostCompany"]       = new[] { "HostCompany", "Host Company", "Host" },
            ["LeadCompany"]       = new[] { "LeadCompany", "Lead Company", "Lead" },
            ["JobType"]           = new[] { "JobType", "Job Type", "Job", "Occupation" },
            ["DemographicGroup"]  = new[] { "DemographicGroup", "Demographic Group", "Demographic", "Race", "Group" },
            ["Sex"]               = new[] { "Sex", "Gender" },
            ["ContactDetails"]    = new[] { "ContactDetails", "Contact Details", "Contact", "Phone", "Cell" },
            ["EmploymentStatus"]  = new[] { "EmploymentStatus", "Employment Status", "Status" },
            ["PersonDisability"]  = new[] { "PersonDisability", "Person Disability", "Disability", "PersonWithDisability" },
        };

        public static readonly string[] RequiredColumns = { "Name", "Surname" };

        /// <summary>Fired whenever the record list changes.</summary>
        public event Action? OnDataChanged;

        public ExcelDataService(ILogger<ExcelDataService> logger)
        {
            _logger = logger;
        }

        // ── Read ────────────────────────────────────────────────────────────

        public List<PersonRecord> GetRecords()
        {
            lock (_lock) return _records.ToList();
        }

        public int GetRecordCount()
        {
            lock (_lock) return _records.Count;
        }

        // ── Import ──────────────────────────────────────────────────────────

        /// <summary>
        /// Validates headers, parses rows and replaces the in-memory dataset.
        /// Returns (count, missingColumns). Throws on hard errors.
        /// </summary>
        public async Task<(int Count, List<string> MissingColumns)> ImportFromExcelAsync(Stream stream)
        {
            var records = new List<PersonRecord>();
            var missingCols = new List<string>();

            try
            {
                using var package = new ExcelPackage(stream);
                var ws = package.Workbook.Worksheets[0];

                if (ws.Dimension == null || ws.Dimension.Rows < 2)
                {
                    _logger.LogWarning("Excel file has no data rows.");
                    return (0, new List<string> { "File appears empty or has only a header row." });
                }

                // Build header → column-index map (case-insensitive)
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= ws.Dimension.Columns; c++)
                {
                    var h = ws.Cells[1, c].Value?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(h) && !colMap.ContainsKey(h))
                        colMap[h] = c;
                }

                _logger.LogInformation("Excel headers detected: {H}", string.Join(", ", colMap.Keys));

                // Validate required columns exist
                foreach (var req in RequiredColumns)
                {
                    bool found = ColumnAliases.TryGetValue(req, out var aliases)
                        && aliases!.Any(a => colMap.ContainsKey(a));
                    if (!found) missingCols.Add(req);
                }

                if (missingCols.Any())
                {
                    _logger.LogWarning("Missing required columns: {Cols}", string.Join(", ", missingCols));
                    return (0, missingCols);
                }

                int autoId = 1;
                int skipped = 0;

                for (int row = 2; row <= ws.Dimension.Rows; row++)
                {
                    var nameVal    = GetStr(ws, row, colMap, ColumnAliases["Name"]);
                    var surnameVal = GetStr(ws, row, colMap, ColumnAliases["Surname"]);

                    // Skip entirely blank rows
                    if (string.IsNullOrWhiteSpace(nameVal) && string.IsNullOrWhiteSpace(surnameVal))
                    {
                        skipped++;
                        continue;
                    }

                    try
                    {
                        records.Add(new PersonRecord
                        {
                            RowNumber         = autoId++,
                            Name              = nameVal,
                            Surname           = surnameVal,
                            Identifier        = GetStr(ws, row, colMap, ColumnAliases["Identifier"]),
                            EmailAddress      = GetStr(ws, row, colMap, ColumnAliases["EmailAddress"]),
                            LocalMunicipality = GetStr(ws, row, colMap, ColumnAliases["LocalMunicipality"]),
                            HostCompany       = GetStr(ws, row, colMap, ColumnAliases["HostCompany"]),
                            LeadCompany       = GetStr(ws, row, colMap, ColumnAliases["LeadCompany"]),
                            JobType           = GetStr(ws, row, colMap, ColumnAliases["JobType"]),
                            DemographicGroup  = GetStr(ws, row, colMap, ColumnAliases["DemographicGroup"]),
                            Sex               = GetStr(ws, row, colMap, ColumnAliases["Sex"]),
                            ContactDetails    = GetStr(ws, row, colMap, ColumnAliases["ContactDetails"]),
                            EmploymentStatus  = GetStr(ws, row, colMap, ColumnAliases["EmploymentStatus"]),
                            PersonDisability  = GetStr(ws, row, colMap, ColumnAliases["PersonDisability"]),
                        });
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Failed to parse row {Row}. Skipping.", row);
                        skipped++;
                    }
                }

                _logger.LogInformation("Imported {Count} records, skipped {Skipped} blank/error rows.", records.Count, skipped);

                lock (_lock)
                {
                    _records = records;
                    _nextId  = records.Count > 0 ? records.Max(r => r.RowNumber) + 1 : 1;
                }

                NotifyChanged();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Fatal error reading Excel file.");
                throw;
            }

            return (records.Count, missingCols);
        }

        // ── CRUD ────────────────────────────────────────────────────────────

        public void AddRecord(PersonRecord record)
        {
            lock (_lock)
            {
                record.RowNumber = _nextId++;
                _records.Add(record);
            }
            NotifyChanged();
        }

        public bool UpdateRecord(PersonRecord updated)
        {
            bool found;
            lock (_lock)
            {
                var idx = _records.FindIndex(r => r.RowNumber == updated.RowNumber);
                found = idx >= 0;
                if (found) _records[idx] = updated;
            }
            if (found) NotifyChanged();
            return found;
        }

        public bool DeleteRecord(int rowNumber)
        {
            bool removed;
            lock (_lock)
            {
                removed = _records.RemoveAll(r => r.RowNumber == rowNumber) > 0;
            }
            if (removed) NotifyChanged();
            return removed;
        }

        public void ClearRecords()
        {
            lock (_lock)
            {
                _records.Clear();
                _nextId = 1;
            }
            NotifyChanged();
        }

        // ── Template Download ───────────────────────────────────────────────

        /// <summary>Generates a blank template .xlsx with the correct headers.</summary>
        public byte[] GenerateTemplate()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Template");

            string[] headers =
            {
                "Name", "Surname", "Identifier", "EmailAddress",
                "LocalMunicipality", "HostCompany", "LeadCompany",
                "JobType", "DemographicGroup", "Sex", "ContactDetails",
                "EmploymentStatus", "PersonDisability"
            };

            for (int c = 0; c < headers.Length; c++)
            {
                var cell = ws.Cells[1, c + 1];
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x10, 0x74, 0xD6));
                cell.Style.Font.Color.SetColor(Color.White);
            }

            // Add 3 example rows
            object[,] examples =
            {
                { "Jane",  "Smith",  "1234567890123", "jane@example.com",   "Cape Town",      "KPMG",      "Collective X", "Data Analyst",  "Youth",    "Female", "0821234567", "Employed",   "N" },
                { "John",  "Doe",    "9876543210987", "john@example.com",   "Stellenbosch",   "Old Mutual", "Collective X", "Tech Support",  "Women",    "Male",   "0831234567", "Unemployed", "N" },
                { "Thabo", "Nkosi",  "1111111111111", "thabo@example.com",  "George",         "CallLab",   "Collective X", "Cloud Admin",   "Youth",    "Male",   "0841234567", "Employed",   "Y" },
            };

            for (int r = 0; r < 3; r++)
                for (int c = 0; c < 13; c++)
                    ws.Cells[r + 2, c + 1].Value = examples[r, c];

            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            // Add notes in row 6
            ws.Cells[6, 1].Value = "Notes:";
            ws.Cells[6, 1].Style.Font.Bold = true;
            ws.Cells[7, 1].Value = "PersonDisability: Use Y/Yes for disability, N/No for none";
            ws.Cells[8, 1].Value = "Sex: Male / Female / Other";
            ws.Cells[9, 1].Value = "Column order does not matter — headers are matched by name";

            return package.GetAsByteArray();
        }

        // ── Export ──────────────────────────────────────────────────────────

        public byte[] ExportToExcel(IEnumerable<PersonRecord> records)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Participants");

            string[] headers =
            {
                "#", "Name", "Surname", "Identifier", "Email Address",
                "Local Municipality", "Host Company", "Lead Company",
                "Job Type", "Demographic Group", "Sex", "Contact Details",
                "Employment Status", "Person Disability"
            };

            for (int c = 0; c < headers.Length; c++)
            {
                var cell = ws.Cells[1, c + 1];
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x10, 0x74, 0xD6));
                cell.Style.Font.Color.SetColor(Color.White);
            }

            int row = 2;
            foreach (var r in records)
            {
                ws.Cells[row,  1].Value = r.RowNumber;
                ws.Cells[row,  2].Value = r.Name;
                ws.Cells[row,  3].Value = r.Surname;
                ws.Cells[row,  4].Value = r.Identifier;
                ws.Cells[row,  5].Value = r.EmailAddress;
                ws.Cells[row,  6].Value = r.LocalMunicipality;
                ws.Cells[row,  7].Value = r.HostCompany;
                ws.Cells[row,  8].Value = r.LeadCompany;
                ws.Cells[row,  9].Value = r.JobType;
                ws.Cells[row, 10].Value = r.DemographicGroup;
                ws.Cells[row, 11].Value = r.Sex;
                ws.Cells[row, 12].Value = r.ContactDetails;
                ws.Cells[row, 13].Value = r.EmploymentStatus;
                ws.Cells[row, 14].Value = r.PersonDisability;

                if (row % 2 == 0)
                {
                    using var rng = ws.Cells[row, 1, row, headers.Length];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xF0, 0xF4, 0xFF));
                }
                row++;
            }

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            return package.GetAsByteArray();
        }

        // ── Private helpers ─────────────────────────────────────────────────

        private void NotifyChanged() => OnDataChanged?.Invoke();

        private static string GetStr(ExcelWorksheet ws, int row,
            Dictionary<string, int> colMap, string[] aliases)
        {
            foreach (var alias in aliases)
            {
                if (colMap.TryGetValue(alias, out int col))
                    return ws.Cells[row, col].Value?.ToString()?.Trim() ?? string.Empty;
            }
            return string.Empty;
        }
    }
}