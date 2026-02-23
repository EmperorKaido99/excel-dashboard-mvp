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
        private List<PersonRecord> _records = new();
        private int _nextId = 1;

        /// <summary>Fired whenever the record list changes.</summary>
        public event Action? OnDataChanged;

        public ExcelDataService(ILogger<ExcelDataService> logger)
        {
            _logger = logger;
        }

        // ── Read ────────────────────────────────────────────────────────────

        public List<PersonRecord> GetRecords() => _records.ToList();
        public int GetRecordCount() => _records.Count;

        // ── Import ──────────────────────────────────────────────────────────

        /// <summary>
        /// Parses an uploaded .xlsx/.xls stream using HEADER-NAME mapping.
        /// Columns can appear in any order — the parser finds them by their header text.
        /// Replaces the current in-memory dataset. Returns the count imported.
        /// </summary>
        public async Task<int> ImportFromExcelAsync(Stream stream)
        {
            var records = new List<PersonRecord>();

            try
            {
                using var package = new ExcelPackage(stream);
                var ws = package.Workbook.Worksheets[0];

                if (ws.Dimension == null || ws.Dimension.Rows < 2)
                {
                    _logger.LogWarning("Excel file has no data rows.");
                    return 0;
                }

                // Build a header → column-index map from row 1
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= ws.Dimension.Columns; c++)
                {
                    var header = ws.Cells[1, c].Value?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(header) && !colMap.ContainsKey(header))
                        colMap[header] = c;
                }

                _logger.LogInformation("Excel headers detected: {H}", string.Join(", ", colMap.Keys));

                int autoId = 1;
                for (int row = 2; row <= ws.Dimension.Rows; row++)
                {
                    // Skip rows where both Name and Surname are blank
                    var nameVal    = GetStr(ws, row, colMap, "Name");
                    var surnameVal = GetStr(ws, row, colMap, "Surname");
                    if (string.IsNullOrWhiteSpace(nameVal) && string.IsNullOrWhiteSpace(surnameVal))
                        continue;

                    try
                    {
                        var record = new PersonRecord
                        {
                            RowNumber         = autoId,
                            Name              = nameVal,
                            Surname           = surnameVal,
                            Identifier        = GetStr(ws, row, colMap, "Identifier"),
                            EmailAddress      = GetStr(ws, row, colMap, "EmailAddress", "Email Address", "Email"),
                            LocalMunicipality = GetStr(ws, row, colMap, "LocalMunicipality", "Local Municipality", "Municipality"),
                            HostCompany       = GetStr(ws, row, colMap, "HostCompany",  "Host Company",  "Host"),
                            LeadCompany       = GetStr(ws, row, colMap, "LeadCompany",  "Lead Company",  "Lead"),
                            JobType           = GetStr(ws, row, colMap, "JobType",       "Job Type",      "Job"),
                            DemographicGroup  = GetStr(ws, row, colMap, "DemographicGroup", "Demographic Group", "Race", "Demographic"),
                            Sex               = GetStr(ws, row, colMap, "Sex",    "Gender"),
                            ContactDetails    = GetStr(ws, row, colMap, "ContactDetails", "Contact Details", "Contact"),
                            EmploymentStatus  = GetStr(ws, row, colMap, "EmploymentStatus", "Employment Status", "Status"),
                            PersonDisability  = GetStr(ws, row, colMap, "PersonDisability", "Person Disability", "Disability", "PersonWithDisability")
                        };

                        records.Add(record);
                        autoId++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Failed to parse row {Row}. Skipping.", row);
                    }
                }

                _records = records;
                _nextId  = _records.Count > 0 ? _records.Max(r => r.RowNumber) + 1 : 1;
                _logger.LogInformation("Imported {Count} records.", _records.Count);
                NotifyChanged();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Fatal error reading Excel file.");
                throw;
            }

            return _records.Count;
        }

        // ── CRUD ────────────────────────────────────────────────────────────

        public void AddRecord(PersonRecord record)
        {
            record.RowNumber = _nextId++;
            _records.Add(record);
            NotifyChanged();
        }

        public bool UpdateRecord(PersonRecord updated)
        {
            var idx = _records.FindIndex(r => r.RowNumber == updated.RowNumber);
            if (idx < 0) return false;
            _records[idx] = updated;
            NotifyChanged();
            return true;
        }

        public bool DeleteRecord(int rowNumber)
        {
            var removed = _records.RemoveAll(r => r.RowNumber == rowNumber) > 0;
            if (removed) NotifyChanged();
            return removed;
        }

        public void ClearRecords()
        {
            _records.Clear();
            _nextId = 1;
            NotifyChanged();
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

        /// <summary>Returns the cell value for the first matching alias in colMap.</summary>
        private static string GetStr(ExcelWorksheet ws, int row,
            Dictionary<string, int> colMap, params string[] aliases)
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
