using OfficeOpenXml;
using OfficeOpenXml.Style;
using ExcelDashboardMVP.Models;
using System.Drawing;

namespace ExcelDashboardMVP.Services
{
    /// <summary>
    /// Singleton service that stores PersonRecords in memory,
    /// handles Excel import/export, and notifies subscribers of changes.
    /// </summary>
    public class ExcelDataService
    {
        private readonly ILogger<ExcelDataService> _logger;
        private List<PersonRecord> _records = new();
        private int _nextId = 1;

        /// <summary>Fired whenever the record list changes (import, add, update, delete).</summary>
        public event Action? OnDataChanged;

        public ExcelDataService(ILogger<ExcelDataService> logger)
        {
            _logger = logger;
        }

        // ── Read ────────────────────────────────────────────────────────────

        /// <summary>Returns a copy of the current record list.</summary>
        public List<PersonRecord> GetRecords() => _records.ToList();

        /// <summary>Returns the total number of records in memory.</summary>
        public int GetRecordCount() => _records.Count;

        // ── Import ──────────────────────────────────────────────────────────

        /// <summary>
        /// Parses an uploaded .xlsx stream and replaces the in-memory list.
        /// Returns the number of records successfully imported.
        /// </summary>
        public async Task<int> ImportFromExcelAsync(Stream stream)
        {
            var records = new List<PersonRecord>();

            try
            {
                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension?.Rows ?? 0;

                if (rowCount < 2)
                {
                    _logger.LogWarning("Excel file has no data rows.");
                    return 0;
                }

                int autoId = 1;
                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        var record = new PersonRecord
                        {
                            Number                    = GetIntValue(worksheet, row, 1) > 0
                                                           ? GetIntValue(worksheet, row, 1)
                                                           : autoId,
                            Surname                   = GetStringValue(worksheet, row, 2),
                            Name                      = GetStringValue(worksheet, row, 3),
                            Identifier                = GetStringValue(worksheet, row, 4),
                            Age                       = GetIntValue(worksheet, row, 5),
                            Sex                       = GetStringValue(worksheet, row, 6),
                            PersonWithDisability      = GetBoolValue(worksheet, row, 7),
                            DemographicGroup          = GetStringValue(worksheet, row, 8),
                            ContactDetails            = GetStringValue(worksheet, row, 9),
                            AlternativeContactDetails = GetStringValue(worksheet, row, 10),
                            EmailAddress              = GetStringValue(worksheet, row, 11),
                            Address                   = GetStringValue(worksheet, row, 12),
                            Suburb                    = GetStringValue(worksheet, row, 13),
                            LocalMunicipality         = GetStringValue(worksheet, row, 14),
                            DistrictMunicipality      = GetStringValue(worksheet, row, 15),
                            EmploymentStatus          = GetStringValue(worksheet, row, 16),
                            StatusAtStartOfProgramme  = GetStringValue(worksheet, row, 17),
                            LeadCompany               = GetStringValue(worksheet, row, 18),
                            LeadCompanyAddress        = GetStringValue(worksheet, row, 19),
                            HostCompany               = GetStringValue(worksheet, row, 20),
                            JobType                   = GetStringValue(worksheet, row, 21),
                            StartDate                 = GetDateValue(worksheet, row, 22),
                            EndDate                   = GetDateValue(worksheet, row, 23),
                            PeriodOfPlacement         = GetIntValue(worksheet, row, 24),
                            DocumentPath              = GetStringValue(worksheet, row, 25)
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
                _nextId = _records.Count > 0 ? _records.Max(r => r.Number) + 1 : 1;

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

        /// <summary>Adds a new record and assigns the next available Number.</summary>
        public void AddRecord(PersonRecord record)
        {
            record.Number = _nextId++;
            _records.Add(record);
            NotifyChanged();
        }

        /// <summary>
        /// Replaces the record whose Number matches <paramref name="updated"/>.Number.
        /// </summary>
        public bool UpdateRecord(PersonRecord updated)
        {
            var idx = _records.FindIndex(r => r.Number == updated.Number);
            if (idx < 0) return false;
            _records[idx] = updated;
            NotifyChanged();
            return true;
        }

        /// <summary>Removes a record by its Number field.</summary>
        public bool DeleteRecord(int number)
        {
            var removed = _records.RemoveAll(r => r.Number == number) > 0;
            if (removed) NotifyChanged();
            return removed;
        }

        /// <summary>Clears all records from memory.</summary>
        public void ClearRecords()
        {
            _records.Clear();
            _nextId = 1;
            NotifyChanged();
            _logger.LogInformation("All records cleared.");
        }

        // ── Export ──────────────────────────────────────────────────────────

        /// <summary>
        /// Exports <paramref name="records"/> to an .xlsx byte array.
        /// Pass <c>GetRecords()</c> for all data or a filtered subset.
        /// </summary>
        public byte[] ExportToExcel(IEnumerable<PersonRecord> records)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Participants");

            // ── Headers ──────────────────────────────────────────────────
            string[] headers =
            {
                "No", "Surname", "Name", "Identifier", "Age", "Sex",
                "Person With Disability", "Demographic Group", "Contact Details",
                "Alternative Contact", "Email Address", "Address", "Suburb",
                "Local Municipality", "District Municipality", "Employment Status",
                "Status At Start", "Lead Company", "Lead Company Address",
                "Host Company", "Job Type", "Start Date", "End Date",
                "Period Of Placement", "Document Path"
            };

            for (int c = 0; c < headers.Length; c++)
            {
                var cell = ws.Cells[1, c + 1];
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x10, 0x74, 0xd6));
                cell.Style.Font.Color.SetColor(Color.White);
            }

            // ── Data rows ─────────────────────────────────────────────────
            int row = 2;
            foreach (var r in records)
            {
                ws.Cells[row, 1].Value  = r.Number;
                ws.Cells[row, 2].Value  = r.Surname;
                ws.Cells[row, 3].Value  = r.Name;
                ws.Cells[row, 4].Value  = r.Identifier;
                ws.Cells[row, 5].Value  = r.Age;
                ws.Cells[row, 6].Value  = r.Sex;
                ws.Cells[row, 7].Value  = r.PersonWithDisability ? "Y" : "N";
                ws.Cells[row, 8].Value  = r.DemographicGroup;
                ws.Cells[row, 9].Value  = r.ContactDetails;
                ws.Cells[row, 10].Value = r.AlternativeContactDetails;
                ws.Cells[row, 11].Value = r.EmailAddress;
                ws.Cells[row, 12].Value = r.Address;
                ws.Cells[row, 13].Value = r.Suburb;
                ws.Cells[row, 14].Value = r.LocalMunicipality;
                ws.Cells[row, 15].Value = r.DistrictMunicipality;
                ws.Cells[row, 16].Value = r.EmploymentStatus;
                ws.Cells[row, 17].Value = r.StatusAtStartOfProgramme;
                ws.Cells[row, 18].Value = r.LeadCompany;
                ws.Cells[row, 19].Value = r.LeadCompanyAddress;
                ws.Cells[row, 20].Value = r.HostCompany;
                ws.Cells[row, 21].Value = r.JobType;

                if (r.StartDate.HasValue)
                {
                    ws.Cells[row, 22].Value = r.StartDate.Value;
                    ws.Cells[row, 22].Style.Numberformat.Format = "yyyy-mm-dd";
                }
                if (r.EndDate.HasValue)
                {
                    ws.Cells[row, 23].Value = r.EndDate.Value;
                    ws.Cells[row, 23].Style.Numberformat.Format = "yyyy-mm-dd";
                }

                ws.Cells[row, 24].Value = r.PeriodOfPlacement;
                ws.Cells[row, 25].Value = r.DocumentPath;

                // Zebra-stripe rows
                if (row % 2 == 0)
                {
                    using var range = ws.Cells[row, 1, row, 25];
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xF0, 0xF4, 0xFF));
                }

                row++;
            }

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            return package.GetAsByteArray();
        }

        // ── Private helpers ─────────────────────────────────────────────────

        private void NotifyChanged() => OnDataChanged?.Invoke();

        private static string GetStringValue(ExcelWorksheet ws, int row, int col)
            => ws.Cells[row, col].Value?.ToString()?.Trim() ?? string.Empty;

        private static int GetIntValue(ExcelWorksheet ws, int row, int col)
        {
            var v = ws.Cells[row, col].Value;
            if (v is null) return 0;
            if (int.TryParse(v.ToString(), out int i)) return i;
            if (double.TryParse(v.ToString(), out double d)) return (int)d;
            return 0;
        }

        private static bool GetBoolValue(ExcelWorksheet ws, int row, int col)
        {
            var v = ws.Cells[row, col].Value?.ToString()?.Trim().ToUpperInvariant() ?? "";
            return v is "Y" or "YES" or "TRUE" or "1";
        }

        private static DateTime? GetDateValue(ExcelWorksheet ws, int row, int col)
        {
            var v = ws.Cells[row, col].Value;
            if (v is null) return null;
            if (v is DateTime dt) return dt;
            if (DateTime.TryParse(v.ToString(), out DateTime p)) return p;
            if (double.TryParse(v.ToString(), out double s))
            {
                try { return DateTime.FromOADate(s); } catch { return null; }
            }
            return null;
        }
    }
}