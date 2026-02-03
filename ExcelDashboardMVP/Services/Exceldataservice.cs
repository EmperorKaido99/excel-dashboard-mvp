using OfficeOpenXml;
using ExcelDashboardMVP.Models;

namespace ExcelDashboardMVP.Services
{
    /// <summary>
    /// Singleton service that handles Excel upload parsing and stores
    /// PersonRecords in memory for access across the application.
    /// </summary>
    public class ExcelDataService
    {
        private readonly ILogger<ExcelDataService> _logger;
        private List<PersonRecord> _records = new();

        public ExcelDataService(ILogger<ExcelDataService> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Returns the currently stored list of PersonRecords.
        /// </summary>
        public List<PersonRecord> GetRecords()
        {
            return _records;
        }

        /// <summary>
        /// Returns the total number of records currently in memory.
        /// </summary>
        public int GetRecordCount()
        {
            return _records.Count;
        }

        /// <summary>
        /// Clears all records from memory.
        /// </summary>
        public void ClearRecords()
        {
            _records.Clear();
            _logger.LogInformation("All records cleared from memory.");
        }

        /// <summary>
        /// Accepts a stream from an uploaded .xlsx file, parses every row
        /// into a PersonRecord, and replaces the in-memory list.
        /// </summary>
        /// <param name="stream">The uploaded Excel file stream.</param>
        /// <returns>The number of records successfully imported.</returns>
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
                    _logger.LogWarning("Excel file has no data rows (only a header or is empty).");
                    return 0;
                }

                // Row 1 is the header — data starts at row 2
                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        var record = new PersonRecord
                        {
                            Number                   = GetIntValue(worksheet, row, 1),
                            Surname                  = GetStringValue(worksheet, row, 2),
                            Name                     = GetStringValue(worksheet, row, 3),
                            Identifier               = GetStringValue(worksheet, row, 4),
                            Age                      = GetIntValue(worksheet, row, 5),
                            Sex                      = GetStringValue(worksheet, row, 6),
                            PersonWithDisability     = GetBoolValue(worksheet, row, 7),
                            DemographicGroup         = GetStringValue(worksheet, row, 8),
                            ContactDetails           = GetStringValue(worksheet, row, 9),
                            AlternativeContactDetails = GetStringValue(worksheet, row, 10),
                            EmailAddress             = GetStringValue(worksheet, row, 11),
                            Address                  = GetStringValue(worksheet, row, 12),
                            Suburb                   = GetStringValue(worksheet, row, 13),
                            LocalMunicipality        = GetStringValue(worksheet, row, 14),
                            DistrictMunicipality     = GetStringValue(worksheet, row, 15),
                            EmploymentStatus         = GetStringValue(worksheet, row, 16),
                            StatusAtStartOfProgramme = GetStringValue(worksheet, row, 17),
                            LeadCompany              = GetStringValue(worksheet, row, 18),
                            LeadCompanyAddress       = GetStringValue(worksheet, row, 19),
                            HostCompany              = GetStringValue(worksheet, row, 20),
                            JobType                  = GetStringValue(worksheet, row, 21),
                            StartDate                = GetDateValue(worksheet, row, 22),
                            EndDate                  = GetDateValue(worksheet, row, 23),
                            PeriodOfPlacement        = GetIntValue(worksheet, row, 24),
                            DocumentPath             = GetStringValue(worksheet, row, 25)
                        };

                        records.Add(record);
                    }
                    catch (Exception ex)
                    {
                        // Log the bad row but keep going so one bad row
                        // does not kill the entire import.
                        _logger.LogError(ex, "Failed to parse row {Row}. Skipping.", row);
                    }
                }

                // Only replace the live list after the whole file has been
                // parsed successfully — avoids a half-imported state.
                _records = records;

                _logger.LogInformation("Imported {Count} records from Excel.", _records.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Fatal error while reading Excel file.");
                throw;
            }

            return _records.Count;
        }

        #region Helper Methods

        private static string GetStringValue(ExcelWorksheet ws, int row, int col)
        {
            return ws.Cells[row, col].Value?.ToString()?.Trim() ?? string.Empty;
        }

        private static int GetIntValue(ExcelWorksheet ws, int row, int col)
        {
            var value = ws.Cells[row, col].Value;

            if (value == null) return 0;

            if (int.TryParse(value.ToString(), out int intResult))
                return intResult;

            if (double.TryParse(value.ToString(), out double dblResult))
                return (int)dblResult;

            return 0;
        }

        private static bool GetBoolValue(ExcelWorksheet ws, int row, int col)
        {
            var value = ws.Cells[row, col].Value?.ToString()?.Trim().ToUpperInvariant() ?? "";

            return value is "Y" or "YES" or "TRUE" or "1";
        }

        private static DateTime? GetDateValue(ExcelWorksheet ws, int row, int col)
        {
            var value = ws.Cells[row, col].Value;

            if (value == null) return null;

            if (value is DateTime dt) return dt;

            if (DateTime.TryParse(value.ToString(), out DateTime parsed))
                return parsed;

            // Excel sometimes stores dates as serial numbers
            if (double.TryParse(value.ToString(), out double serial))
            {
                try { return DateTime.FromOADate(serial); }
                catch { return null; }
            }

            return null;
        }

        #endregion
    }
}
