using OfficeOpenXml;
using ExcelDashboardMVP.Models;
using System.Globalization;

namespace ExcelDashboardMVP.Services
{
    /// <summary>
    /// Service for importing PersonRecord data from Excel files
    /// </summary>
    public class ExcelImportService
    {
        private readonly ILogger<ExcelImportService> _logger;

        public ExcelImportService(ILogger<ExcelImportService> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Imports person records from an Excel file
        /// </summary>
        /// <param name="stream">Stream containing the Excel file data</param>
        /// <returns>List of PersonRecord objects</returns>
        public async Task<List<PersonRecord>> ImportFromExcelAsync(Stream stream)
        {
            var records = new List<PersonRecord>();

            try
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Get first worksheet
                    var rowCount = worksheet.Dimension?.Rows ?? 0;

                    if (rowCount < 2)
                    {
                        _logger.LogWarning("Excel file is empty or has no data rows");
                        return records;
                    }

                    // Start from row 2 (assuming row 1 is headers)
                    for (int row = 2; row <= rowCount; row++)
                    {
                        try
                        {
                            var record = new PersonRecord
                            {
                                Number = GetIntValue(worksheet, row, 1),
                                Surname = GetStringValue(worksheet, row, 2),
                                Name = GetStringValue(worksheet, row, 3),
                                Identifier = GetStringValue(worksheet, row, 4),
                                Age = GetIntValue(worksheet, row, 5),
                                Sex = GetStringValue(worksheet, row, 6),
                                PersonWithDisability = GetBoolValue(worksheet, row, 7),
                                DemographicGroup = GetStringValue(worksheet, row, 8),
                                ContactDetails = GetStringValue(worksheet, row, 9),
                                AlternativeContactDetails = GetStringValue(worksheet, row, 10),
                                EmailAddress = GetStringValue(worksheet, row, 11),
                                Address = GetStringValue(worksheet, row, 12),
                                Suburb = GetStringValue(worksheet, row, 13),
                                LocalMunicipality = GetStringValue(worksheet, row, 14),
                                DistrictMunicipality = GetStringValue(worksheet, row, 15),
                                EmploymentStatus = GetStringValue(worksheet, row, 16),
                                StatusAtStartOfProgramme = GetStringValue(worksheet, row, 17),
                                LeadCompany = GetStringValue(worksheet, row, 18),
                                LeadCompanyAddress = GetStringValue(worksheet, row, 19),
                                HostCompany = GetStringValue(worksheet, row, 20),
                                JobType = GetStringValue(worksheet, row, 21),
                                StartDate = GetDateValue(worksheet, row, 22),
                                EndDate = GetDateValue(worksheet, row, 23),
                                PeriodOfPlacement = GetIntValue(worksheet, row, 24),
                                DocumentPath = GetStringValue(worksheet, row, 25)
                            };

                            records.Add(record);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"Error parsing row {row}");
                            // Continue processing other rows
                        }
                    }
                }

                _logger.LogInformation($"Successfully imported {records.Count} records from Excel");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing Excel file");
                throw;
            }

            return records;
        }

        /// <summary>
        /// Imports person records from an Excel file path
        /// </summary>
        public async Task<List<PersonRecord>> ImportFromExcelFileAsync(string filePath)
        {
            using (var stream = File.OpenRead(filePath))
            {
                return await ImportFromExcelAsync(stream);
            }
        }

        #region Helper Methods

        private string GetStringValue(ExcelWorksheet worksheet, int row, int col)
        {
            var value = worksheet.Cells[row, col].Value;
            return value?.ToString()?.Trim() ?? string.Empty;
        }

        private int GetIntValue(ExcelWorksheet worksheet, int row, int col)
        {
            var value = worksheet.Cells[row, col].Value;
            
            if (value == null)
                return 0;

            if (int.TryParse(value.ToString(), out int result))
                return result;

            if (double.TryParse(value.ToString(), out double doubleResult))
                return (int)doubleResult;

            return 0;
        }

        private bool GetBoolValue(ExcelWorksheet worksheet, int row, int col)
        {
            var value = GetStringValue(worksheet, row, col).ToUpper();
            
            // Check for Y/N or Yes/No
            if (value == "Y" || value == "YES" || value == "TRUE" || value == "1")
                return true;

            if (value == "N" || value == "NO" || value == "FALSE" || value == "0")
                return false;

            return false;
        }

        private DateTime? GetDateValue(ExcelWorksheet worksheet, int row, int col)
        {
            var value = worksheet.Cells[row, col].Value;

            if (value == null)
                return null;

            // If it's already a DateTime
            if (value is DateTime dateTime)
                return dateTime;

            // Try to parse string
            if (DateTime.TryParse(value.ToString(), out DateTime result))
                return result;

            // Try to parse Excel serial date
            if (double.TryParse(value.ToString(), out double serialDate))
            {
                try
                {
                    return DateTime.FromOADate(serialDate);
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        #endregion
    }
}