using OfficeOpenXml;
using ExcelDashboardMVP.Models;

namespace ExcelDashboardMVP.Services
{
    /// <summary>
    /// Legacy import service — kept for backward compatibility with Import.razor.
    /// For the main data pipeline, use ExcelDataService which supports header-name mapping.
    /// This service uses position-based parsing aligned to the new 13-column structure.
    /// </summary>
    public class ExcelImportService
    {
        private readonly ILogger<ExcelImportService> _logger;

        public ExcelImportService(ILogger<ExcelImportService> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Imports person records from an Excel stream using header-name detection.
        /// Falls back to column-position mapping if headers are not recognised.
        /// </summary>
        public async Task<List<PersonRecord>> ImportFromExcelAsync(Stream stream)
        {
            var records = new List<PersonRecord>();

            try
            {
                using var package = new ExcelPackage(stream);
                var ws = package.Workbook.Worksheets[0];

                if (ws.Dimension == null || ws.Dimension.Rows < 2)
                {
                    _logger.LogWarning("Excel file has no data rows.");
                    return records;
                }

                // Build header → column-index map (case-insensitive)
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= ws.Dimension.Columns; c++)
                {
                    var h = ws.Cells[1, c].Value?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(h) && !colMap.ContainsKey(h))
                        colMap[h] = c;
                }

                int id = 1;
                for (int row = 2; row <= ws.Dimension.Rows; row++)
                {
                    try
                    {
                        var name    = GetStr(ws, row, colMap, "Name");
                        var surname = GetStr(ws, row, colMap, "Surname");
                        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(surname))
                            continue;

                        records.Add(new PersonRecord
                        {
                            RowNumber         = id++,
                            Name              = name,
                            Surname           = surname,
                            Identifier        = GetStr(ws, row, colMap, "Identifier"),
                            EmailAddress      = GetStr(ws, row, colMap, "EmailAddress", "Email Address", "Email"),
                            LocalMunicipality = GetStr(ws, row, colMap, "LocalMunicipality", "Local Municipality", "Municipality"),
                            HostCompany       = GetStr(ws, row, colMap, "HostCompany", "Host Company", "Host"),
                            LeadCompany       = GetStr(ws, row, colMap, "LeadCompany", "Lead Company", "Lead"),
                            JobType           = GetStr(ws, row, colMap, "JobType", "Job Type", "Job"),
                            DemographicGroup  = GetStr(ws, row, colMap, "DemographicGroup", "Demographic Group", "Race", "Demographic"),
                            Sex               = GetStr(ws, row, colMap, "Sex", "Gender"),
                            ContactDetails    = GetStr(ws, row, colMap, "ContactDetails", "Contact Details", "Contact"),
                            EmploymentStatus  = GetStr(ws, row, colMap, "EmploymentStatus", "Employment Status", "Status"),
                            PersonDisability  = GetStr(ws, row, colMap, "PersonDisability", "Person Disability", "Disability")
                        });
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error parsing row {Row}. Skipping.", row);
                    }
                }

                _logger.LogInformation("Imported {Count} records.", records.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading Excel file.");
                throw;
            }

            return records;
        }

        public async Task<List<PersonRecord>> ImportFromExcelFileAsync(string filePath)
        {
            using var stream = File.OpenRead(filePath);
            return await ImportFromExcelAsync(stream);
        }

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