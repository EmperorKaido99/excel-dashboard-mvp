using System.ComponentModel.DataAnnotations;

namespace ExcelDashboardMVP.Models
{
    /// <summary>
    /// Represents a person record with employment and demographic information.
    /// Columns match the uploaded Excel structure exactly.
    /// </summary>
    public class PersonRecord
    {
        [Key]
        public int RowNumber { get; set; }

        public string Name { get; set; } = string.Empty;
        public string Surname { get; set; } = string.Empty;
        public string Identifier { get; set; } = string.Empty;
        public string EmailAddress { get; set; } = string.Empty;
        public string LocalMunicipality { get; set; } = string.Empty;
        public string HostCompany { get; set; } = string.Empty;
        public string LeadCompany { get; set; } = string.Empty;
        public string JobType { get; set; } = string.Empty;
        public string DemographicGroup { get; set; } = string.Empty;
        public string Sex { get; set; } = string.Empty;
        public string ContactDetails { get; set; } = string.Empty;
        public string EmploymentStatus { get; set; } = string.Empty;
        public string PersonDisability { get; set; } = string.Empty;

        /// <summary>Full name computed from Name + Surname.</summary>
        public string FullName => $"{Name} {Surname}".Trim();

        /// <summary>
        /// Returns true when PersonDisability column indicates a disability.
        /// Accepts: Y, Yes, 1, True (case-insensitive).
        /// </summary>
        public bool HasDisability =>
            PersonDisability.Equals("Y",    StringComparison.OrdinalIgnoreCase) ||
            PersonDisability.Equals("Yes",  StringComparison.OrdinalIgnoreCase) ||
            PersonDisability.Equals("1",    StringComparison.OrdinalIgnoreCase) ||
            PersonDisability.Equals("True", StringComparison.OrdinalIgnoreCase);
    }
}
