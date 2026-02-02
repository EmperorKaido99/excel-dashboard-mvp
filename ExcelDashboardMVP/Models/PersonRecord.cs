using System.ComponentModel.DataAnnotations;

namespace ExcelDashboardMVP.Models
{
    /// <summary>
    /// Represents a person record with employment and demographic information
    /// </summary>
    public class PersonRecord
    {
        [Key]
        public int Number { get; set; }

        [Required]
        [StringLength(100)]
        public string Surname { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;

        [StringLength(50)]
        public string Identifier { get; set; } = string.Empty;

        public int Age { get; set; }

        [StringLength(10)]
        public string Sex { get; set; } = string.Empty;

        public bool PersonWithDisability { get; set; }

        [StringLength(100)]
        public string DemographicGroup { get; set; } = string.Empty;

        [StringLength(200)]
        public string ContactDetails { get; set; } = string.Empty;

        [StringLength(200)]
        public string AlternativeContactDetails { get; set; } = string.Empty;

        [EmailAddress]
        [StringLength(200)]
        public string EmailAddress { get; set; } = string.Empty;

        [StringLength(300)]
        public string Address { get; set; } = string.Empty;

        [StringLength(100)]
        public string Suburb { get; set; } = string.Empty;

        [StringLength(100)]
        public string LocalMunicipality { get; set; } = string.Empty;

        [StringLength(100)]
        public string DistrictMunicipality { get; set; } = string.Empty;

        [StringLength(100)]
        public string EmploymentStatus { get; set; } = string.Empty;

        [StringLength(100)]
        public string StatusAtStartOfProgramme { get; set; } = string.Empty;

        [StringLength(200)]
        public string LeadCompany { get; set; } = string.Empty;

        [StringLength(300)]
        public string LeadCompanyAddress { get; set; } = string.Empty;

        [StringLength(200)]
        public string HostCompany { get; set; } = string.Empty;

        [StringLength(100)]
        public string JobType { get; set; } = string.Empty;

        public DateTime? StartDate { get; set; }

        public DateTime? EndDate { get; set; }

        public int PeriodOfPlacement { get; set; }

        [StringLength(500)]
        public string DocumentPath { get; set; } = string.Empty;

        /// <summary>
        /// Calculated property for full name
        /// </summary>
        public string FullName => $"{Name} {Surname}";

        /// <summary>
        /// Calculated property to check if placement is currently active
        /// </summary>
        public bool IsActivePlacement
        {
            get
            {
                if (!StartDate.HasValue || !EndDate.HasValue)
                    return false;

                var today = DateTime.Today;
                return today >= StartDate.Value && today <= EndDate.Value;
            }
        }
    }
}