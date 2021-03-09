using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SyncfusionAPI.Models
{
    public partial class FA_Accomodation
    {
        [Key]
        public int RequestID { get; set; }
        [StringLength(12)]
        public string CaseNumber { get; set; }
        [StringLength(12)]
        public string IndividualID { get; set; }
        [StringLength(3)]
        public string Nationality { get; set; }
        [StringLength(30)]
        public string GivenName { get; set; }
        [StringLength(50)]
        public string FamilyName { get; set; }
       
        public DateTime? StartingDate { get; set; }
    
        public DateTime? EndingDate { get; set; }
        public byte? NumberOfPersons { get; set; }
        [StringLength(30)]
        public string HotelName { get; set; }
        [StringLength(70)]
        public string HotelAddress { get; set; }
        [StringLength(250)]
        public string RequestComments { get; set; }
 
        public DateTime? RequestedOn { get; set; }
        [StringLength(12)]
        public string RequestedBy { get; set; }
        [StringLength(12)]
        public string DecidedBy { get; set; }

        public DateTime? DecidedOn { get; set; }
        [Required]
        [StringLength(1)]
        public string RequestStatus { get; set; }
        [StringLength(250)]
        public string DecisionComments { get; set; }
        public byte? TotalFemale { get; set; }
        public byte? TotalMale { get; set; }
        [Required]
        [StringLength(30)]
        public string AccomodationType { get; set; }
        public byte? TotalDays { get; set; }
        [StringLength(250)]
        public string SpecialComments { get; set; }
    }
}
