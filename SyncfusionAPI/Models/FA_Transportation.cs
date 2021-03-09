using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SyncfusionAPI.Models
{
    public partial class FA_Transportation
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
        [StringLength(15)]
        public string TravelingFrom { get; set; }
        [StringLength(15)]
        public string TravelingTo { get; set; }
        public byte? NumberOfPersons { get; set; }
        [StringLength(60)]
        public string NameOfTheFirm { get; set; }
        [StringLength(30)]
        public string ToTheAttentionOf { get; set; }
  
        public DateTime? TravelingOn { get; set; }
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
        [StringLength(70)]
        public string Travelers { get; set; }
        [StringLength(250)]
        public string SpecialComments { get; set; }
        public short? TravelType { get; set; }
        [StringLength(20)]
        public string RTravelingFrom { get; set; }
        [StringLength(20)]
        public string RTravelingTo { get; set; }

        public DateTime? RTravelingOn { get; set; }
    }
}
