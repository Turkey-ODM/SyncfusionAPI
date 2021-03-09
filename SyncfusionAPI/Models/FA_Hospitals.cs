using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SyncfusionAPI.Models
{
    public partial class FA_Hospitals
    {


        [Key]
        public int RequestID { get; set; }
        [Required]
        [StringLength(12)]
        public string CaseNumber { get; set; }
        [Required]
        [StringLength(12)]
        public string IndividualID { get; set; }
        [Required]
        [StringLength(8)]
        public string RequestType { get; set; }
        [StringLength(3)]
        public string Nationality { get; set; }
        [StringLength(1)]
        public string Sex { get; set; }
        [StringLength(30)]
        public string GivenName { get; set; }
        [StringLength(50)]
        public string FamilyName { get; set; }
        [StringLength(15)]
        public string City { get; set; }
        [StringLength(256)]
        public string Hospital { get; set; }
        public short? Amount { get; set; }
        [StringLength(50)]
        public string AmountInWritten { get; set; }
        [StringLength(12)]
        public string RequestedBy { get; set; }

        public DateTime RequestedOn { get; set; }
        [StringLength(250)]
        public string RequestComments { get; set; }
        [StringLength(12)]
        public string DecidedBy { get; set; }

        public DateTime DecidedOn { get; set; }
        [StringLength(1)]
        public string RequestStatus { get; set; }
        [StringLength(250)]
        public string DecisionComments { get; set; }

        public byte FamilySize { get; set; }

        public short? AmountRequested { get; set; }
        [StringLength(250)]
        public string SpecialComments { get; set; }



        //[Key]
        //public int RequestID { get; set; }
        //[Required]
        //[StringLength(12)]
        //public string CaseNumber { get; set; }
        //[Required]
        //[StringLength(12)]
        //public string IndividualID { get; set; }
        //[Required]
        //[StringLength(8)]
        //public string RequestType { get; set; }
        //[StringLength(3)]
        //public string Nationality { get; set; }
        //[StringLength(1)]
        //public string Sex { get; set; }
        //[StringLength(30)]
        //public string GivenName { get; set; }
        //[StringLength(50)]
        //public string FamilyName { get; set; }
        //[StringLength(15)]
        //public string City { get; set; }
        //[StringLength(250)]
        //public string Hospital { get; set; }
        //public short? Amount { get; set; }
        //[StringLength(50)]
        //public string AmountInWritten { get; set; }
        //[StringLength(12)]
        //public string RequestedBy { get; set; }
        //[Column(TypeName = "datetime")]
        //public DateTime? RequestedOn { get; set; }
        //[StringLength(250)]
        //public string RequestComments { get; set; }
        //[StringLength(12)]
        //public string DecidedBy { get; set; }
        //[Column(TypeName = "datetime")]
        //public DateTime? DecidedOn { get; set; }
        //[Required]
        //[StringLength(1)]
        //public string RequestStatus { get; set; }
        //[StringLength(250)]
        //public string DecisionComments { get; set; }
        //public byte? FamilySize { get; set; }
        //public short? AmountRequested { get; set; }
        //[StringLength(250)]
        //public string SpecialComments { get; set; }
    }
}
