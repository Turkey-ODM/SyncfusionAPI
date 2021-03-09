using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SyncfusionAPI.Models
{
    public partial class _DepDetailsTable
    {
        [Key]
        public string IndividualId { get; set; }
        public string Rel { get; set; }
        public string GivenName { get; set; }
        public string FamilyName { get; set; }
        public string Nationality { get; set; }
        public string Religion { get; set; }
        public string Ethnicity { get; set; }
        public string RegDate { get; set; }
        public string BirthCity { get; set; }
        public string Sex { get; set; }
        public short? Age { get; set; }
        public byte[] Photo { get; set; }

        public string IndividualGUID { get; set; }

        public string ProcessingGroupGUID { get; set; }
        public string HasRST107 { get; set; }
        public string HasREG50 { get; set; }
        public string LanguageText { get; set; }
        public string PreAssessmentLanguage { get; set; }

    }
}
