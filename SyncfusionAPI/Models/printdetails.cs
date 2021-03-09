using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SyncfusionAPI.Models
{
    public class printdetails
    {
        public string ProcessingGroupNumber { get; set; }

        public string rel { get; set; }
        public string GivenName { get; set; }
        public string OriginCountryCode { get; set; }
        public DateTime DateofBirth { get; set; }

        public string BirthCityTownVillage { get; set; }

        public byte[] Photo { get; set; }

       
        public string IndividualID { get; set; }



        public string SexCode { get; set; }

        public string FamilyName { get; set; }


        public Int16 ProcessingGroupSize { get; set; }

        public string LocationLevel1Description { get; set; }
    }
}
