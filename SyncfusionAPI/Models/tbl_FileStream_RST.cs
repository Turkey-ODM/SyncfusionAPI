using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace SyncfusionAPI.Models
{
    public class tbl_FileStream_RST
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string CaseNumber { get; set; }
        [Key]
        public int ID_Case { get; set; }
        public string Description { get; set; }
        public DateTime? Date { get; set; }
        public string TagName { get; set; }
        public string IndividualID { get; set; }
        public string IndividualName { get; set; }
        public string User_Name { get; set; }
        public Boolean Active { get; set; }
        public string UpdateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public int TagNameValue { get; set; }
        public Guid uniqueID { get; set; }
        public int IDD { get; set; }
       
    }
}
