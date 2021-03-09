using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace SyncfusionAPI.Models
{
    public class tbl_FileStream_Data
    {
        [Key]
        public int ID_Case { get; set; }

        //public Guid Id { get; set; }

        public byte[] Data { get; set; }

        public string CaseNumber { get; set; }
       

        public DateTime Date { get; set; }

        public string Name { get; set; }
    }
}
