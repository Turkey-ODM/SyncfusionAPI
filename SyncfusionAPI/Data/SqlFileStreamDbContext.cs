using Microsoft.EntityFrameworkCore;
using SyncfusionAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SyncfusionAPI.Data
{
    public class SqlFileStreamDbContext :DbContext
    {
        public SqlFileStreamDbContext(DbContextOptions<SqlFileStreamDbContext> options)
          : base(options)
        {
        }

        public DbSet<tbl_FileStream_RST> tbl_FileStream_RST { get; set; }
        public DbSet<tbl_FileStream_Data> tbl_FileStream_Data { get; set; }
        //public DbSet<EDRMS_ScannedGrid> EDRMS_ScannedGrid { get; set; }
        //public DbSet<EDRMS_AssessmentsGrid> EDRMS_AssessmentsGrid { get; set; }
        //public DbSet<tblDropDownList> tblDropDownList { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
           // modelBuilder.Entity<EDRMS_ScannedGrid>().HasNoKey().ToView(null);
           // modelBuilder.Entity<EDRMS_AssessmentsGrid>().HasNoKey().ToView(null);
        }

    }
}
