
using Microsoft.EntityFrameworkCore;
using SyncfusionAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SyncfusionAPI.Data
{
    public class ApplicationDbContext : DbContext
    {
      
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }


        public DbSet<FA_Hospitals> FA_Hospitals { get; set; }
        public DbSet<printdetails> printdetails { get; set; }
        public DbSet<FA_Accomodation> FA_Accomodation { get; set; }
        public DbSet<FA_Transportation> FA_Transportation { get; set; }
        public DbSet<RSTCaseBioData> RSTCaseBioData { get; set; }
        public DbSet<_DepDetailsTable> _DepDetailsTable { get; set; }

        


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<printdetails>().HasNoKey().ToView(null);
            modelBuilder.Entity<RSTCaseBioData>().HasNoKey().ToView(null);
            modelBuilder.Entity<_DepDetailsTable>().HasNoKey().ToView(null);


        }

    }
}
