using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace WEB_PHANTICHLOI.Models
{
    public class ErrorAnalysisDbContext : DbContext
    {
        public ErrorAnalysisDbContext() : base("DB_PHANTICHLOI")
        {
        }

        public DbSet<ErrorAnalysis> ErrorAnalyses { get; set; }
        public DbSet<ErrorAnalysisLookup> ErrorAnalysisLookups { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ErrorAnalysis>().ToTable("ErrorAnalyses");
            modelBuilder.Entity<ErrorAnalysisLookup>().ToTable("ErrorAnalysisLookups");

            base.OnModelCreating(modelBuilder);
        }
    }
}