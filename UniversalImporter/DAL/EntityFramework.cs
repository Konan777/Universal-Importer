using System.Data.Common;
using System.Data.Entity;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Configuration;
using System.Data.SqlClient;


namespace UniversalImporter.DAL
{
    public class MSSQL_Database 
    {
        public TMContext GetNewContext()
        {
            return new TMContext(new SqlConnection(ConfigurationSettings.AppSettings["ConnectionString"].ToString()), false);
        }
    }
    public class TMContext : DbContext
    {
        public TMContext(DbConnection dbConnection, bool contextOwnsConnection) : base(dbConnection, contextOwnsConnection)
        {
            Configuration.ProxyCreationEnabled = false;
        }

        public DbSet<AAA_300> AAA_300 { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //System.Data.Entity.Database.SetInitializer<TMContext>(null);
            base.OnModelCreating(modelBuilder);
            // Map entity to table
            modelBuilder.Entity<AAA_300>().ToTable("AAA_300");
        }

    }
    public class AAA_300
    {
        [Key]

        public int Id { get; set; }
        public DateTime? DT01 { get; set; }
        public string TXT01 { get; set; }
        public string TXT02 { get; set; }
        public string TXT03 { get; set; }
        public string TIME01 { get; set; }
        public string TIME02 { get; set; }
        public double? COST { get; set; }
    }
}
