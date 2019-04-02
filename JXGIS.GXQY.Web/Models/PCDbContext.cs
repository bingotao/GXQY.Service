using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace JXGIS.GXQY.Web.Models
{
    public class PCDbContext : DbContext
    {
        private static string conStr = System.Configuration.ConfigurationManager.ConnectionStrings["PCDbContext"].ToString();
        public static PCDbContext NewDbContext
        {
            get
            {
                return new PCDbContext();
            }
        }

        public PCDbContext() : base(conStr)
        {
            this.Database.Initialize(false);
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.HasDefaultSchema("dbo");
            modelBuilder.Entity<Project>().HasMany(b => b.Workers).WithMany(c => c.Projects).Map(m =>
            {
                m.MapLeftKey("ProjectId");
                m.MapRightKey("WorkerId");
                m.ToTable("Project_Worker");
            });

            //modelBuilder.Entity<Department>().HasMany(d => d.SubDepartments).WithOptional(d => d.PDepartment).Map(m =>
            //{
            //    m.MapKey("P_ID");
            //    m.ToTable("Department");
            //});
        }


        public DbSet<Department> Department { get; set; }

        public DbSet<Project> Project { get; set; }

        public DbSet<Worker> Worker { get; set; }

        public DbSet<WorkerSalary> WorkerSalary { get; set; }

        public DbSet<WorkTime> WorkTime { get; set; }

        public DbSet<DateType> DateType { get; set; }

        //public DbSet<Project_Worker> Project_Worker { get; set; }


    }
}