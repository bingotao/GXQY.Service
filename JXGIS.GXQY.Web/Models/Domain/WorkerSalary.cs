using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("WorkerSalary")]
    public class WorkerSalary : BaseEntity
    {
        [Key]
        public string Id { get; set; }

        public string WorkerId { get; set; }

        [NotMapped]
        public string WorkerName { get; set; }


        public int? Month { get; set; }

        public double BasePay { get; set; }

        public double Bonus { get; set; }

        public double AccumulationFund { get; set; }

        public double SocialSecurity { get; set; }

        public double SS1 { get; set; }

        public double SS2 { get; set; }

        public double SS3 { get; set; }

        public double SS4 { get; set; }

        public double SS5 { get; set; }

        public double Bonus1 { get; set; }

        public int MonthX { get; set; }
    }
}