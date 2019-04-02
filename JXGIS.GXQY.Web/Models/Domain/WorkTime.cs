using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("WorkTime")]
    public class WorkTime : BaseEntity
    {
        [Key]
        public string Id { get; set; }

        public string WorkerId { get; set; }

        public Worker Worker { get; set; }

        public string ProjectId { get; set; }

        public Project Project { get; set; }

        public int? Date { get; set; }

        public int? Month { get; set; }

        public string WorkType { get; set; }
    }
}