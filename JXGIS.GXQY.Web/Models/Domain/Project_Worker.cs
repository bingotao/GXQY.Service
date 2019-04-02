using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("Project_Worker")]
    public class Project_Worker
    {
        [Key, Column(Order = 1)]
        public string ProjectId { get; set; }

        [Key, Column(Order = 2)]
        public string WorkerId { get; set; }

        [NotMapped]
        public string Name { get; set; }

        public int? Index { get; set; }

        public string ProjectRole { get; set; }
    }
}