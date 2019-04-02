using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("DateType")]
    public class DateType
    {
        [Key]
        public string Id { get; set; }

        public int Date { get; set; }

        public int Month { get; set; }

        public string Type { get; set; }
    }
}