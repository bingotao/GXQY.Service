using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("Department")]
    public class Department : BaseEntity
    {
        [Key]
        public string Id { get; set; }

        public string P_Id { get; set; }

        public string Name { get; set; }

        [NotMapped]
        public int Level { get; set; }
        //[ForeignKey("P_Id")]
        public Department PDepartment { get; set; }

        //[ForeignKey("P_Id")]
        public List<Department> SubDepartments { get; set; }
    }
}