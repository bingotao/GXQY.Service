using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("Project")]
    public class Project : BaseEntity
    {
        [Key]
        public string Id { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }

        public double? ContractAmount { get; set; }

        public DateTime? StartTime { get; set; }

        public DateTime? EndTime { get; set; }

        public string SerialNumber { get; set; }

        public string Year { get; set; }

        public string ResolutionNumber { get; set; }

        public DateTime? CheckTime { get; set; }

        public string Department { get; set; }


        public List<Worker> Workers { get; set; }


        public bool Validate(PCDbContext db, out string msg)
        {
            StringBuilder sb = new StringBuilder();
            if (string.IsNullOrEmpty(this.Name))
            {
                sb.Append("项目名称不能为空！\n");
            }

            //if (this.ContractAmount == null)
            //{
            //    sb.Append("项目金额不能为空！\n");
            //}

            //if (StartTime == null || EndTime == null)
            //{
            //    sb.Append("项目起始时间不能为空！\n");
            //}

            var cnt = db.Project.Where(p => p.Id != this.Id && p.Name == this.Name).Count();
            if (cnt > 0)
            {
                sb.Append("已存在该项目名！\n");
            }
            msg = sb.ToString();

            return string.IsNullOrEmpty(msg);
        }
    }
}