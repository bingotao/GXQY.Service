using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Web;

namespace JXGIS.GXQY.Web.Models
{
    [Table("Worker")]
    public class Worker : BaseEntity
    {
        [Key]
        public string Id { get; set; }

        public string Name { get; set; }

        public string Department { get; set; }

        public List<Project> Projects { get; set; }

        internal bool Validate(PCDbContext db, out string msg)
        {
            StringBuilder sb = new StringBuilder();
            if (string.IsNullOrEmpty(this.Name))
            {
                sb.Append("姓名不能为空！\n");
            }

            var cnt = db.Worker.Where(p => p.Id != this.Id && p.Name == this.Name && p.IsValid == 1).Count();
            if (cnt > 0)
            {
                sb.Append("已存在该姓名人员！\n");
            }
            msg = sb.ToString();

            return string.IsNullOrEmpty(msg);
        }
    }
}