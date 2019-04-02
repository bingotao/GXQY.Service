using JXGIS.GXQY.Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json.Linq;
using System.Reflection;
using JXGIS.GXQY.Web.Base;
using System.Transactions;
using System.Data.SqlClient;

namespace JXGIS.GXQY.Web.Controllers
{
    public class ProjectController : Controller
    {
        public ContentResult GetProjects(string prjState, int month)
        {
            string s = null;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    List<Project> projects = null;
                    var query = db.Project;
                    var selectedMonth = DateTime.Parse($"{(int)(month / 100)}-{month % 100}-01");

                    if (prjState == "已结束")
                    {
                        projects = query.Where(p => p.EndTime != null && p.EndTime < selectedMonth).OrderBy(p => p.Year).OrderBy(p => p.SerialNumber).ToList();
                    }
                    else if (prjState == "进行中")
                    {
                        var startTime = selectedMonth.AddMonths(1);
                        projects = query.Where(p => (p.EndTime == null || p.EndTime >= selectedMonth) && p.StartTime < startTime).OrderBy(p => p.Year).OrderBy(p => p.SerialNumber).ToList();
                    }
                    else
                    {
                        projects = query.OrderBy(p => p.Year).OrderBy(p => p.SerialNumber).ToList();
                    }

                    RtObj.Serialize(projects, out s, new Newtonsoft.Json.Converters.IsoDateTimeConverter() { DateTimeFormat = "yyyy年MM月dd日" });
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }

        public ActionResult ProjectForm()
        {
            return PartialView();
        }


        public static PropertyInfo[] properties = typeof(Project).GetProperties();

        public ActionResult ModifyProject(string project)
        {
            string s = null;
            try
            {
                var nPrj = Newtonsoft.Json.JsonConvert.DeserializeObject<Project>(project);
                if (nPrj != null)
                {
                    using (var db = PCDbContext.NewDbContext)
                    {
                        Project rp = nPrj;
                        // 新增
                        if (string.IsNullOrEmpty(nPrj.Id))
                        {
                            nPrj.Id = Guid.NewGuid().ToString();
                            nPrj.IsValid = 1;

                            string msg = null;
                            if (!nPrj.Validate(db, out msg))
                            {
                                throw new Exception(msg);
                            }

                            db.Project.Add(nPrj);
                        }
                        // 修改
                        else
                        {
                            var oPrj = db.Project.Find(nPrj.Id);
                            if (oPrj == null)
                            {
                                throw new Exception("未找到要修改的数据！");
                            }

                            var prj = JToken.Parse(project);

                            foreach (JProperty p in prj)
                            {
                                var name = p.Name;
                                var property = properties.Where(pt => pt.Name == name).FirstOrDefault();
                                if (property != null)
                                {
                                    property.SetValue(oPrj, property.GetValue(nPrj));
                                }
                            }

                            string msg = null;
                            if (!oPrj.Validate(db, out msg))
                            {
                                throw new Exception(msg);
                            }

                            rp = oPrj;
                        }
                        db.SaveChanges();
                        RtObj.Serialize(rp, out s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }

            return Content(s);
        }

        public ActionResult GetProject(string id, bool includeWorkers = false)
        {
            string s = null;
            try
            {
                if (string.IsNullOrEmpty(id))
                {
                    throw new Exception("参数不正确");
                }
                else
                {
                    using (var db = PCDbContext.NewDbContext)
                    {
                        var prj = includeWorkers ? db.Project.Include("Workers").Where(p => p.Id == id).FirstOrDefault() : db.Project.Find(id);
                        RtObj.Serialize(prj, out s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }


        public ActionResult RemoveProject(string id)
        {
            string s = null;
            try
            {
                if (string.IsNullOrEmpty(id))
                {
                    throw new Exception("参数不正确");
                }
                else
                {
                    using (var ts = new TransactionScope())
                    {
                        using (var db = PCDbContext.NewDbContext)
                        {

                            var prj = db.Project.Find(id);
                            if (prj != null)
                            {
                                db.Project.Remove(prj);
                                // 删除项目员工关系
                                db.Database.ExecuteSqlCommand("delete project_worker where projectid=@projectid", new SqlParameter("@projectid", id));
                                // 删除该员工所有该项目科研数据
                                db.Database.ExecuteSqlCommand("delete worktime where projectid=@projectid", new SqlParameter("@projectid", id));
                                db.SaveChanges();
                                ts.Complete();
                            }
                            RtObj.Serialize("", out s);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);

        }

        public ActionResult GetProjectWorker(string prjId)
        {
            string s = null;
            try
            {
                if (string.IsNullOrEmpty(prjId))
                {
                    throw new Exception("参数不正确");
                }
                else
                {
                    using (var db = PCDbContext.NewDbContext)
                    {
                        var workers = db.Database.SqlQuery<Project_Worker>(@"select wk.Name,pw.* from Project_Worker pw
  left join  Worker wk on pw.WorkerId=wk.Id
  where pw.ProjectId=@prjId
  order by pw.[Index] asc,wk.Name asc", new SqlParameter("@prjId", prjId)).ToList();

                        for (int i = 0, j = workers.Count; i < j; i++)
                        {
                            workers[i].Index = workers[i].Index ?? i;
                        }

                        RtObj.Serialize(workers, out s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }

        public ActionResult SaveProjectWorkers(List<Project_Worker> projectWorkers)
        {
            string s = null;
            try
            {
                if (projectWorkers == null)
                {
                    throw new Exception("数据有误");
                }
                else
                {
                    using (var db = PCDbContext.NewDbContext)
                    {
                        var sql = string.Empty;
                        var sql1 = "update Project_Worker set [Index]={0} where WorkerId='{1}' and ProjectId='{2}';";
                        var sql2 = "update Project_Worker set ProjectRole='{0}' where WorkerId='{1}' and ProjectId='{2}';";
                        var sql3 = "update Project_Worker set [Index]={0},ProjectRole='{1}' where WorkerId='{2}' and ProjectId='{3}';";
                        foreach (var pw in projectWorkers)
                        {
                            if (pw.Index != null && !string.IsNullOrEmpty(pw.ProjectRole))
                            {
                                sql += string.Format(sql3, pw.Index, pw.ProjectRole, pw.WorkerId, pw.ProjectId);
                            }
                            else if (pw.Index != null)
                            {
                                sql += string.Format(sql1, pw.Index, pw.WorkerId, pw.ProjectId);
                            }
                            else if (!string.IsNullOrEmpty(pw.ProjectRole))
                            {
                                sql += string.Format(sql2, pw.ProjectRole, pw.WorkerId, pw.ProjectId);
                            }
                        }
                        int x = db.Database.ExecuteSqlCommand(sql);
                        RtObj.Serialize("", out s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }
    }
}