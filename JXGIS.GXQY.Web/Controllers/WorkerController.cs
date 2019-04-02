using JXGIS.GXQY.Web.Base;
using JXGIS.GXQY.Web.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Transactions;
using System.Web;
using System.Web.Mvc;

namespace JXGIS.GXQY.Web.Controllers
{
    public class WorkerController : Controller
    {
        public static PropertyInfo[] properties = typeof(Worker).GetProperties();


        public ActionResult Index()
        {
            return View();
        }


        public ActionResult WorkerList()
        {
            return PartialView();
        }

        public ActionResult WorkerForm()
        {
            return PartialView();
        }

        public ActionResult GetWorkers()
        {
            string s = null;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    var workers = db.Worker.OrderBy(w => w.Name).ToList();
                    RtObj.Serialize(workers, out s);
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }


        public ActionResult ModifyWorker(string worker)
        {
            string s = null;
            try
            {
                var nData = Newtonsoft.Json.JsonConvert.DeserializeObject<Worker>(worker);
                if (nData != null)
                {
                    using (var db = PCDbContext.NewDbContext)
                    {
                        Worker wk = nData;
                        // 新增
                        if (string.IsNullOrEmpty(nData.Id))
                        {
                            nData.Id = Guid.NewGuid().ToString();

                            string msg = null;
                            if (!nData.Validate(db, out msg))
                            {
                                throw new Exception(msg);
                            }

                            db.Worker.Add(nData);
                        }
                        // 修改
                        else
                        {
                            var oData = db.Worker.Find(nData.Id);
                            if (oData == null)
                            {
                                throw new Exception("未找到要修改的数据！");
                            }

                            var jToken = JToken.Parse(worker);

                            foreach (JProperty p in jToken)
                            {
                                var name = p.Name;
                                var property = properties.Where(pt => pt.Name == name).FirstOrDefault();
                                if (property != null)
                                {
                                    property.SetValue(oData, property.GetValue(nData));
                                }
                            }

                            string msg = null;
                            if (!oData.Validate(db, out msg))
                            {
                                throw new Exception(msg);
                            }

                            wk = oData;
                        }
                        db.SaveChanges();
                        RtObj.Serialize(wk, out s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }

            return Content(s);
        }

        public ActionResult GetWorker(string id)
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
                        var wk = db.Worker.Find(id);
                        RtObj.Serialize(wk, out s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }

        public ActionResult RemoveWorker(string id)
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

                            var wk = db.Worker.Find(id);
                            if (wk != null)
                            {
                                // 删除员工
                                db.Worker.Remove(wk);
                                // 删除项目员工关系
                                db.Database.ExecuteSqlCommand("delete project_worker where workerid=@workerid", new SqlParameter("@workerid", id));
                                // 删除该员工所有该项目科研数据
                                db.Database.ExecuteSqlCommand("delete worktime where workerid=@workerid", new SqlParameter("@workerid", id));
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

        public ActionResult AddToProject(string prjId, string wkId)
        {
            string s = null;
            try
            {
                if (string.IsNullOrEmpty(prjId) || string.IsNullOrEmpty(wkId))
                {
                    throw new Exception("参数不正确");
                }
                else
                {
                    using (var ts = new TransactionScope())
                    {
                        using (var db = PCDbContext.NewDbContext)
                        {
                            // 删除项目员工关系
                            var cnt = db.Database.SqlQuery<int>("select count(1) from  project_worker where workerId=@workerId and projectId=@projectId", new SqlParameter("@workerId", wkId), new SqlParameter("@projectId", prjId)).FirstOrDefault();
                            if (cnt > 0)
                            {
                                throw new Exception("人员已存在该项目中");
                            }
                            db.Database.ExecuteSqlCommand("insert into project_worker(workerId,projectId) values(@workerId,@projectId)", new SqlParameter("@workerId", wkId), new SqlParameter("@projectId", prjId));
                            db.SaveChanges();
                            ts.Complete();
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


        public ActionResult RemoveWorkerFromProject(string prjId, string wkId)
        {
            string s = null;
            try
            {
                if (string.IsNullOrEmpty(prjId) || string.IsNullOrEmpty(wkId))
                {
                    throw new Exception("参数不正确");
                }
                else
                {
                    using (var ts = new TransactionScope())
                    {
                        using (var db = PCDbContext.NewDbContext)
                        {
                            var count = db.Database.SqlQuery<int>("select count(1) from worktime where projectid=@projectId and workerid=@workerId", new SqlParameter("@projectId", prjId), new SqlParameter("@workerId", wkId)).FirstOrDefault();
                            if (count != 0)
                                throw new Exception("该人员已参与本项目实际工作，如需删除请联系管理员！");

                            db.Database.ExecuteSqlCommand("delete project_worker where projectid=@projectId and workerid=@workerId", new SqlParameter("@projectId", prjId), new SqlParameter("@workerId", wkId));

                            db.SaveChanges();
                            ts.Complete();
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
    }
}