using JXGIS.GXQY.Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace JXGIS.GXQY.Web.Controllers
{
    public class ImportController : Controller
    {
        public static Dictionary<string, object> Data = null;

        // GET: Import
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Upload()
        {
            var date = Request.Form["date"];
            var d = DateTime.Parse(date);
            var f = Request.Files.Get("file");

            var wb = new Aspose.Cells.Workbook(f.InputStream);
            var ws = wb.Worksheets[0];

            int i = 0;
            int WorkerIdIdx = -1,
                NameIdx = -1,
                BasePayIdx = -1,
                BonusIdx = -1,
                AccumulationFundIdx = -1,
                S1Idx = -1,
                S2Idx = -1,
                S3Idx = -1,
                S4Idx = -1,
                S5Idx = -1;

            while (ws.Cells[0, i] != null && !string.IsNullOrEmpty(ws.Cells[0, i].StringValue))
            {
                var x = ws.Cells[0, i].Value.ToString();

                if (x.Contains("编号")) WorkerIdIdx = i;
                if (x.Contains("姓名")) NameIdx = i;
                if (x.Contains("工资")) BasePayIdx = i;
                if (x.Contains("公积金")) AccumulationFundIdx = i;
                if (x.Contains("绩效")) BonusIdx = i;
                if (x.Contains("医")) S1Idx = i;
                if (x.Contains("老")) S2Idx = i;
                if (x.Contains("失")) S3Idx = i;
                if (x.Contains("育")) S4Idx = i;
                if (x.Contains("伤")) S5Idx = i;
                i++;
            }
            i = 1;
            List<string> errors = new List<string>();
            List<string> infos = new List<string>();
            List<WorkerSalary> wsList = new List<WorkerSalary>();
            List<Worker> workers = new List<Worker>();

            while (ws.Cells[i, 0] != null && !string.IsNullOrEmpty(ws.Cells[i, 0].StringValue))
            {
                var w = new WorkerSalary()
                {
                    Id = Guid.NewGuid().ToString(),
                    Month = int.Parse(d.AddMonths(-1).ToString("yyyyMM")),
                    MonthX = int.Parse(d.ToString("yyyyMM"))
                };

                if (string.IsNullOrEmpty(ws.Cells[i, WorkerIdIdx].StringValue))
                    errors.Add($"第{i}行“WorkerId”为空");
                else
                    w.WorkerId = ws.Cells[i, WorkerIdIdx].StringValue;

                if (string.IsNullOrEmpty(ws.Cells[i, NameIdx].StringValue))
                    errors.Add($"第{i}行“Name”为空");
                else
                    w.WorkerName = ws.Cells[i, NameIdx].StringValue;

                if (string.IsNullOrEmpty(ws.Cells[i, BasePayIdx].StringValue))
                    errors.Add($"第{i}行“BasePay”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, BasePayIdx].StringValue, out v))
                    {
                        w.BasePay = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“BasePay”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, AccumulationFundIdx].StringValue))
                    errors.Add($"第{i}行“AccumulationFund”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, AccumulationFundIdx].StringValue, out v))
                    {
                        w.AccumulationFund = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“AccumulationFund”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, BonusIdx].StringValue))
                    errors.Add($"第{i}行“Bonus”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, BonusIdx].StringValue, out v))
                    {
                        w.Bonus = v;
                        w.Bonus1 = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“Bonus”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, S1Idx].StringValue))
                    errors.Add($"第{i}行“SS1”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, S1Idx].StringValue, out v))
                    {
                        w.SS1 = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“SS1”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, S2Idx].StringValue))
                    errors.Add($"第{i}行“SS2”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, S2Idx].StringValue, out v))
                    {
                        w.SS2 = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“SS2”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, S3Idx].StringValue))
                    errors.Add($"第{i}行“SS3”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, S3Idx].StringValue, out v))
                    {
                        w.SS3 = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“SS3”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, S4Idx].StringValue))
                    errors.Add($"第{i}行“SS4”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, S4Idx].StringValue, out v))
                    {
                        w.SS4 = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“SS4”格式不正确");
                    }
                }

                if (string.IsNullOrEmpty(ws.Cells[i, S5Idx].StringValue))
                    errors.Add($"第{i}行“SS5”为空");
                else
                {
                    double v = 0;
                    if (double.TryParse(ws.Cells[i, S5Idx].StringValue, out v))
                    {
                        w.SS5 = v;
                    }
                    else
                    {
                        errors.Add($"第{i}行“SS5”格式不正确");
                    }
                }
                w.SocialSecurity = w.SS1 + w.SS2 + w.SS3 + w.SS4 + w.SS5;
                wsList.Add(w);
                i++;
            }

            var cf = (from t in wsList
                      group t by new
                      {
                          t.WorkerId,
                          t.WorkerName
                      } into g
                      where g.Count() > 1
                      select new { g.Key.WorkerId, g.Key.WorkerName }).ToList();
            foreach (var c in cf)
            {
                errors.Add($"WokerId，WorkName重复${c.WorkerName}(${c.WorkerId})");
            }

            using (var db = new PCDbContext())
            {
                var wks = db.Worker.ToList();

                foreach (var t in wsList)
                {
                    var wn = wks.Where(x => x.Id == t.WorkerId).FirstOrDefault();
                    if (wn == null)
                    {
                        infos.Add($"新增员工{t.WorkerName}({t.WorkerId})");
                        workers.Add(new Worker()
                        {
                            Id = t.WorkerId,
                            Name = t.WorkerName
                        });
                    }
                    if (wn != null && wn.Name != t.WorkerName)
                    {
                        errors.Add($"Id为{t.WorkerId}与系统内员工名称不一致");
                    }
                }
            }

            var rt = new Dictionary<string, object>();
            rt.Add("Errors", errors);
            rt.Add("Infos", infos);
            rt.Add("WorkerSalary", wsList);
            rt.Add("Workers", workers);

            Data = rt;

            var s = Newtonsoft.Json.JsonConvert.SerializeObject(rt);
            return Content(s);
        }

        public ActionResult Update(string date)
        {
            using (var db = new PCDbContext())
            {
                // 数据备份
                var d = DateTime.Parse(date).ToString("yyyyMMdd");
                var row = db.Database.SqlQuery<int>($"select count(1) row where OBJECT_ID('Worker{d}','u') is not null").FirstOrDefault();
                if (row == 0)
                {
                    db.Database.ExecuteSqlCommand($"select * into Worker{d} from worker");
                }

                row = db.Database.SqlQuery<int>($"select count(1) row where OBJECT_ID('WorkerSalary{d}','u') is not null").FirstOrDefault();
                if (row == 0)
                {
                    db.Database.ExecuteSqlCommand($"select * into WorkerSalary{d} from WorkerSalary");
                }

                // 导入数据
                var wks = Data["Workers"] as List<Worker>;
                var wss = Data["WorkerSalary"] as List<WorkerSalary>;
                db.Worker.AddRange(wks);
                db.WorkerSalary.AddRange(wss);

                db.SaveChanges();
            }

            return null;
        }
    }
}