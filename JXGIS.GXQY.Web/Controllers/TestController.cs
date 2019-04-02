using Aspose.Cells;
using JXGIS.GXQY.Web.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace JXGIS.GXQY.Web.Controllers
{
    public class WorkTime3
    {
        public string WorkerId { get; set; }

        public string WorkerName { get; set; }

        public int Date { get; set; }

        public int Day { get; set; }

        public int Month { get; set; }

        public string WorkType { get; set; }

        public string PWorkType { get; set; }

    }

    public class WorkerSalary3
    {
        public string WorkerId { get; set; }

        public string WorkerName { get; set; }

        public double BasePay { get; set; }

        public double Bonus { get; set; }
    }


    public class WorkMonth3
    {
        public string WorkerId { get; set; }

        public string WorkerName { get; set; }

        public List<WorkTime3> WorkTime { get; set; }

        public int WorkDay { get; set; }

        public int ResearchDay { get; set; }

        public int PResearchDay { get; set; }

        public double Bonus_R { get; set; }

        public double BasePay_R { get; set; }

        public double Bonus_W { get; set; }

        public double BasePay_W { get; set; }

        public double Bonus { get; set; }

        public double BasePay { get; set; }
    }
    public class Department
    {
        public string Id { get; set; }

        public string P_Id { get; set; }

        public string Name { get; set; }

        public int Level { get; set; }

        public Department PDepartment { get; set; }

        public List<Department> SubDepartments { get; set; }
    }
    public class TestController : Controller
    {


        // GET: Test
        public ActionResult Test()
        {
            //GetGZR();
            //GetProjectMonthTable();
            //GetMonthTable();
            //ImportWorker();

            //            using (var db = PCDbContext.NewDbContext)
            //            {
            //                var sql = @"with dps as(
            //select *,0 as level from department where p_id is null
            //union all
            //select d.* ,level+1 as level from department d inner join dps on d.p_id=dps.id
            //)
            //select * from dps;";

            //                var dpts = db.Database.SqlQuery<Department>(sql).ToList();

            //                var dpt = new Department()
            //                {
            //                    Level = -1,
            //                    Id = null,
            //                };
            //                GetNodes(dpt, dpts);
            //            }

            return Content(null);
        }

        public void GetNodes(Department dpt, List<Department> dpts)
        {
            var dpts0 = new List<Department>();
            var dpts1 = new List<Department>();

            foreach (var t in dpts) {
                if (t.Level == dpt.Level + 1 && t.P_Id == dpt.Id)
                {
                    dpts0.Add(t);
                }
                else {
                    dpts1.Add(t);
                }
            }


            dpt.SubDepartments = dpts0;
            foreach (var dp in dpts0)
            {
                GetNodes(dp, dpts1);
            }
        }

        public void GetProjectMonthTable()
        {
            string prjId = "1";
            int month = 201805;

            var sqlWT = @"
--获取工作清单
with
--日期 
dt as (select dt.Date,dt.Month,dt.type from DateType dt where month=@month),
--项目人员
wk as (select distinct wk.Id,wk.Name from Project_Worker pw left join worker wk on pw.WorkerId=wk.id where pw.ProjectId=@prjId),
--本月科研情况
wt as (
select workerid,date,WorkType from worktime wt 
where wt.month=@month
),
--本项目科研情况
wt1 as (
select workerid,date,WorkType from worktime wt 
where wt.month=@month and wt.ProjectId=@prjId
)
--人员本月科研情况
select t.date,t.workerid,t.workername,
(case when wt.WorkType is null then t.type else '科研' end) worktype,
(case when wt1.WorkType is null then t.type else '科研' end) pworktype
 from (select dt.*,wk.id workerid,wk.name workername from dt,wk) t
left join wt on t.date=wt.date and t.workerid=wt.workerid
left join wt1 on t.date=wt1.date and t.workerid=wt1.workerid";
            var sqlWS = @"
select t.*,ws.BasePay,ws.bonus from (select distinct wk.Id workerid,wk.Name workername from Project_Worker pw left join worker wk on pw.WorkerId=wk.id where pw.ProjectId=@prjId) t 
left join WorkerSalary ws on ws.WorkerId=t.workerid where ws.Month=@month
";
            var sqlSTWT = @"select dt.Date,dt.Month,dt.type worktype from DateType dt where month=@month order by dt.Date";

            List<WorkTime3> wts = null;
            List<WorkerSalary3> wss = null;
            List<WorkTime3> stwts = null;
            Project prj = null;

            using (var db = PCDbContext.NewDbContext)
            {
                prj = db.Project.Find(prjId);
                stwts = db.Database.SqlQuery<WorkTime3>(sqlSTWT, new SqlParameter("@month", month)).ToList();
                wts = db.Database.SqlQuery<WorkTime3>(sqlWT, new SqlParameter("@prjId", prjId), new SqlParameter("@month", month)).ToList();
                wss = db.Database.SqlQuery<WorkerSalary3>(sqlWS, new SqlParameter("@prjId", prjId), new SqlParameter("@month", month)).ToList();
            }

            var wks = (from wt in wts
                       group wt by new { wt.WorkerId, wt.WorkerName } into g
                       select new WorkMonth3
                       {
                           WorkerId = g.Key.WorkerId,
                           WorkerName = g.Key.WorkerName,
                           WorkTime = g.OrderBy(t => t.Date).ToList(),
                           WorkDay = g.Where(t => t.WorkType != "节假日").Count(),
                           ResearchDay = g.Where(t => t.WorkType == "科研").Count(),
                           PResearchDay = g.Where(t => t.PWorkType == "科研").Count(),
                       }).ToList();

            var wkms = (from wk in wks
                        from ws in wss
                        where wk.WorkerId == ws.WorkerId
                        select new WorkMonth3
                        {
                            WorkerId = wk.WorkerId,
                            WorkerName = wk.WorkerName,
                            WorkTime = wk.WorkTime,
                            WorkDay = wk.WorkDay,
                            ResearchDay = wk.ResearchDay,
                            PResearchDay = wk.PResearchDay,
                            Bonus_R = wk.PResearchDay * (ws.Bonus / wk.WorkDay),
                            BasePay_R = wk.PResearchDay * (ws.BasePay / wk.WorkDay),
                        }).OrderBy(t => t.WorkerName).ToList();

            var sum = (from wk in wkms
                       group wk by 1 into g
                       select new WorkMonth3
                       {
                           WorkerId = "HJ",
                           WorkerName = "合计",
                           WorkDay = g.Sum(t => t.WorkDay),
                           ResearchDay = g.Sum(t => t.ResearchDay),
                           PResearchDay = g.Sum(t => t.PResearchDay),
                           Bonus_R = g.Sum(t => t.Bonus_R),
                           BasePay_R = g.Sum(t => t.BasePay_R),
                       }).FirstOrDefault();

            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = (int)(month / 100) + "年" + month % 100 + "月项目考勤表";

            Style st1 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st1.HorizontalAlignment = TextAlignmentType.Center;//文字居中
            st1.Font.Name = "等线";//文字字体
            st1.Font.Size = 12;//文字大小
            st1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线

            Style st4 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st4.HorizontalAlignment = TextAlignmentType.Center;//文字居中
            st4.Font.Name = "等线";//文字字体
            st4.Font.Size = 12;//文字大小

            Style st2 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st2.HorizontalAlignment = TextAlignmentType.Right;//文字居中
            st2.Font.Name = "等线";//文字字体
            st2.Font.Size = 12;//文字大小
            st2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线
            st2.Number = 4;

            Style st3 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st3.HorizontalAlignment = TextAlignmentType.Center;//文字居中
            st3.Font.Name = "等线";//文字字体
            st3.Font.Size = 12;//文字大小
            st3.Font.Color = System.Drawing.Color.Red;
            st3.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st3.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st3.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线


            Cells cells = sheet.Cells;//单元格

            int d = stwts.Count;
            cells.Merge(0, 0, 1, 2); cells[0, 0].PutValue("项目名称");
            cells.Merge(0, 2, 1, 20); cells[0, 2].PutValue(prj.Name);
            cells.Merge(0, 2 + 20, 1, d + 5 - 20 - 3); cells[0, 2 + 20].PutValue("（研发人员）考勤表");

            cells.Merge(1, 0, 1, 2); cells[1, 0].PutValue("项目编号");
            cells.Merge(1, 2, 1, 20); cells[1, 2].PutValue(prj.Id);
            cells.Merge(1, 2 + 20, 1, d + 5 - 20 - 3); cells[1, 2 + 20].PutValue((int)(month / 100) + "年" + month % 100 + "月");

            cells.Merge(2, 0, 2, 1); cells[2, 0].PutValue("序号");
            cells.Merge(2, 1, 2, 1); cells[2, 1].PutValue("姓名");
            cells.Merge(2, 2, 1, stwts.Count); cells[2, 2].PutValue("出    勤    情    况");
            cells.Merge(2, 2 + d, 2, 1); cells[2, 2 + d].PutValue("研发出勤\n（天）");
            cells.Merge(2, 2 + d + 1, 2, 1); cells[2, 2 + d + 1].PutValue("总工作日\n（天）");


            cells.Merge(0, d + 4, 2, 3); cells[0, d + 4].PutValue("科研费用");
            cells.Merge(2, d + 4, 2, 1); cells[2, d + 4].PutValue("基本工资");
            cells.Merge(2, d + 4 + 1, 2, 1); cells[2, d + 4 + 1].PutValue("绩效工资");
            cells.Merge(2, d + 4 + 2, 2, 1); cells[2, d + 4 + 2].PutValue("合计");
            for (int i = 0; i <= 1 + wkms.Count + 3; i++)
            {
                for (int j = 0; j <= 2 + d + 4; j++)
                {
                    cells[i, j].SetStyle(st1);
                }
            }

            for (int i = 4; i <= 1 + wkms.Count + 3; i++)
            {
                for (int j = d + 4; j <= d + 6; j++)
                {
                    cells[i, j].SetStyle(st2);
                }
            }

            for (int i = 0; i < d; i++)
            {
                cells[3, 2 + i].PutValue(stwts[i].Date % 100);
                if (stwts[i].WorkType != "工作日") cells[3, 2 + i].SetStyle(st3);
            }



            for (int j = 1; j <= wkms.Count; j++)
            {
                var r0 = j + 3;
                var wkm = wkms[j - 1];
                cells[r0, 0].PutValue(j);
                cells[r0, 1].PutValue(wkm.WorkerName);

                for (int i = 0; i < d; i++)
                {
                    var wtm = wkm.WorkTime[i];
                    cells[r0, 2 + i + 1].PutValue(wtm.PWorkType == "科研" ? "√" : "");
                }

                cells[r0, 2 + d].PutValue(wkm.PResearchDay);
                cells[r0, 2 + d + 1].PutValue(wkm.WorkDay);
                cells[r0, 2 + d + 2].PutValue(wkm.BasePay_R);
                cells[r0, 2 + d + 3].PutValue(wkm.Bonus_R);
                cells[r0, 2 + d + 4].PutValue(wkm.BasePay_R + wkm.Bonus_R);
            }

            int r = 3 + wkms.Count + 1;

            cells.Merge(r, 0, 1, 2); cells[r, 0].PutValue("合计");
            cells[r, 2 + d].PutValue(sum.PResearchDay);
            cells[r, 2 + d + 1].PutValue(sum.WorkDay);
            cells[r, 2 + d + 2].PutValue(sum.BasePay_R);
            cells[r, 2 + d + 3].PutValue(sum.Bonus_R);
            cells[r, 2 + d + 4].PutValue(sum.BasePay_R + sum.Bonus_R);

            cells[r + 2, 2 + d - 20].PutValue("项目组长签字："); cells[r + 2, 2 + d - 20].SetStyle(st4);
            cells[r + 2, 2 + d - 5].PutValue("考勤员签字："); cells[r + 2, 2 + d - 5].SetStyle(st4);


            for (var i = 2; i < 2 + d; i++)
            {
                cells.SetColumnWidth(i, 3);
            }
            cells.SetColumnWidth(0, 5);
            cells.SetColumnWidth(1, 15);
            for (var i = 2 + d; i < 2 + d + 5; i++)
            {
                cells.SetColumnWidth(i, 12);
            }

            workbook.Save(@"D:\test.xls");
        }


        public void GetMonthTable()
        {
            int month = 201812;
            var sqlWM = @"with
--人员当月研发情况
wt as (select wt.WorkerId,count(1) yf from WorkTime wt where wt.Month=@month group by wt.workerid),
--日期 
dt as (select dt.Date,dt.Month,dt.type from DateType dt where month=@month),
--工作类型
wtp as (
select t.*,(case when wt.WorkType is not null then '科研' else t.Type end) worktype from (
select dt.*,wt.WorkerId from dt,wt )t left join worktime wt on wt.WorkerId=t.workerid and wt.date=t.date),
--科研，工作日数量
mt as (
select workerid,sum(case when worktype='节假日' then 0 else 1 end) gzr,sum(case when worktype='科研' then 1 else 0 end) ky from wtp
group by workerid)

select mt.workerid,wk.name workername,gzr workday,ky researchday,ky*ws.basepay/gzr basepay_r,ky*ws.bonus/gzr bonus_r,ws.basepay,ws.bonus from mt 
left join WorkerSalary ws on ws.WorkerId=mt.WorkerId 
left join Worker wk on mt.WorkerId=wk.Id
where ws.Month=@month
order by workername";
            List<WorkMonth3> wms = null;
            using (var db = PCDbContext.NewDbContext)
            {
                wms = db.Database.SqlQuery<WorkMonth3>(sqlWM, new SqlParameter("@month", month)).ToList();

                foreach (var wm in wms)
                {
                    wm.BasePay_W = wm.BasePay - wm.BasePay_R;
                    wm.Bonus_W = wm.Bonus - wm.Bonus_R;
                }
            }
            var hj = new WorkMonth3()
            {
                WorkerId = "HJ",
                WorkerName = "合计",
                BasePay = wms.Sum(t => t.BasePay),
                BasePay_R = wms.Sum(t => t.BasePay_R),
                BasePay_W = wms.Sum(t => t.BasePay_W),
                Bonus = wms.Sum(t => t.Bonus),
                Bonus_R = wms.Sum(t => t.Bonus_R),
                Bonus_W = wms.Sum(t => t.Bonus_W),
                WorkDay = wms.Sum(t => t.WorkDay),
                ResearchDay = wms.Sum(t => t.ResearchDay)
            };

            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = (int)(month / 100) + "年" + month % 100 + "月科研、生产工资";

            Cells cells = sheet.Cells;//单元格

            Style st1 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st1.HorizontalAlignment = TextAlignmentType.Center;//文字居中
            st1.Font.Name = "等线";//文字字体
            st1.Font.Size = 12;//文字大小
            st1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线

            Style st2 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st2.HorizontalAlignment = TextAlignmentType.Right;//文字居中
            st2.Font.Name = "等线";//文字字体
            st2.Font.Size = 12;//文字大小
            st2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线
            st2.Number = 4;

            // 表头
            cells.Merge(0, 0, 1, 8);
            cells[0, 0].PutValue((int)(month / 100) + "年" + month % 100 + "月科研、生产工资");
            //cells[0, 0].SetStyle(st1);

            cells.Merge(1, 0, 2, 1);
            cells[1, 0].PutValue("序号");
            //cells[1, 0].SetStyle(st1);

            cells.Merge(1, 1, 2, 1);
            cells[1, 1].PutValue("姓名");
            //cells[1, 1].SetStyle(st1);

            cells.Merge(1, 2, 1, 3);
            cells[1, 2].PutValue("基本工资");
            //cells[1, 2].SetStyle(st1);

            cells.Merge(1, 5, 1, 3);
            cells[1, 5].PutValue("绩效奖金"); //cells[1, 5].SetStyle(st1);
            cells[2, 2].PutValue("研发费用"); //cells[2, 2].SetStyle(st1);
            cells[2, 3].PutValue("生产费用"); //cells[2, 3].SetStyle(st1);
            cells[2, 4].PutValue("合计"); //cells[2, 4].SetStyle(st1);
            cells[2, 5].PutValue("研发费用"); //cells[2, 5].SetStyle(st1);
            cells[2, 6].PutValue("生产费用"); //cells[2, 6].SetStyle(st1);
            cells[2, 7].PutValue("合计");// cells[2, 7].SetStyle(st1);
            int i = 3;
            for (int j = 0; i < 3 + wms.Count; i++, j++)
            {
                var w = wms[j];
                cells[i, 0].PutValue(j + 1);// cells[i, 0].SetStyle(st1);
                cells[i, 1].PutValue(w.WorkerName); //cells[i, 1].SetStyle(st1);
                cells[i, 2].PutValue(w.BasePay_R); //cells[i, 2].SetStyle(st2);
                cells[i, 3].PutValue(w.BasePay_W); //cells[i, 3].SetStyle(st2);
                cells[i, 4].PutValue(w.BasePay); //cells[i, 4].SetStyle(st2);
                cells[i, 5].PutValue(w.Bonus_R);// cells[i, 5].SetStyle(st2);
                cells[i, 6].PutValue(w.Bonus_W); //cells[i, 6].SetStyle(st2);
                cells[i, 7].PutValue(w.Bonus);// cells[i, 7].SetStyle(st2);
            }
            cells.Merge(i, 0, 1, 2);
            cells[i, 0].PutValue("合计"); //cells[i, 0].SetStyle(st1);
            cells[i, 2].PutValue(hj.BasePay_R); //cells[i, 2].SetStyle(st2);
            cells[i, 3].PutValue(hj.BasePay_W); //cells[i, 3].SetStyle(st2);
            cells[i, 4].PutValue(hj.BasePay); //cells[i, 4].SetStyle(st2);
            cells[i, 5].PutValue(hj.Bonus_R); //cells[i, 5].SetStyle(st2);
            cells[i, 6].PutValue(hj.Bonus_W); //cells[i, 6].SetStyle(st2);
            cells[i, 7].PutValue(hj.Bonus); //cells[i, 7].SetStyle(st2);

            for (int n = 0; n < 8; n++)
            {
                for (int m = 0; m < i + 1; m++)
                {
                    if (m >= 3 && n >= 2)
                        cells[m, n].SetStyle(st2);
                    else
                        cells[m, n].SetStyle(st1);

                }

            }
            sheet.AutoFitRows();
            sheet.AutoFitColumns();
            workbook.Save(@"D:\test.xls");
        }

        private void ImportWorker()
        {
            string filePath = "E:\\x.xlsx";
            DataTable dt = new DataTable();
            OleDbConnection con = new OleDbConnection(
                string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended properties=\"Excel 12.0;Imex=1;HDR=Yes;\"", filePath));
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [Sheet1$]", con);
            adapter.Fill(dt);
            using (var db = PCDbContext.NewDbContext)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    var id = dr["ID"].ToString();
                    var name = dr["Name"].ToString();
                    var worker = new Worker
                    {
                        Id = id,
                        Name = name
                    };
                    db.Worker.Add(worker);

                    for (int i = 201801; i <= 201809; i++)
                    {
                        var fgz = "GZ-" + i;
                        var fjx = "JX-" + i;
                        var sgz = dr[fgz] == DBNull.Value ? "0" : dr[fgz].ToString();
                        var sjx = dr[fjx] == DBNull.Value ? "0" : dr[fjx].ToString();
                        sgz = string.IsNullOrEmpty(sgz) ? "0" : sgz;
                        sjx = string.IsNullOrEmpty(sjx) ? "0" : sjx;

                        var gz = float.Parse(sgz);
                        var jx = float.Parse(sjx);

                        var workerSalary = new WorkerSalary
                        {
                            Id = Guid.NewGuid().ToString(),
                            WorkerId = id,
                            Month = i,
                            BasePay = gz,
                            Bonus = jx

                        };
                        db.WorkerSalary.Add(workerSalary);
                    }
                }
                db.SaveChanges();
            }

        }

        private void GetGZR()
        {
            DateTime start = DateTime.Parse("2019-01-21");
            DateTime end = DateTime.Parse("2020-01-20");

            using (var db = PCDbContext.NewDbContext)
            {
                for (var i = start; i <= end;)
                {

                    db.DateType.Add(new DateType
                    {
                        Id = Guid.NewGuid().ToString(),
                        Date = i.Year * 10000 + i.Month * 100 + i.Day,
                        Month = i.Day > 20 ? (i.AddMonths(1).Year * 100 + (i.Month + 1) % 12) : (i.Year * 100 + i.Month),
                        Type = i.DayOfWeek == DayOfWeek.Sunday || i.DayOfWeek == DayOfWeek.Saturday ? "假日" : "工作日"
                    });

                    i = i.AddDays(1);
                }

                db.SaveChanges();
            }
        }

        public ActionResult TestPartialView()
        {
            return PartialView();
        }
    }
}