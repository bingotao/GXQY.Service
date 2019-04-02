using Aspose.Cells;
using JXGIS.GXQY.Web.Base;
using JXGIS.GXQY.Web.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace JXGIS.GXQY.Web.Controllers
{
    public class ExportController : Controller
    {
        public class WorkTime3
        {
            public int? Index { get; set; }

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

            public double AccumulationFund { get; set; }

            public double SocialSecurity { get; set; }
        }


        public class WorkMonth3
        {
            public int? Index { get; set; }

            public string WorkerId { get; set; }

            public string WorkerName { get; set; }

            public List<WorkTime3> WorkTime { get; set; }

            public int WorkDay { get; set; }

            public int ResearchDay { get; set; }

            public int PResearchDay { get; set; }

            public double Bonus_R { get; set; }

            public double BasePay_R { get; set; }

            public double AccumulationFund_R { get; set; }

            public double SocialSecurity_R { get; set; }

            public double Bonus_W { get; set; }

            public double BasePay_W { get; set; }

            public double AccumulationFund_W { get; set; }

            public double SocialSecurity_W { get; set; }

            public double Bonus { get; set; }

            public double BasePay { get; set; }

            public double AccumulationFund { get; set; }

            public double SocialSecurity { get; set; }
        }

        public class MaintainProject
        {
            public string ID { get; set; }
            public string Name { get; set; }
            public string Department { get; set; }
            public string Workers { get; set; }

            public DateTime? StartTime { get; set; }

            public DateTime? EndTime { get; set; }
        }

        public ActionResult DownloadMaintainProjects(int month)
        {
            string s = null;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    var sql = @"with p as(
select * from project p where (p.endtime is null or p.EndTime>CONVERT(datetime,@month)) and p.starttime<CONVERT(datetime,@month2)),
pw as(
select p.id,w.name from p 
left join Project_Worker pw on p.Id=pw.ProjectId
left join Worker w on pw.WorkerId=w.id)

select pp.year+pp.serialnumber id,pp.name,pp.starttime,pp.endtime,pp.Department,x.workers from (
select id,STUFF(( SELECT '、'+ name FROM pw b WHERE b.id = a.id FOR XML PATH('')),1 ,1, '') workers from pw a
group by id)x left join Project pp on x.Id=pp.id
order by id";
                    var projs = db.Database.SqlQuery<MaintainProject>(sql, new SqlParameter("@month", month + "01"), new SqlParameter("@month2", DateTime.Parse($"{(int)(month / 100)}-{month % 100}-01").AddMonths(1).ToString("yyyyMMdd"))).ToList();
                    var fileName = (int)(month / 100) + "年" + (month % 100) + "月项目存续表";
                    var wk = GetMaintainProjectExcel(projs, fileName);
                    var ms = wk.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", fileName + ".xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }

        public Workbook GetMaintainProjectExcel(List<MaintainProject> prjs, string sheetName)
        {
            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = sheetName;

            Style st1 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st1.HorizontalAlignment = TextAlignmentType.Center;//文字居中
            st1.Font.Name = "等线";//文字字体
            st1.Font.Size = 12;//文字大小
            st1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线

            Style st4 = workbook.Styles[workbook.Styles.Add()];//新增样式
            st4.HorizontalAlignment = TextAlignmentType.Left;//文字左对齐
            st4.Font.Name = "等线";//文字字体
            st4.Font.Size = 12;//文字大小
            st4.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            st4.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            st4.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            st4.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线

            Cells cells = sheet.Cells;//单元格

            cells.Merge(0, 0, 1, 6); cells[0, 0].PutValue(sheetName);

            cells[1, 0].PutValue("编号"); cells[1, 0].SetStyle(st1);
            cells[1, 1].PutValue("名称"); cells[1, 1].SetStyle(st1);
            cells[1, 2].PutValue("起始时间"); cells[1, 2].SetStyle(st1);
            cells[1, 3].PutValue("截止时间"); cells[1, 3].SetStyle(st1);
            cells[1, 4].PutValue("所属研究中心"); cells[1, 4].SetStyle(st1);
            cells[1, 5].PutValue("参与人员"); cells[1, 5].SetStyle(st1);

            int rowindex = 2;
            foreach (var prj in prjs)
            {
                cells[rowindex, 0].PutValue(prj.ID); cells[rowindex, 0].SetStyle(st1);
                cells[rowindex, 1].PutValue(prj.Name); cells[rowindex, 1].SetStyle(st1);
                cells[rowindex, 2].PutValue(prj.StartTime != null ? ((DateTime)prj.StartTime).ToString("yyyy-MM-dd") : string.Empty); cells[rowindex, 2].SetStyle(st1);
                cells[rowindex, 3].PutValue(prj.EndTime != null ? ((DateTime)prj.EndTime).ToString("yyyy-MM-dd") : string.Empty); cells[rowindex, 3].SetStyle(st1);
                cells[rowindex, 4].PutValue(prj.Department); cells[rowindex, 4].SetStyle(st1);
                cells[rowindex, 5].PutValue(prj.Workers); cells[rowindex, 5].SetStyle(st4);

                rowindex += 1;
            }
            sheet.AutoFitColumns();
            return workbook;
        }


        public ActionResult DownloadProjectMonthTable1(string prjId, int month)
        {
            var s = string.Empty;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    var prj = db.Project.Find(prjId);
                    if (prj == null) throw new Exception("未能找到指定的项目");
                    var data = GetProjectMonthData(prjId, month);
                    var wk1 = GetProjectMonthWorkbook1(month, data);
                    var ms = wk1.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", $"[{prj.Year}-{prj.SerialNumber}]{prj.Name}-科研费用-工资公积金社保[{(int)(month / 100)}年{month % 100}月].xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex.Message, out s);
            }
            return Content(s);
        }

        public ActionResult DownloadProjectMonthTable2(string prjId, int month)
        {
            var s = string.Empty;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    var prj = db.Project.Find(prjId);
                    if (prj == null) throw new Exception("未能找到指定的项目");
                    var data = GetProjectMonthData(prjId, month);
                    var wk1 = GetProjectMonthWorkbook2(month, data);
                    var ms = wk1.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", $"[{prj.Year}-{prj.SerialNumber}]{prj.Name}-科研费用-绩效[{(int)(month / 100)}年{month % 100}月].xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex.Message, out s);
            }
            return Content(s);
        }

        public ActionResult DownloadMonthTable1(int month)
        {
            var s = string.Empty;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    var data = GetMonthData(month);
                    var wk = GetMonthWorkbook1(month, data);
                    var ms = wk.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", $"科研费用-工资公积金社保[{(int)(month / 100)}年{month % 100}月].xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex.Message, out s);
            }
            return Content(s);
        }

        public ActionResult DownloadMonthTable2(int month)
        {
            var s = string.Empty;
            try
            {
                using (var db = PCDbContext.NewDbContext)
                {
                    var data = GetMonthData(month);
                    var wk = GetMonthWorkbook2(month, data);
                    var ms = wk.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", $"科研费用-绩效[{(int)(month / 100)}年{month % 100}月].xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex.Message, out s);
            }
            return Content(s);
        }

        #region 废弃
        /*
        public ActionResult DownloadProjectMonthTable(string prjId, int? month)
        {
            string s = null;
            try
            {
                if (string.IsNullOrEmpty(prjId) || month == null)
                {
                    throw new Exception("参数不正确！");
                }
                using (var db = PCDbContext.NewDbContext)
                {
                    var prj = db.Project.Find(prjId);
                    var wk = GetProjectMonthTable(prjId, (int)month);
                    var ms = wk.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", prj.Name + (int)(month / 100) + "年" + (month % 100) + "月项目考勤表" + ".xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }

        public ActionResult DownloadMonthTable(int? month)
        {
            string s = null;
            try
            {
                if (month == null)
                {
                    throw new Exception("参数不正确！");
                }
                using (var db = PCDbContext.NewDbContext)
                {
                    var wk = GetMonthTable((int)month);
                    var ms = wk.SaveToStream();
                    ms.Seek(0, SeekOrigin.Begin);
                    return File(ms, "application/vnd.ms-excel", (int)(month / 100) + "年" + month % 100 + "月科研、生产工资" + ".xls");
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }
            return Content(s);
        }


        public Workbook GetProjectMonthTable(string prjId, int month)
        {
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
                            AccumulationFund_R = wk.PResearchDay * (ws.AccumulationFund / wk.WorkDay),
                            SocialSecurity_R = wk.PResearchDay * (ws.SocialSecurity / wk.WorkDay),
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
                           AccumulationFund_R = g.Sum(t => t.AccumulationFund_R),
                           SocialSecurity_R = g.Sum(t => t.SocialSecurity_R),
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
            cells.Merge(1, 2, 1, 20); cells[1, 2].PutValue($"{prj.Year}-{prj.SerialNumber}");
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
                    cells[r0, 2 + i].PutValue(wtm.PWorkType == "科研" ? "√" : "");
                }

                cells[r0, 2 + d].PutValue(wkm.PResearchDay);
                cells[r0, 2 + d + 1].PutValue(wkm.WorkDay);
                cells[r0, 2 + d + 2].PutValue(wkm.BasePay_R);
                cells[r0, 2 + d + 3].PutValue(wkm.Bonus_R);
                cells[r0, 2 + d + 4].PutValue(wkm.AccumulationFund_R);
                cells[r0, 2 + d + 5].PutValue(wkm.SocialSecurity_R);
                cells[r0, 2 + d + 6].PutValue(wkm.BasePay_R + wkm.Bonus_R + wkm.AccumulationFund_R + wkm.SocialSecurity_R);
            }

            int r = 3 + wkms.Count + 1;

            cells.Merge(r, 0, 1, 2); cells[r, 0].PutValue("合计");
            cells[r, 2 + d].PutValue(sum.PResearchDay);
            cells[r, 2 + d + 1].PutValue(sum.WorkDay);
            cells[r, 2 + d + 2].PutValue(sum.BasePay_R);
            cells[r, 2 + d + 3].PutValue(sum.Bonus_R);
            cells[r, 2 + d + 4].PutValue(sum.BasePay_R);
            cells[r, 2 + d + 5].PutValue(sum.Bonus_R);
            cells[r, 2 + d + 6].PutValue(sum.BasePay_R + sum.Bonus_R + sum.AccumulationFund_R + sum.SocialSecurity_R);

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

            return workbook;
        }


        public Workbook GetMonthTable(int month)
        {
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

select mt.workerid,wk.name workername,gzr workday,ky researchday,
ky*ws.basepay/gzr basepay_r,
ky*ws.bonus/gzr bonus_r,
ky*ws.AccumulationFund/gzr AccumulationFund_r,
ky*ws.SocialSecurity/gzr SocialSecurity_r,
ws.basepay,
ws.bonus,
ws.AccumulationFund,
ws.SocialSecurity from mt 
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
                    wm.AccumulationFund_W = wm.AccumulationFund - wm.AccumulationFund_R;
                    wm.SocialSecurity_W = wm.SocialSecurity - wm.SocialSecurity_R;
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
                AccumulationFund = wms.Sum(t => t.AccumulationFund),
                AccumulationFund_R = wms.Sum(t => t.AccumulationFund_R),
                AccumulationFund_W = wms.Sum(t => t.AccumulationFund_W),
                SocialSecurity = wms.Sum(t => t.SocialSecurity),
                SocialSecurity_R = wms.Sum(t => t.SocialSecurity_R),
                SocialSecurity_W = wms.Sum(t => t.SocialSecurity_W),

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
            cells.Merge(0, 0, 1, 8 + 2);
            cells[0, 0].PutValue((int)(month / 100) + "年" + month % 100 + "月科研、生产工资");
            //cells[0, 0].SetStyle(st1);

            cells.Merge(1, 0, 2, 1);
            cells[1, 0].PutValue("序号");
            //cells[1, 0].SetStyle(st1);

            cells.Merge(1, 1, 2, 1);
            cells[1, 1].PutValue("姓名");
            //cells[1, 1].SetStyle(st1);

            cells.Merge(1, 2, 2, 1);
            cells[1, 2].PutValue("研发天数");
            cells.Merge(1, 3, 2, 1);
            cells[1, 3].PutValue("总工作日");

            cells.Merge(1, 4, 1, 3);
            cells[1, 2 + 2].PutValue("基本工资");
            //cells[1, 2].SetStyle(st1);

            cells.Merge(1, 5 + 2, 1, 3);
            cells[1, 5 + 2].PutValue("绩效奖金"); //cells[1, 5].SetStyle(st1);
            cells[2, 2 + 2].PutValue("研发费用"); //cells[2, 2].SetStyle(st1);
            cells[2, 3 + 2].PutValue("生产费用"); //cells[2, 3].SetStyle(st1);
            cells[2, 4 + 2].PutValue("合计"); //cells[2, 4].SetStyle(st1);
            cells[2, 5 + 2].PutValue("研发费用"); //cells[2, 5].SetStyle(st1);
            cells[2, 6 + 2].PutValue("生产费用"); //cells[2, 6].SetStyle(st1);
            cells[2, 7 + 2].PutValue("合计");// cells[2, 7].SetStyle(st1);
            int i = 3;
            for (int j = 0; i < 3 + wms.Count; i++, j++)
            {
                var w = wms[j];
                cells[i, 0].PutValue(j + 1);// cells[i, 0].SetStyle(st1);
                cells[i, 1].PutValue(w.WorkerName); //cells[i, 1].SetStyle(st1
                cells[i, 2].PutValue(w.ResearchDay); //cells[i, 1].SetStyle(st1);
                cells[i, 3].PutValue(w.WorkDay); //cells[i, 1].SetStyle(st1);
                cells[i, 2 + 2].PutValue(w.BasePay_R); //cells[i, 2].SetStyle(st2);
                cells[i, 3 + 2].PutValue(w.BasePay_W); //cells[i, 3].SetStyle(st2);
                cells[i, 4 + 2].PutValue(w.BasePay); //cells[i, 4].SetStyle(st2);
                cells[i, 5 + 2].PutValue(w.Bonus_R);// cells[i, 5].SetStyle(st2);
                cells[i, 6 + 2].PutValue(w.Bonus_W); //cells[i, 6].SetStyle(st2);
                cells[i, 7 + 2].PutValue(w.Bonus);// cells[i, 7].SetStyle(st2);
            }
            cells.Merge(i, 0, 1, 2);
            cells[i, 0].PutValue("合计"); //cells[i, 0].SetStyle(st1);
            cells[i, 2].PutValue("——"); //cells[i, 1].SetStyle(st1
            cells[i, 3].PutValue("——"); //cells[i, 1].SetStyle(st1);
            cells[i, 2 + 2].PutValue(hj.BasePay_R); //cells[i, 2].SetStyle(st2);
            cells[i, 3 + 2].PutValue(hj.BasePay_W); //cells[i, 3].SetStyle(st2);
            cells[i, 4 + 2].PutValue(hj.BasePay); //cells[i, 4].SetStyle(st2);
            cells[i, 5 + 2].PutValue(hj.Bonus_R); //cells[i, 5].SetStyle(st2);
            cells[i, 6 + 2].PutValue(hj.Bonus_W); //cells[i, 6].SetStyle(st2);
            cells[i, 7 + 2].PutValue(hj.Bonus); //cells[i, 7].SetStyle(st2);

            for (int n = 0; n < 8 + 2; n++)
            {
                for (int m = 0; m < i + 1; m++)
                {
                    if (m >= 3 && n >= 2 + 2)
                        cells[m, n].SetStyle(st2);
                    else
                        cells[m, n].SetStyle(st1);

                }

            }
            sheet.AutoFitRows();
            sheet.AutoFitColumns();
            return workbook;
        }

        public ActionResult GetAll()
        {
            var baseFile = "D:\\高新企业";

            if (!Directory.Exists(baseFile))
            {
                Directory.CreateDirectory(baseFile);
            }

            using (var db = PCDbContext.NewDbContext)
            {
                var prjs = db.Project.ToList();
                for (var i = 201801; i < 201813; i++)
                {
                    var mName = string.Format("{0}年{1}月", (int)(i / 100), i % 100);
                    var nPath = string.Format("{0}\\{1}", baseFile, mName);
                    if (!Directory.Exists(nPath))
                    {
                        Directory.CreateDirectory(nPath);
                    }
                    foreach (var prj in prjs)
                    {
                        var filePath = string.Format("{0}\\{1}——科研、生产费用（{2}）.xls", nPath, prj.Name.Trim(), mName);
                        var wk = GetProjectMonthTable(prj.Id, i);
                        wk.Save(filePath);
                    }

                    var wkm = GetMonthTable(i);

                    wkm.Save(string.Format("{0}\\科研、生产费用合计（{1}）.xls", nPath, mName));
                }
            }

            return null;
        }

        */
        #endregion

        public Dictionary<string, object> GetMonthData(int month)
        {
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

select mt.workerid,wk.name workername,gzr workday,ky researchday,
ky*ws.basepay/gzr basepay_r,
ky*ws.bonus/gzr bonus_r,
ky*ws.AccumulationFund/gzr AccumulationFund_r,
ky*ws.SocialSecurity/gzr SocialSecurity_r,
ws.basepay,
ws.bonus,
ws.AccumulationFund,
ws.SocialSecurity from mt 
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
                    wm.AccumulationFund_W = wm.AccumulationFund - wm.AccumulationFund_R;
                    wm.SocialSecurity_W = wm.SocialSecurity - wm.SocialSecurity_R;
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
                AccumulationFund = wms.Sum(t => t.AccumulationFund),
                AccumulationFund_R = wms.Sum(t => t.AccumulationFund_R),
                AccumulationFund_W = wms.Sum(t => t.AccumulationFund_W),
                SocialSecurity = wms.Sum(t => t.SocialSecurity),
                SocialSecurity_R = wms.Sum(t => t.SocialSecurity_R),
                SocialSecurity_W = wms.Sum(t => t.SocialSecurity_W),

                WorkDay = wms.Sum(t => t.WorkDay),
                ResearchDay = wms.Sum(t => t.ResearchDay)
            };

            var dic = new Dictionary<string, object>();
            dic.Add("wms", wms);
            dic.Add("hj", hj);

            return dic;
        }

        public Workbook GetMonthWorkbook1(int month, Dictionary<string, object> dic)
        {
            var wms = dic["wms"] as List<WorkMonth3>;
            var hj = dic["hj"] as WorkMonth3;
            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = (int)(month / 100) + "年" + month % 100 + "月科研、生产费用（1）";

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
            cells.Merge(0, 0, 1, 4 + 3 * 3);
            cells[0, 0].PutValue((int)(month / 100) + "年" + month % 100 + "月科研、生产费用（1）");

            cells.Merge(1, 0, 2, 1);
            cells[1, 0].PutValue("序号");

            cells.Merge(1, 1, 2, 1);
            cells[1, 1].PutValue("姓名");

            cells.Merge(1, 2, 2, 1);
            cells[1, 2].PutValue("研发天数");
            cells.Merge(1, 3, 2, 1);
            cells[1, 3].PutValue("总工作日");

            cells.Merge(1, 4, 1, 3);
            cells[1, 1 + 3].PutValue("基本工资");

            cells.Merge(1, 4 + 3, 1, 3);
            cells[1, 4 + 3].PutValue("公积金");

            cells.Merge(1, 4 + 3 + 3, 1, 3);
            cells[1, 4 + 3 + 3].PutValue("社保");

            cells[2, 2 + 2].PutValue("研发费用");
            cells[2, 3 + 2].PutValue("生产费用");
            cells[2, 4 + 2].PutValue("合计");
            cells[2, 5 + 2].PutValue("研发费用");
            cells[2, 6 + 2].PutValue("生产费用");
            cells[2, 7 + 2].PutValue("合计");
            cells[2, 8 + 2].PutValue("研发费用");
            cells[2, 9 + 2].PutValue("生产费用");
            cells[2, 10 + 2].PutValue("合计");

            int i = 3;
            for (int j = 0; i < 3 + wms.Count; i++, j++)
            {
                var w = wms[j];
                cells[i, 0].PutValue(j + 1);
                cells[i, 1].PutValue(w.WorkerName);
                cells[i, 2].PutValue(w.ResearchDay);
                cells[i, 3].PutValue(w.WorkDay);
                cells[i, 2 + 2].PutValue(w.BasePay_R);
                cells[i, 3 + 2].PutValue(w.BasePay_W);
                cells[i, 4 + 2].PutValue(w.BasePay);
                cells[i, 5 + 2].PutValue(w.AccumulationFund_R);
                cells[i, 6 + 2].PutValue(w.AccumulationFund_W);
                cells[i, 7 + 2].PutValue(w.AccumulationFund);
                cells[i, 8 + 2].PutValue(w.SocialSecurity_R);
                cells[i, 9 + 2].PutValue(w.SocialSecurity_W);
                cells[i, 10 + 2].PutValue(w.SocialSecurity);
            }
            cells.Merge(i, 0, 1, 2);
            cells[i, 0].PutValue("合计");
            cells[i, 2].PutValue("——");
            cells[i, 3].PutValue("——");
            cells[i, 2 + 2].PutValue(hj.BasePay_R);
            cells[i, 3 + 2].PutValue(hj.BasePay_W);
            cells[i, 4 + 2].PutValue(hj.BasePay);
            cells[i, 5 + 2].PutValue(hj.AccumulationFund_R);
            cells[i, 6 + 2].PutValue(hj.AccumulationFund_W);
            cells[i, 7 + 2].PutValue(hj.AccumulationFund);
            cells[i, 8 + 2].PutValue(hj.SocialSecurity_R);
            cells[i, 9 + 2].PutValue(hj.SocialSecurity_W);
            cells[i, 10 + 2].PutValue(hj.SocialSecurity);

            for (int n = 0; n < 8 + 2 + 3; n++)
            {
                for (int m = 0; m < i + 1; m++)
                {
                    if (m >= 3 && n >= 2 + 2)
                        cells[m, n].SetStyle(st2);
                    else
                        cells[m, n].SetStyle(st1);

                }

            }
            sheet.AutoFitRows();
            sheet.AutoFitColumns();
            return workbook;
        }

        public Workbook GetMonthWorkbook2(int month, Dictionary<string, object> dic)
        {
            var wms = dic["wms"] as List<WorkMonth3>;
            var hj = dic["hj"] as WorkMonth3;
            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = (int)(month / 100) + "年" + month % 100 + "月科研、生产费用（2）";

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
            cells.Merge(0, 0, 1, 4 + 1 * 3);
            cells[0, 0].PutValue((int)(month / 100) + "年" + month % 100 + "月科研、生产费用（2）");

            cells.Merge(1, 0, 2, 1);
            cells[1, 0].PutValue("序号");

            cells.Merge(1, 1, 2, 1);
            cells[1, 1].PutValue("姓名");

            cells.Merge(1, 2, 2, 1);
            cells[1, 2].PutValue("研发天数");
            cells.Merge(1, 3, 2, 1);
            cells[1, 3].PutValue("总工作日");

            cells.Merge(1, 4, 1, 3);
            cells[1, 1 + 3].PutValue("绩效奖金");

            cells[2, 2 + 2].PutValue("研发费用");
            cells[2, 3 + 2].PutValue("生产费用");
            cells[2, 4 + 2].PutValue("合计");

            int i = 3;
            for (int j = 0; i < 3 + wms.Count; i++, j++)
            {
                var w = wms[j];
                cells[i, 0].PutValue(j + 1);
                cells[i, 1].PutValue(w.WorkerName);
                cells[i, 2].PutValue(w.ResearchDay);
                cells[i, 3].PutValue(w.WorkDay);
                cells[i, 2 + 2].PutValue(w.Bonus_R);
                cells[i, 3 + 2].PutValue(w.Bonus_W);
                cells[i, 4 + 2].PutValue(w.Bonus);
            }
            cells.Merge(i, 0, 1, 2);
            cells[i, 0].PutValue("合计");
            cells[i, 2].PutValue("——");
            cells[i, 3].PutValue("——");
            cells[i, 2 + 2].PutValue(hj.Bonus_R);
            cells[i, 3 + 2].PutValue(hj.Bonus_W);
            cells[i, 4 + 2].PutValue(hj.Bonus);

            for (int n = 0; n < 8 + 2 - 3; n++)
            {
                for (int m = 0; m < i + 1; m++)
                {
                    if (m >= 3 && n >= 2 + 2)
                        cells[m, n].SetStyle(st2);
                    else
                        cells[m, n].SetStyle(st1);

                }

            }
            sheet.AutoFitRows();
            sheet.AutoFitColumns();
            return workbook;
        }

        public Dictionary<string, object> GetProjectMonthData(string prjId, int month)
        {
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
select wt.workerid,date,WorkType from worktime wt 
where wt.month=@month and wt.ProjectId=@prjId
)
--人员本月科研情况
select t.date,t.workerid,t.workername,pw.[Index],
(case when wt.WorkType is null then t.type else '科研' end) worktype,
(case when wt1.WorkType is null then t.type else '科研' end) pworktype
 from (select dt.*,wk.id workerid,wk.name workername from dt,wk) t
left join wt on t.date=wt.date and t.workerid=wt.workerid
left join wt1 on t.date=wt1.date and t.workerid=wt1.workerid
left join Project_Worker pw on t.workerid=pw.workerid and pw.ProjectId=@prjId";
            var sqlWS = @"
select t.*,ws.BasePay,ws.bonus,ws.AccumulationFund,ws.SocialSecurity from (select distinct wk.Id workerid,wk.Name workername from Project_Worker pw left join worker wk on pw.WorkerId=wk.id where pw.ProjectId=@prjId) t 
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
                       group wt by new { wt.WorkerId, wt.WorkerName, wt.Index } into g
                       select new WorkMonth3
                       {
                           Index = g.Key.Index,
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
                            Index = wk.Index,
                            WorkerId = wk.WorkerId,
                            WorkerName = wk.WorkerName,
                            WorkTime = wk.WorkTime,
                            WorkDay = wk.WorkDay,
                            ResearchDay = wk.ResearchDay,
                            PResearchDay = wk.PResearchDay,
                            Bonus_R = wk.PResearchDay * (ws.Bonus / wk.WorkDay),
                            BasePay_R = wk.PResearchDay * (ws.BasePay / wk.WorkDay),
                            AccumulationFund_R = wk.PResearchDay * (ws.AccumulationFund / wk.WorkDay),
                            SocialSecurity_R = wk.PResearchDay * (ws.SocialSecurity / wk.WorkDay),
                        }).OrderBy(t => t.Index).ThenBy(t=>t.WorkerName).ToList();

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
                           AccumulationFund_R = g.Sum(t => t.AccumulationFund_R),
                           SocialSecurity_R = g.Sum(t => t.SocialSecurity_R),
                       }).FirstOrDefault();

            var dic = new Dictionary<string, object>();
            dic.Add("wts", wts);
            dic.Add("wss", wss);
            dic.Add("stwts", stwts);
            dic.Add("prj", prj);
            dic.Add("wkms", wkms);
            dic.Add("sum", sum);
            return dic;
        }

        public Workbook GetProjectMonthWorkbook1(int month, Dictionary<string, object> dic)
        {
            List<WorkTime3> wts = dic["wts"] as List<WorkTime3>;
            List<WorkerSalary3> wss = dic["wss"] as List<WorkerSalary3>;
            List<WorkTime3> stwts = dic["stwts"] as List<WorkTime3>;
            List<WorkMonth3> wkms = dic["wkms"] as List<WorkMonth3>;
            WorkMonth3 sum = dic["sum"] as WorkMonth3;
            Project prj = dic["prj"] as Project;

            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = (int)(month / 100) + "年" + month % 100 + "月项目考勤表（1）";

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
            cells.Merge(1, 2, 1, 20); cells[1, 2].PutValue($"{prj.Year}-{prj.SerialNumber}");
            cells.Merge(1, 2 + 20, 1, d + 5 - 20 - 3); cells[1, 2 + 20].PutValue((int)(month / 100) + "年" + month % 100 + "月");

            cells.Merge(2, 0, 2, 1); cells[2, 0].PutValue("序号");
            cells.Merge(2, 1, 2, 1); cells[2, 1].PutValue("姓名");
            cells.Merge(2, 2, 1, stwts.Count); cells[2, 2].PutValue("出    勤    情    况");
            cells.Merge(2, 2 + d, 2, 1); cells[2, 2 + d].PutValue("研发出勤\n（天）");
            cells.Merge(2, 2 + d + 1, 2, 1); cells[2, 2 + d + 1].PutValue("总工作日\n（天）");


            cells.Merge(0, d + 4, 2, 4); cells[0, d + 4].PutValue("科研费用");
            cells.Merge(2, d + 4, 2, 1); cells[2, d + 4].PutValue("基本工资");
            cells.Merge(2, d + 4 + 1, 2, 1); cells[2, d + 4 + 1].PutValue("公积金");
            cells.Merge(2, d + 4 + 2, 2, 1); cells[2, d + 4 + 2].PutValue("社保");
            cells.Merge(2, d + 4 + 3, 2, 1); cells[2, d + 4 + 3].PutValue("合计");

            for (int i = 0; i <= 1 + wkms.Count + 3; i++)
            {
                for (int j = 0; j <= 2 + d + 5; j++)
                {
                    cells[i, j].SetStyle(st1);
                }
            }

            for (int i = 4; i <= 1 + wkms.Count + 3; i++)
            {
                for (int j = d + 4; j <= d + 7; j++)
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
                    cells[r0, 2 + i].PutValue(wtm.PWorkType == "科研" ? "√" : "");
                }

                cells[r0, 2 + d].PutValue(wkm.PResearchDay);
                cells[r0, 2 + d + 1].PutValue(wkm.WorkDay);
                cells[r0, 2 + d + 2].PutValue(wkm.BasePay_R);
                cells[r0, 2 + d + 3].PutValue(wkm.AccumulationFund_R);
                cells[r0, 2 + d + 4].PutValue(wkm.SocialSecurity_R);
                cells[r0, 2 + d + 5].PutValue(wkm.BasePay_R + wkm.AccumulationFund_R + wkm.SocialSecurity_R);
            }

            int r = 3 + wkms.Count + 1;

            cells.Merge(r, 0, 1, 2); cells[r, 0].PutValue("合计");
            cells[r, 2 + d].PutValue(sum == null ? 0 : sum.PResearchDay);
            cells[r, 2 + d + 1].PutValue(sum == null ? 0 : sum.WorkDay);
            cells[r, 2 + d + 2].PutValue(sum == null ? 0 : sum.BasePay_R);
            cells[r, 2 + d + 3].PutValue(sum == null ? 0 : sum.AccumulationFund_R);
            cells[r, 2 + d + 4].PutValue(sum == null ? 0 : sum.SocialSecurity_R);
            cells[r, 2 + d + 5].PutValue((sum == null ? 0 : sum.BasePay_R) + (sum == null ? 0 : sum.AccumulationFund_R) + (sum == null ? 0 : sum.SocialSecurity_R));

            cells[r + 2, 2 + d - 20].PutValue("项目组长签字："); cells[r + 2, 2 + d - 20].SetStyle(st4);
            cells[r + 2, 2 + d - 5].PutValue("考勤员签字："); cells[r + 2, 2 + d - 5].SetStyle(st4);


            for (var i = 2; i < 2 + d; i++)
            {
                cells.SetColumnWidth(i, 3);
            }
            cells.SetColumnWidth(0, 5);
            cells.SetColumnWidth(1, 15);
            for (var i = 2 + d; i < 2 + d + 6; i++)
            {
                cells.SetColumnWidth(i, 12);
            }

            return workbook;
        }

        public Workbook GetProjectMonthWorkbook2(int month, Dictionary<string, object> dic)
        {
            List<WorkTime3> wts = dic["wts"] as List<WorkTime3>;
            List<WorkerSalary3> wss = dic["wss"] as List<WorkerSalary3>;
            List<WorkTime3> stwts = dic["stwts"] as List<WorkTime3>;
            List<WorkMonth3> wkms = dic["wkms"] as List<WorkMonth3>;
            WorkMonth3 sum = dic["sum"] as WorkMonth3;
            Project prj = dic["prj"] as Project;

            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            sheet.Name = (int)(month / 100) + "年" + month % 100 + "月项目考勤表（2）";

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
            cells.Merge(1, 2, 1, 20); cells[1, 2].PutValue($"{prj.Year}-{prj.SerialNumber}");
            cells.Merge(1, 2 + 20, 1, d + 5 - 20 - 3); cells[1, 2 + 20].PutValue((int)(month / 100) + "年" + month % 100 + "月");

            cells.Merge(2, 0, 2, 1); cells[2, 0].PutValue("序号");
            cells.Merge(2, 1, 2, 1); cells[2, 1].PutValue("姓名");
            cells.Merge(2, 2, 1, stwts.Count); cells[2, 2].PutValue("出    勤    情    况");
            cells.Merge(2, 2 + d, 2, 1); cells[2, 2 + d].PutValue("研发出勤\n（天）");
            cells.Merge(2, 2 + d + 1, 2, 1); cells[2, 2 + d + 1].PutValue("总工作日\n（天）");


            cells.Merge(0, d + 4, 2, 1); cells[0, d + 4].PutValue("科研费用");
            cells.Merge(2, d + 4, 2, 1); cells[2, d + 4].PutValue("绩效奖金");

            for (int i = 0; i <= 1 + wkms.Count + 3; i++)
            {
                for (int j = 0; j <= 2 + d + 2; j++)
                {
                    cells[i, j].SetStyle(st1);
                }
            }

            for (int i = 4; i <= 1 + wkms.Count + 3; i++)
            {
                for (int j = d + 4; j <= d + 4; j++)
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
                    cells[r0, 2 + i].PutValue(wtm.PWorkType == "科研" ? "√" : "");
                }

                cells[r0, 2 + d].PutValue(wkm.PResearchDay);
                cells[r0, 2 + d + 1].PutValue(wkm.WorkDay);
                cells[r0, 2 + d + 2].PutValue(wkm.Bonus_R);
            }

            int r = 3 + wkms.Count + 1;

            cells.Merge(r, 0, 1, 2); cells[r, 0].PutValue("合计");
            cells[r, 2 + d].PutValue(sum == null ? 0 : sum.PResearchDay);
            cells[r, 2 + d + 1].PutValue(sum == null ? 0 : sum.WorkDay);
            cells[r, 2 + d + 2].PutValue(sum == null ? 0 : sum.Bonus_R);

            cells[r + 2, 2 + d - 20].PutValue("项目组长签字："); cells[r + 2, 2 + d - 20].SetStyle(st4);
            cells[r + 2, 2 + d - 5].PutValue("考勤员签字："); cells[r + 2, 2 + d - 5].SetStyle(st4);


            for (var i = 2; i < 2 + d; i++)
            {
                cells.SetColumnWidth(i, 3);
            }
            cells.SetColumnWidth(0, 5);
            cells.SetColumnWidth(1, 15);
            for (var i = 2 + d; i < 2 + d + 3; i++)
            {
                cells.SetColumnWidth(i, 12);
            }

            return workbook;
        }

        /// <summary>
        /// 批量导出  新
        /// </summary>
        /// <returns></returns>
        public ActionResult GetAll2()
        {
            var baseFile = "D:\\高新企业";

            if (!Directory.Exists(baseFile))
            {
                Directory.CreateDirectory(baseFile);
            }

            using (var db = PCDbContext.NewDbContext)
            {
                var prjs = db.Project.ToList();
                for (var i = 201901; i < 201902; i++)
                {
                    var mName = string.Format("{0}年{1}月", (int)(i / 100), i % 100);
                    var nPath = string.Format("{0}\\{1}", baseFile, mName);
                    if (!Directory.Exists(nPath))
                    {
                        Directory.CreateDirectory(nPath);
                    }
                    foreach (var prj in prjs)
                    {
                        var filePath1 = string.Format("{0}\\{1}——科研、生产费用（{2}-1）.xls", nPath, prj.Name.Trim(), mName);
                        var filePath2 = string.Format("{0}\\{1}——科研、生产费用（{2}-2）.xls", nPath, prj.Name.Trim(), mName);
                        var data = GetProjectMonthData(prj.Id, i);
                        var wk1 = GetProjectMonthWorkbook1(i, data);
                        var wk2 = GetProjectMonthWorkbook2(i, data);
                        wk1.Save(filePath1);
                        wk2.Save(filePath2);
                    }

                    var data2 = GetMonthData(i);
                    var wkm1 = GetMonthWorkbook1(i, data2);
                    var wkm2 = GetMonthWorkbook2(i, data2);
                    wkm1.Save(string.Format("{0}\\科研、生产费用合计（{1}-1）.xls", nPath, mName));
                    wkm2.Save(string.Format("{0}\\科研、生产费用合计（{1}-2）.xls", nPath, mName));
                }
            }

            return null;
        }

        public ActionResult Test()
        {
            string path = @"D:\test.xls";
            Workbook wb = new Workbook(path);
            Worksheet ws = wb.Worksheets[0];

            ws.Cells[20, 1].PutValue("hello world!");
            wb.Save(path);
            return null;
        }
    }
}