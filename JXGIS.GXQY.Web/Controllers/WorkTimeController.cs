using JXGIS.GXQY.Web.Base;
using JXGIS.GXQY.Web.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Web.Mvc;

namespace JXGIS.GXQY.Web.Controllers
{
    public class WorkTimeController : Controller
    {
        public static int start = 21;
        public static int end = 20;


        public class Series
        {
            public int Date { get; set; }

            public int Date2
            {
                get
                {
                    return Date % 100;
                }
            }

            public int Month { get; set; }

            public string Type { get; set; }

            public string ProjectId { get; set; }
        }

        #region 废弃

        public class WorkerMonth
        {
            public int Month { get; set; }
            public string WorkerId { get; set; }

            public string WorkerName { get; set; }

            public double BasePay { get; set; }

            public double Bonus { get; set; }

            public List<Series> WorkTime { get; set; }

            public List<Project> Projects { get; set; }

            /// <summary>
            /// 基本工资生产费用
            /// </summary>
            public double BasePay_W
            {
                get
                {
                    var wkd = Workday;
                    var rsd = ResearchDay;
                    return Math.Round(wkd * BasePay / (wkd + rsd), 2, MidpointRounding.AwayFromZero);
                }
            }

            /// <summary>
            /// 基本工资研发费用
            /// </summary>
            public double BasePay_R
            {
                get
                {
                    return BasePay - BasePay_W;
                }
            }


            /// <summary>
            /// 绩效工资生产费用
            /// </summary>
            public double Bonus_W
            {
                get
                {
                    var wkd = Workday;
                    var rsd = ResearchDay;
                    return Math.Round(wkd * Bonus / (wkd + rsd), 2, MidpointRounding.AwayFromZero);
                }
            }


            /// <summary>
            /// 绩效工资研发费用
            /// </summary>
            public double Bonus_R
            {
                get
                {
                    return Bonus - Bonus_W;
                }
            }


            public int Workday
            {
                get
                {
                    return WorkTime == null ? 0 : WorkTime.Where(w => w.Type == "工作日").Count();
                }
            }


            public int Holiday
            {
                get
                {
                    return WorkTime == null ? 0 : WorkTime.Where(w => w.Type == "节假日").Count();
                }
            }

            public int ResearchDay
            {
                get
                {
                    return WorkTime == null ? 0 : WorkTime.Where(w => w.Type != "工作日" && w.Type != "节假日").Count();
                }
            }

        }
        #endregion
        public class WorkTime2
        {
            public string WorkerId { get; set; }

            public string WorkerName { get; set; }

            public int Date { get; set; }

            public int Date2 { get; set; }

            public int Month { get; set; }

            public string Type { get; set; }

            public string ProjectId { get; set; }
        }
        public class WorkerMonth2
        {
            public string ProjectId { get; set; }
            public int Month { get; set; }
            public string WorkerId { get; set; }

            public string WorkerName { get; set; }

            public int? Index { get; set; }

            public string ProjectRole { get; set; }

            public double? BasePay { get; set; }

            public double? Bonus { get; set; }

            public double? AccumulationFund { get; set; }

            public double? SocialSecurity { get; set; }

            public List<WorkTime2> WorkTime { get; set; }

            public List<Project2> Projects { get; set; }

            /// <summary>
            /// 基本工资生产费用
            /// </summary>
            public double? BasePay_W { get; set; }

            /// <summary>
            /// 基本工资研发费用
            /// </summary>
            public double? BasePay_R { get; set; }

            public double? BasePay_AVG { get; set; }

            /// <summary>
            /// 绩效工资生产费用
            /// </summary>
            public double? Bonus_W { get; set; }


            /// <summary>
            /// 绩效工资研发费用
            /// </summary>
            public double? Bonus_R { get; set; }

            public double? Bonus_AVG { get; set; }

            public double? AccumulationFund_W { get; set; }

            public double? AccumulationFund_R { get; set; }

            public double? AccumulationFund_AVG { get; set; }

            public double? SocialSecurity_W { get; set; }

            public double? SocialSecurity_R { get; set; }

            public double? SocialSecurity_AVG { get; set; }

            /// <summary>
            /// 生产天数
            /// </summary>
            public int Workday { get; set; }

            /// <summary>
            /// 节假日天数
            /// </summary>
            public int Holiday { get; set; }

            /// <summary>
            /// 科研天数
            /// </summary>
            public int ResearchDay { get; set; }

            /// <summary>
            /// 总天数
            /// </summary>
            public int TotalDay { get; set; }

            /// <summary>
            /// 总工作日
            /// </summary>
            public int TotalWorkDay { get; set; }

            /// <summary>
            /// 本项目的研究天数
            /// </summary>
            public int? PResearchDay { get; set; }

            public double? PBonus { get; set; }

            public double? PBasePay { get; set; }

            public double? PAccumulationFund { get; set; }

            public double? PSocialSecurity { get; set; }

        }


        public class Project2
        {
            public string WorkerId { get; set; }

            public string ProjectId { get; set; }

            public string ProjectName { get; set; }

            public DateTime? StartTime { get; set; }
            public DateTime? EndTime { get; set; }

        }
        #region 废弃
        public ActionResult GetWorkTime(string prjId, int year, int month)
        {
            string s = null;
            try
            {
                var bMonth = year * 100 + month;

                using (var db = PCDbContext.NewDbContext)
                {
                    var projectTimes = db.Database.SqlQuery<WorkTime>("select wk.* from project p inner join project_worker r on p.id=r.projectid inner join worker w on r.workerid=w.id inner join worktime wk on w.id=wk.workerid where p.id=@projectId", new SqlParameter("@projectId", prjId)).ToList();
                    var dateSeries = db.DateType.Where(t => t.Month == bMonth).ToList();
                    var workerSalary = db.Database.SqlQuery<WorkerSalary>("select w.* from project_worker as t left join workerSalary as w on t.workerid=w.workerid where t.projectid=@projectid and w.month=@month",
                        new SqlParameter("@projectid", prjId),
                        new SqlParameter("@month", bMonth)
                        ).ToList();


                    var projects = db.Project.Include("Workers").ToList();
                    var prj = projects.Where(p => p.Id == prjId).FirstOrDefault();
                    var workers = prj.Workers;

                    var wMonth = new List<WorkerMonth>();
                    foreach (var wk in workers)
                    {
                        var wt = new List<Series>();
                        var tms = projectTimes.Where(w => w.WorkerId == wk.Id).ToList();

                        foreach (var d in dateSeries)
                        {
                            var tm = tms.Where(t => t.Date == d.Date).FirstOrDefault();
                            wt.Add(new Series
                            {
                                Date = d.Date,
                                Month = d.Month,
                                Type = tm == null ? d.Type : tm.WorkType,
                                ProjectId = tm == null ? null : tm.ProjectId
                            });
                        }
                        var ws = workerSalary.Where(w => w.WorkerId == wk.Id).FirstOrDefault();
                        wMonth.Add(new WorkerMonth
                        {
                            WorkerId = wk.Id,
                            WorkerName = wk.Name,
                            Month = bMonth,
                            BasePay = ws != null ? ws.BasePay : 0,
                            Bonus = ws != null ? ws.Bonus : 0,
                            WorkTime = wt.OrderBy(t => t.Date).ToList(),
                            Projects = (from p in wk.Projects select new Project { Id = p.Id, Name = p.Name }).ToList()
                        });
                    }

                    var woMonth = new List<object>();
                    woMonth.AddRange(wMonth);
                    woMonth.Add(new
                    {
                        Id = "HJ",
                        WorkerName = "合计",
                        BasePay = wMonth.Sum(w => w.BasePay),
                        Bonus = wMonth.Sum(w => w.Bonus),
                        BasePay_W = wMonth.Sum(w => w.BasePay_W),
                        Bonus_W = wMonth.Sum(w => w.Bonus_W),
                        BasePay_R = wMonth.Sum(w => w.BasePay_R),
                        Bonus_R = wMonth.Sum(w => w.Bonus_R),
                    });

                    RtObj.Serialize(woMonth, out s);
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }

            return Content(s);
        }
        #endregion


        public ActionResult GetWorkTime2(string prjId, int year, int month)
        {
            string s = null;
            try
            {
                var bMonth = year * 100 + month;

                var sqlProject = @"with dt as(
select dt.date,dt.month,dt.type,a.workerid from datetype as dt,(select distinct workerid from project_worker where ProjectId=@projectId) as a),
wkt as (
select dt.*,wk.ProjectId,wk.WorkType from dt left join worktime as wk on dt.date=wk.date and wk.WorkerId=dt.WorkerId),
yf as (
select month,workerid,sum(case
when worktype is not null then 1
else 0 end ) ResearchDay,
sum(case
when type='节假日'  then 1
else 0 end ) Holiday,
sum(case
when worktype is null and type<>'节假日'  then 1
else 0 end ) Workday,
sum(case
when worktype is not null or type='工作日'  then 1
else 0 end ) TotalWorkday,
count(1) Totalday from wkt group by month ,workerid
),
yf_w as (
select yf.*,ws.basepay,ws.bonus,ws.AccumulationFund,ws.SocialSecurity,
ResearchDay*ws.basepay/TotalWorkday basepay_r,ws.basepay/TotalWorkday basepay_avg,
ResearchDay*ws.bonus/TotalWorkday bonus_r ,ws.bonus/TotalWorkday bonus_avg,
ResearchDay*ws.AccumulationFund/TotalWorkday AccumulationFund_r,ws.AccumulationFund/TotalWorkday AccumulationFund_avg,
ResearchDay*ws.SocialSecurity/TotalWorkday SocialSecurity_r,ws.SocialSecurity/TotalWorkday SocialSecurity_avg,
basepay-ResearchDay*ws.basepay/TotalWorkday basepay_w,
bonus- ResearchDay*ws.bonus/TotalWorkday bonus_w,
AccumulationFund- ResearchDay*ws.AccumulationFund/TotalWorkday AccumulationFund_w,
SocialSecurity- ResearchDay*ws.SocialSecurity/TotalWorkday SocialSecurity_w
from yf 
left join WorkerSalary ws on yf.month=ws.Month and yf.WorkerId=ws.WorkerId)
select pw.[Index],pw.projectrole,yf_w.*,wk.Name workerName,x.workdates PResearchDay,x.workdates*bonus_avg PBonus,x.workdates*basepay_avg PBasePay,x.workdates*AccumulationFund_avg PAccumulationFund,x.workdates*SocialSecurity_avg PSocialSecurity from yf_w left join worker wk on wk.Id = yf_w.workerid
left join (select wt.WorkerId,wt.month,count(1) workdates from WorkTime wt where wt.ProjectId=@projectId2
group by wt.WorkerId,wt.month) x on x.WorkerId=yf_w.WorkerId and x.month=yf_w.month
left join (select * from Project_Worker where ProjectId=@projectId3 ) pw on yf_w.WorkerId=pw.WorkerId;";
                var sqlWorkTime = @"with w as (select * from (select date,month,type from datetype dt where dt.month=@month)dt,
(select distinct WorkerId from Project_Worker pw where pw.ProjectId=@projectId) wkids)

select w.WorkerId,w.date,w.date%100 date2,w.month,
(case when wt.worktype is not null then wt.ProjectId else null end ) projectid
,(case when wt.worktype is not null then wt.WorkType else w.type end ) type from w left join  WorkTime wt on w.date=wt.Date and wt.WorkerId=w.WorkerId
";
                var sqlPW = @"with wks as (select distinct workerid from project p left join Project_Worker pw on p.id=pw.ProjectId where id=@projectId)
select p.Id projectid,p.Name projectname,wks.WorkerId,p.StartTime,p.EndTime from Project_Worker pw 
inner join Project p on pw.ProjectId=p.id 
inner join wks on wks.WorkerId=pw.WorkerId where p.EndTime is null or p.EndTime>convert(datetime,@date)";

                /*
                --写法二 
                var sqlPW = @"select distinct p.id,p.name,pw1.WorkerId from Project p
                inner join Project_Worker pw1 on p.id = pw1.ProjectId
inner join Project_Worker pw2 on pw1.WorkerId = pw2.WorkerId
where pw2.ProjectId = 'e74f39aa-43be-4dfb-b6bc-5ab27f6932ee' and(p.EndTime is null or p.EndTime > convert(datetime, '20190201'))";*/

                using (var db = PCDbContext.NewDbContext)
                {
                    var wms = db.Database.SqlQuery<WorkerMonth2>(sqlProject, new SqlParameter("@projectId", prjId), new SqlParameter("@projectId2", prjId), new SqlParameter("@projectId3", prjId)).ToList();
                    var wms_month = wms.Where(w => w.Month == bMonth).ToList();
                    var wkts = db.Database.SqlQuery<WorkTime2>(sqlWorkTime, new SqlParameter("@projectId", prjId), new SqlParameter("@month", bMonth)).ToList();
                    var prjs = db.Database.SqlQuery<Project2>(sqlPW, new SqlParameter("@projectId", prjId), new SqlParameter("@date", bMonth + "01")).ToList();

                    foreach (var w in wms_month)
                    {
                        w.WorkTime = wkts.Where(x => x.WorkerId == w.WorkerId).OrderBy(x => x.Date).ToList();
                        w.Projects = prjs.Where(p => p.WorkerId == w.WorkerId).ToList();
                        w.ProjectId = prjId;
                    }

                    wms_month = wms_month.OrderBy(w => w.Index).ThenBy(w => w.WorkerName).ToList();

                    wms_month.Add(new WorkerMonth2()
                    {
                        WorkerId = "HJ",
                        WorkerName = "合计",
                        ResearchDay = wms_month.Sum(w => w.ResearchDay),
                        PResearchDay = wms_month.Sum(w => w.PResearchDay ?? 0),
                        BasePay = wms_month.Sum(w => w.BasePay),
                        Bonus = wms_month.Sum(w => w.Bonus),
                        AccumulationFund = wms_month.Sum(w => w.AccumulationFund),
                        SocialSecurity = wms_month.Sum(w => w.SocialSecurity),

                        BasePay_W = wms_month.Sum(w => w.BasePay_W),
                        Bonus_W = wms_month.Sum(w => w.Bonus_W),
                        AccumulationFund_W = wms_month.Sum(w => w.AccumulationFund_W),
                        SocialSecurity_W = wms_month.Sum(w => w.SocialSecurity_W),

                        BasePay_R = wms_month.Sum(w => w.BasePay_R),
                        Bonus_R = wms_month.Sum(w => w.Bonus_R),
                        AccumulationFund_R = wms_month.Sum(w => w.AccumulationFund_R),
                        SocialSecurity_R = wms_month.Sum(w => w.SocialSecurity_R),

                        PBasePay = wms_month.Sum(w => w.BasePay_AVG * (w.PResearchDay ?? 0)),
                        PBonus = wms_month.Sum(w => w.Bonus_AVG * (w.PResearchDay ?? 0)),
                        PAccumulationFund = wms_month.Sum(w => w.AccumulationFund_AVG * (w.PResearchDay ?? 0)),
                        PSocialSecurity = wms_month.Sum(w => w.SocialSecurity_AVG * (w.PResearchDay ?? 0)),
                        Workday = wms_month.Sum(w => w.Workday),
                    });

                    var prj = (from w in wms
                               group w by w.Month into g
                               select new
                               {
                                   Month = g.Key,
                                   BasePay_R = g.Sum(t => (t.BasePay_AVG ?? 0) * (t.PResearchDay ?? 0)),
                                   Bonus_R = g.Sum(t => (t.Bonus_AVG ?? 0) * (t.PResearchDay ?? 0)),
                                   AccumulationFund_R = g.Sum(t => (t.AccumulationFund_AVG ?? 0) * (t.PResearchDay ?? 0)),
                                   SocialSecurity_R = g.Sum(t => (t.SocialSecurity_AVG ?? 0) * (t.PResearchDay ?? 0))
                               }).Where(x => x.BasePay_R != 0 && x.Bonus_R != 0 && x.AccumulationFund_R != 0 && x.SocialSecurity_R != 0).OrderBy(x => x.Month).ToList();

                    RtObj ro = new RtObj();
                    ro.Add("WorkMonth", wms_month);
                    ro.Add("Project", prj);
                    RtObj.Serialize(ro.Data, out s, new Newtonsoft.Json.Converters.IsoDateTimeConverter() { DateTimeFormat = "yyyy年MM月dd日" });
                }

            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }

            return Content(s);

        }

        public ActionResult AddWorkTime(string prjId, string wkId, int date)
        {
            string s = null;
            try
            {
                using (var ts = new TransactionScope())
                {
                    using (var db = PCDbContext.NewDbContext)
                    {
                        var wtm = db.WorkTime.Where(wt => wt.Date == date && wt.WorkerId == wkId).FirstOrDefault();
                        if (string.IsNullOrEmpty(prjId))
                        {
                            db.WorkTime.Remove(wtm);
                        }
                        else
                        {
                            if (wtm != null)
                            {
                                wtm.ProjectId = prjId;
                                wtm.WorkType = db.Project.Find(prjId).Name;
                            }
                            else
                            {
                                var prj = db.Project.Find(prjId);
                                db.WorkTime.Add(new WorkTime()
                                {
                                    Id = Guid.NewGuid().ToString(),
                                    Date = date,
                                    Month = db.DateType.Where(x => x.Date == date).FirstOrDefault().Month,
                                    ProjectId = prjId,
                                    WorkerId = wkId,
                                    WorkType = prj != null ? prj.Name : null
                                });
                            }
                        }
                        db.SaveChanges();
                        ts.Complete();
                        RtObj.Serialize(string.Empty, out s);
                    }
                }
            }
            catch (Exception ex)
            {
                RtObj.Serialize(ex, out s);
            }

            return Content(s);
        }
    }
}