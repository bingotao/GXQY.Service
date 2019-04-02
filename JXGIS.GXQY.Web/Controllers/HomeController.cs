using JXGIS.GXQY.Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;

namespace JXGIS.GXQY.Web.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult GetProjects()
        {
            string s = null;
            using (var db = PCDbContext.NewDbContext)
            {
                var projects = db.Project.OrderByDescending(p => p.StartTime).ToList();
                s = Newtonsoft.Json.JsonConvert.SerializeObject(projects, new Newtonsoft.Json.JsonSerializerSettings { ReferenceLoopHandling = ReferenceLoopHandling.Ignore });
            }
            return Content(s);
        }
    }
}