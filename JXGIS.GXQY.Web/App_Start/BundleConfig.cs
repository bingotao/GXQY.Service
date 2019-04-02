using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Optimization;

namespace JXGIS.GXQY.Web
{
    public class BundleConfig
    {
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/home_index_js").IncludeDirectory("~/Views/Home/js", "*.js"));
            bundles.Add(new StyleBundle("~/bundles/home_index_css").IncludeDirectory("~/Views/Home/css", "*.css"));
            bundles.Add(new StyleBundle("~/bundles/com_js").IncludeDirectory("~/Refers/extends", "*.js"));
        }
    }
}