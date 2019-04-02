using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace JXGIS.GXQY.Web.Base
{
    public class RtObj
    {
        public RtObj()
        {

        }

        public RtObj(string erMessage)
        {
            this.ErrorMessage = erMessage;
        }

        public RtObj(object data)
        {
            this.Data = data;
        }

        public RtObj(Exception ex)
        {
            this.ErrorMessage = ex.Message;
        }

        public string ErrorMessage { get; set; }

        public object Data { get; set; } = new Dictionary<string, object>();

        public void Add(string key, object value)
        {
            if (!(this.Data is Dictionary<string, object>))
            {
                this.Data = new Dictionary<string, object>();
            }

            (this.Data as Dictionary<string, object>).Add(key, value);
        }

        public string Serialize(params JsonConverter[] converters)
        {
            return JsonConvert.SerializeObject(this, new JsonSerializerSettings { ReferenceLoopHandling = ReferenceLoopHandling.Ignore, Converters = converters });
        }

        public static string Serialize(string error, out string s, params JsonConverter[] converters)
        {
            s = new RtObj(error).Serialize(converters);
            return s;
        }

        public static string Serialize(Exception er, out string s, params JsonConverter[] converters)
        {
            s = new RtObj(er).Serialize(converters);
            return s;
        }

        public static string Serialize(object obj, out string s, params JsonConverter[] converters)
        {
            s = new RtObj(obj).Serialize(converters);
            return s;
        }
    }
}