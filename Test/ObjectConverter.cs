using log4net.Core;
using log4net.Layout;
using log4net.Layout.Pattern;
using log4net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    /// <summary>
    /// 根据键值获取值的对象
    /// </summary>
    public interface IGetObjectValueByKey
    {
        string GetByKey(string name);
    }

    /// <summary>
    /// 自定义的LayoutConverter
    /// </summary>
    /// <remarks>用于PatternLayout</remarks>
    public class ObjectConverter : PatternLayoutConverter
    {
        static Func<object, string, object> funcs;
        static ObjectConverter()
        {
            //********根据键值获取值的顺序
            //从接口获取值
            funcs += GetValueByInterface;
            //反射获取属性值
            funcs += GetValueByReflection;
            //从索引值获取值
            funcs += GetValueByIndexer;
        }

        /// <summary>
        /// 实现PatternLayoutConverter.Convert抽象方法
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="loggingEvent"></param>
        protected override void Convert(TextWriter writer, LoggingEvent loggingEvent)
        {
            //获取传入的消息对象
            object objMsg = loggingEvent.MessageObject;

            if (objMsg == null)
            {
                //如果对象为空输出log4net默认的null字符串
                writer.Write(SystemInfo.NullText);
                return;
            }
            if (string.IsNullOrEmpty(this.Option))
            {
                //如果属性为空，输出消息对象的ToString()
                writer.Write(objMsg.ToString());
                return;
            }

            // 获取属性并输出
            object val = GetValue(funcs, objMsg, Option);
            writer.Write(val == null ? string.Empty : val.ToString());
        }

        #region 静态方法
        /// <summary>
        /// 循环方法列表，根据键值获取值
        /// </summary>
        /// <param name="func">方法列表委托</param>
        /// <param name="obj">对象</param>
        /// <param name="name">键值</param>
        /// <returns></returns>
        private static object GetValue(Func<object, string, object> func, object obj, string name)
        {
            object val = null;
            if (func != null)
            {
                foreach (Func<object, string, object> del in func.GetInvocationList())
                {
                    val = del(obj, name);
                    //如果获取的值不为null，则跳出循环
                    if (val != null)
                    {
                        break;
                    }
                }
            }
            return val;
        }

        /// <summary>
        /// 使用接口方式取值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <remarks>效率最高，避免了反射带来的效能损耗</remarks>
        private static object GetValueByInterface(object obj, string name)
        {
            object val = null;
            IGetObjectValueByKey objConverter = obj as IGetObjectValueByKey;
            if (objConverter != null)
            {
                val = objConverter.GetByKey(name);
            }
            return val;
        }

        /// <summary>
        /// 反射对象的获取属性，获取属性值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        private static object GetValueByReflection(object obj, string name)
        {
            object val = null;
            Type t = obj.GetType();
            var propertyInfo = t.GetProperty(name);
            if (propertyInfo != null)
            {
                val = propertyInfo.GetValue(obj, null);
            }

            return val;
        }

        /// <summary>
        /// 反射对象的索引器，获取值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        private static object GetValueByIndexer(object obj, string name)
        {
            object val = null;

            MethodInfo getValueMethod = obj.GetType().GetMethod("get_Item");
            if (getValueMethod != null)
            {
                val = getValueMethod.Invoke(obj, new object[] { name });
            }

            return val; 
        }
        #endregion
    }

    /// <summary>
    /// 自定义的Object布局
    /// </summary>
    public class ObjectPatternLayout : PatternLayout
    {
        public ObjectPatternLayout()
        {
            // 添加名为%o的Converter
            this.AddConverter("o", typeof(ObjectConverter));
        }
    }
}
