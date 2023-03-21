using log4net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logger
{
    public static class MyLogger
    {
        private static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static System.Object locker = new object();

        public static void log(string msg, string logtype)
        {
            //lock (locker)
            {
                var stackFrames = new StackTrace().GetFrames();
                var callingframe = stackFrames.ElementAt(2);
                var method = callingframe.GetMethod();
                //var fn = callingframe.GetFileName() + ":" + callingframe.GetFileLineNumber().ToString();
                //fn = callingframe.ToString();
                //fn = new StackTrace().ToString();
                //fn = fn.Trim();
                //string[] locations = fn.Split(new string[] { "at " }, StringSplitOptions.RemoveEmptyEntries);
                //fn = locations[2].Trim();
                //var p1 = fn.IndexOf('(');
                //var name = fn.Substring(0, p1);
                var name = method.Name;
                ThreadContext.Properties["method"] = name;

                if (logtype == "error")
                    ThreadContext.Properties["status"] = "E";
                else if (logtype == "warn")
                    ThreadContext.Properties["status"] = "W";
                else
                    ThreadContext.Properties["status"] = "I";

                //if (logtype == "error")
                //    logger.Error(msg);
                //else if (logtype == "warn")
                //    logger.Warn(msg);
                //else
                //    logger.Info(msg);
                logger.Info(msg);
            }
        }

        public static void Info(string msg)
        {
            log(msg, "info");
        }

        public static void Error(string msg)
        {
            log(msg, "error");
        }
        public static void Warn(string msg)
        {
            log(msg, "warn");
        }
    }
}
