using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ReportStat
{
    public class XStatHelper
    {
        private static object m_lock_report_list = new object();
        private static List<KReportBase> m_report_list = new List<KReportBase>();
        private GReport4Retailers report_4retailers;
        public XStatHelper()
        {
            report_4retailers = new GReport4Retailers();
        }
        static public void update_report(KReportBase report)
        {
            lock (m_lock_report_list)
            {
                m_report_list.Add(report);
            }
        }
        private KReportBase take_report()
        {
            KReportBase report = null;
            lock (m_lock_report_list)
            {
                if (m_report_list.Count > 0)
                {
                    report = m_report_list[0];
                    m_report_list.RemoveAt(0);
                }
            }
            return report;
        }
        public void start_report()
        {
            lock (m_lock_report_list)
            {
                m_report_list.Clear();
            }
            Task.Run(() =>
            {
                MyLogger.Info("...start report");

                DateTime last_confirmed_time = DateTime.MinValue;

                while (!Program.g_must_end)
                {
                    try
                    {
                        KReportBase report = take_report();

                        bool must_track = false;
                        if (last_confirmed_time == DateTime.MinValue)
                        {
                            must_track = true;
                        }
                        else
                        {
                            TimeSpan time_diff = DateTime.Now - last_confirmed_time;
                            if (time_diff.TotalHours >= 1)
                                must_track = true;
                        }

                        if (report != null)
                        {
                            if (report.is_4_retailers())
                            {
                                report_4retailers.report(report);

                                lock (m_lock_report_list)
                                {
                                    MyLogger.Info($"remained num to report to google sheets = {m_report_list.Count}");
                                }
                            }
                        }
                        else
                        {
                            Thread.Sleep(1000);
                        }

                        if (must_track)
                        {
                            report_4retailers.monitor();
                            last_confirmed_time = DateTime.Now;
                        }
                    }
                    catch (Exception exception)
                    {
                        MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                    }
                }
                MyLogger.Info("...stop report");
            });
        }
    }
}
