using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportStat
{
    public class KReportJsonHelper
    {
        public List<object> report_list;
        public KReportJsonHelper()
        {
            report_list = new List<object>();
        }
        public void load_from_json_text(string json_text)
        {
            try
            {
                report_list.Clear();

                bool is_start = false;
                List<string> report_lines = new List<string>();

                string[] lines = json_text.Replace("\r", "").Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];

                    if (line == "[")
                    {
                        report_lines.Clear();

                        is_start = true;
                        continue;
                    }
                    if (line == "{")
                    {
                        report_lines.Clear();

                        is_start = true;
                        report_lines.Add(line);
                        continue;
                    }
                    if (!is_start)
                        continue;
                    if (line == "]," || line == "},")
                    {
                        is_start = false;

                        if (line == "},")
                            report_lines.Add("}");

                        string report_json_text = "";
                        for (int k = 0; k < report_lines.Count; k++)
                            report_json_text += (report_json_text == "") ? report_lines[k] : "\n" + report_lines[k];

                        object report =  KReportBase.generate_report_from_json(report_json_text);
                        if (report != null)
                            report_list.Add(report);

                        continue;
                    }
                    report_lines.Add(line);
                }

                File.Delete("cr2_2.txt");
                for (int i = 0; i < report_list.Count; i++)
                {
                    KReportBase report = (KReportBase)report_list[i];

                    if (report.m_mail_type == KReportBase.MailType.CR_2_2)
                    {
                        if (report.m_giftcard_details.Count(s => s.m_pin == "") > 0)
                        {
                            File.AppendAllText("cr2_2.txt", $"{report.m_mail_id}" + "\n");
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
    }
}
