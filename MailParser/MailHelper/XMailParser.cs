using MailParser;
using Logger;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailHelper
{
    public class XMailParser
    {
        private int m_parser_thread_num = 1;
        private static object m_lock_unchecked_ids = new object();
        private static List<int> m_unchecked_list = new List<int>();

        private List<KMailBaseParser> m_mail_parsers = new List<KMailBaseParser>();

        public XMailParser()
        {
            m_mail_parsers.Add(new KMailCR1());
            m_mail_parsers.Add(new KMailCR2());
            m_mail_parsers.Add(new KMailCR3());
            m_mail_parsers.Add(new KMailCR4());
            m_mail_parsers.Add(new KMailCR7());
            m_mail_parsers.Add(new KMailCR8());
            m_mail_parsers.Add(new KMailBaseOP());
            m_mail_parsers.Add(new KMailBaseSC());
            m_mail_parsers.Add(new KMailBaseCC());
        }
        static public void add_unchecked_mail_id(int id)
        {
            lock (m_lock_unchecked_ids)
            {
                m_unchecked_list.Add(id);
            }
        }
        public void start_parse_mails()
        {
            int i, k;

            MyLogger.Info($"Start parser : thread num = {m_parser_thread_num}....");

            m_unchecked_list = new List<int>();

            get_unchecked_mails_ids_from_db();

            for (i = 0; i < m_parser_thread_num; i++)
            {
                new Thread(() =>
                {
                    MyLogger.Info($"...start parsing thread. {i} th thread ID = {Thread.CurrentThread.ManagedThreadId}");

                    while (!Program.g_must_end)
                    {
                        try
                        {
                            int id = take_unchecked_mails_id();
                            if (id == -1)
                            {
                                MyLogger.Info("...parsing : no unchecked mail. will wait.");
                                for (k = 0; !Program.g_must_end && k < 600; k++)
                                    Thread.Sleep(100);
                                continue;
                            }

                            string eml_file = get_eml_file_path_from_id(id);

                            for (k = 0; k < m_mail_parsers.Count; k++)
                            {
                                if (m_mail_parsers[k].parse(id, eml_file))
                                    break;
                            }
                            if (k == m_mail_parsers.Count)
                            {
                                MyLogger.Error($"***Parsing Failed*** : mail.id = {id}");

                                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_UPDATE_DB))
                                    Program.g_db.set_mail_checked_flag(id, ConstEnv.MAIL_PARSING_FAILED);
                            }
                        }
                        catch (Exception exception)
                        {
                            MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                        }
                    }
                    MyLogger.Info($"...stop parsing thread. {i} th thread ID = {Thread.CurrentThread.ManagedThreadId}");

                }).Start();
            }
        }
        public void parse_specific_mails(int[] mail_ids)
        {
            int i, k;

            MyLogger.Info($"[TEST] Start parser : thread num = {m_parser_thread_num}....");

            m_unchecked_list = new List<int>();

            foreach (int mail_id in mail_ids)
                m_unchecked_list.Add(mail_id);

            for (i = 0; i < m_parser_thread_num; i++)
            {
                new Thread(() =>
                {
                    MyLogger.Info($"[TEST] ...start parsing thread. {i} th thread ID = {Thread.CurrentThread.ManagedThreadId}");

                    while (!Program.g_must_end)
                    {
                        try
                        {
                            int id = take_unchecked_mails_id();
                            if (id == -1)
                            {
                                MyLogger.Info("[TEST] ...parsing : no unchecked mail. will wait.");
                                for (k = 0; !Program.g_must_end && k < 600; k++)
                                    Thread.Sleep(100);
                                continue;
                            }

                            string eml_file = get_eml_file_path_from_id(id);

                            for (k = 0; k < m_mail_parsers.Count; k++)
                            {
                                if (m_mail_parsers[k].parse(id, eml_file))
                                    break;
                            }
                            if (k == m_mail_parsers.Count)
                            {
                                MyLogger.Error($"[TEST] ***Parsing Failed*** : mail.id = {id}");

                                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_UPDATE_DB))
                                    Program.g_db.set_mail_checked_flag(id, ConstEnv.MAIL_PARSING_FAILED);
                            }
                        }
                        catch (Exception exception)
                        {
                            MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                        }
                    }
                    MyLogger.Info($"[TEST] ...stop parsing thread. {i} th thread ID = {Thread.CurrentThread.ManagedThreadId}");

                }).Start();
            }
        }
        private void get_unchecked_mails_ids_from_db()
        {
            DataTable dt = Program.g_db.get_unchecked_fetched_mail_ids();

            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    int id = int.Parse(row["id"].ToString());

                    lock (m_lock_unchecked_ids)
                    {
                        m_unchecked_list.Add(id);
                    }
                }
            }
        }
        private int take_unchecked_mails_id()
        {
            int id = -1;
            lock (m_lock_unchecked_ids)
            {
                if (m_unchecked_list.Count > 0)
                {
                    id = m_unchecked_list[0];
                    m_unchecked_list.RemoveAt(0);
                }
            }
            return id;
        }
        private string get_eml_file_path_from_id(int id)
        {
            string eml_folder_path = Program.g_db.get_mail_folder_path(id);
            if (eml_folder_path == "")
                return "";

            string eml_path = Path.Combine(eml_folder_path, ConstEnv.LOCAL_MAIL_FILE_NAME);
            if (!File.Exists(eml_path))
                return "";

            return eml_path;
        }
    }
}
