using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using BaseModule;
using DbHelper;
using Logger;
using MailHelper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ReportStat;
using StatementHelper;
using UserHelper;
using Utils;
using WebAuto;

namespace MailParser
{
    static class Program
    {
        public static UserSetting g_setting = new UserSetting();
        public static int m_thread_count = 0;
        public static DbMgr g_db = null;
        public static XUserHelper g_user = null;
        public static bool g_must_end = false;
        public static string g_working_directory = "";

        static void Main(string[] args)
        {
            try
            {
                ConstEnv.get_os_type();
                string cur_process_name = ProcessInfo.get_current_process_name();
                MyLogger.Info($"current process name = {cur_process_name}, PID = {Process.GetCurrentProcess().Id}");
                MyLogger.Info($"{Process.GetCurrentProcess().ProcessName} started.... argument number = {args.Length}");
                for (int i = 0; i < args.Length; i++)
                    MyLogger.Info($"{i + 1}th argument : {args[i]}");

                g_working_directory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

                if (args.Length < 1)
                {
                    MyLogger.Error("Invalid command line arguments.");
                    return;
                }

                g_setting = UserSetting.Load();
                if (g_setting == null)
                {
                    g_setting = new UserSetting();
                    g_setting.Save();
                }
                MyLogger.Info("load program settings.");

                g_db = new DbMgr(g_setting.database_name, g_setting.db_hostname, g_setting.db_port, g_setting.db_username, g_setting.db_password,
                        g_setting.db_use_ssh, g_setting.db_ssh_hostname, g_setting.db_ssh_port, g_setting.db_ssh_username, g_setting.db_ssh_password, g_setting.db_ssh_keyfile);
                if (!g_db.connectable())
                {
                    MyLogger.Error("Can not connect to the local db.");
                    return;
                }
                MyLogger.Info("db connected.");

                g_user = new XUserHelper();

                InterComm.clear_intercomm_file();

                if (args[0] == "-set")
                {
                    if (args.Length != 3)
                    {
                        MyLogger.Error("Invalid command line arguments.");
                        return;
                    }
                    set_configuration(args[1], args[2]);
                }
                else if (args[0].ToUpper() == "-START")
                {
                    start_work(args);
                }
                else if (args[0] == "-stop")
                {
                    stop_work();
                }
                else if (args[0] == "-parse")
                {
                    if (args.Length < 2)
                    {
                        MyLogger.Error("Invalid command line arguments.");
                        return;
                    }

                    ConstEnv.app_work_mode = ConstEnv.APP_WORK_MODE_PARSE_MAILS;

                    int i = 1;
                    for (; i < args.Length; i++)
                    {
                        if (args[i][0] != '-')
                            break;

                        string temp = args[i].ToLower().Trim();

                        if (temp == "-pdf")
                            ConstEnv.app_work_mode |= ConstEnv.APP_WORK_MODE_CREATE_PDF_IN_PARSING;
                        if (temp == "-scrap")
                            ConstEnv.app_work_mode |= ConstEnv.APP_WORK_MODE_WEB_SCRAP;
                        if (temp == "-db")
                            ConstEnv.app_work_mode |= ConstEnv.APP_WORK_MODE_UPDATE_DB;
                        if (temp == "-report")
                            ConstEnv.app_work_mode |= ConstEnv.APP_WORK_MODE_REPORT;
                        if (temp == "-del_mail")
                            ConstEnv.app_work_mode |= ConstEnv.APP_WORK_MODE_DELETE_LOCAL_MAIL;
                    }

                    if (i < args.Length - 2)
                        test_parse_mail(args[i], "");
                    else
                        test_parse_mail(args[i], args[i + 1]);
                }
                else if (args[0] == "-fetch")
                {
                    test_fetch_mail();
                }
                else if (args[0] == "-get")
                {
                    if (args.Length == 2 && args[1] == "log_summary")
                        get_log_summary();
                    else if (args.Length == 2 && args[1] == "order_cancelled_mails")
                        download_order_cancelled_mails();
                    else if (args.Length == 3 && args[1] == "mails_by_type")
                    {
                        get_mails_by_type(args[2]);
                    }
                    else if (args.Length == 4 && args[1] == "mails_by_id")
                        get_mails_by_id(args[2], args[3]);
                }
                else if (args[0] == "-scrap")
                {
                    if (args.Length == 4 && args[1] == "gift_card_zen" && args[2] != "" && args[3] != "")
                        ScrapGiftcardZen(args[2], args[3]);
                }
                else if (args[0] == "-statement")
                {
                    if (args.Length < 4)
                    {
                        MyLogger.Error("Invalid command line arguments.");
                        return;
                    }
                    if (args[1] == "-bank" && args[2] != "" && args[3] != "")
                    {
                        if (args[2] == "-dir")
                            parse_statement_directory(args[3]);
                        else if (args[2] == "-file")
                            parse_statement_file(args[3]);
                    }
                }
                else if (args[0] == "-test")
                {
                    //collect_tax_data();
                    //recheck_already_fetched_mails();
                    //update_file_hash();
                    //load_result_card_file();
                    //delete_all_order_reports();
                    //reparse_cr2_mails();
                    //delete_duplicated_mails();
                    //create_old_crm_table();
                    create_old_crm_diff_csv();
                    //make_order_status_table();
                }

                g_setting.Save();

                MyLogger.Info($"Terminating process... {Process.GetCurrentProcess().ProcessName}, PID = {Process.GetCurrentProcess().Id}");
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        
        static void set_configuration(string type, string cfg_fpath)
        {
            bool already_running = (InterComm.find_working_process().Length > 0);
            MyLogger.Info("running worker process ? - " + ((already_running) ? "Yes" : "No"));

            if (already_running)
                InterComm.kill_working_Process();

            if (type == "user")
                g_user.load_account_info(cfg_fpath);
            else if (type == "proxy")
                BaseData.load_proxy_info(cfg_fpath);

            if (already_running)
                InterComm.start_working_Process();
        }
        static void start_work(string[] args)
        {
            try
            {
                CancellationTokenSource cts = new CancellationTokenSource();

                g_must_end = false;

                InterComm.monitor_interprocess_command(cts);

                if (args.Length < 3)
                {
                    g_user.load_user_info_from_db();
                }
                else
                {
                    int report_mode = int.Parse(args[1]);
                    if (report_mode != ConstEnv.USER_REPORT_MODE_CRM && report_mode != ConstEnv.USER_REPORT_MODE_GS)
                        throw new Exception($"Invalid work mode.");

                    List<string> user_mail_list = new List<string>();
                    for (int i = 2; i < args.Length; i++)
                        user_mail_list.Add(args[i]);

                    g_user.load_user_info_from_db(report_mode, user_mail_list);
                }

                if (g_user.get_user_num() == 0)
                    throw new Exception("Not found user");

                BaseData.load_proxy_info_from_db();

                XStatHelper reporter = new XStatHelper();
                reporter.start_report();

                XMailParser mail_parser = new XMailParser();
                mail_parser.start_parse_mails();

                XMailHelper mail_fetcher = new XMailHelper();
                mail_fetcher.start_fetch_mails(cts);

                while (!g_must_end)
                    Thread.Sleep(500);
            }
            catch(Exception ex)
            {
                g_must_end = true;
                MyLogger.Error($"Error catched in start_work - {ex.Message}");
            }            
        }
        static void stop_work()
        {
            bool already_running = (InterComm.find_working_process().Length > 0);
            MyLogger.Info("running worker process ? - " + ((already_running) ? "Yes" : "No"));

            if (already_running)
                InterComm.kill_working_Process();
        }

        static void test_fetch_mail()
        {
            try
            {
                CancellationTokenSource cts = new CancellationTokenSource();

                g_must_end = false;

                InterComm.monitor_interprocess_command(cts);

                g_user.load_user_info_from_db();
                BaseData.load_proxy_info_from_db();

                XMailHelper mail_fetcher = new XMailHelper();
                mail_fetcher.start_fetch_mails(cts);

                while (!g_must_end)
                    Thread.Sleep(500);
            }
            catch (Exception exception)
            {
                g_must_end = true;
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        static void test_parse_mail(string parse_type, string parse_dest)
        {
            try
            {
                CancellationTokenSource cts = new CancellationTokenSource();

                g_must_end = false;

                InterComm.monitor_interprocess_command(cts);

                //g_user.load_user_info_from_db();
                g_user.load_user_info_from_db(ConstEnv.USER_REPORT_MODE_CRM, new List<string>() { "info@doublegame.net" });
                BaseData.load_proxy_info_from_db();

                XStatHelper reporter = new XStatHelper();
                reporter.start_report();

                XMailParser mail_parser = new XMailParser();

                if (parse_type.ToUpper() == "ALL")
                {
                    mail_parser.start_parse_mails();
                }
                else if (parse_type.ToUpper() == "BY_ID")
                {
                    string[] id_str_list = parse_dest.Split(',');
                    int[] mail_ids = new int[id_str_list.Length];
                    for (int i = 0; i < mail_ids.Length; i++)
                    {
                        mail_ids[i] = int.Parse(id_str_list[i].Trim());
                        MyLogger.Info($"Add mail id to parse : {mail_ids[i]}");
                    }
                    mail_parser.parse_specific_mails(mail_ids);
                }
                else if (parse_type.ToUpper() == "BY_FILE")
                {
                    string failed_ids = File.ReadAllText(parse_dest);
                    failed_ids.Replace("\r\n", "\n");
                    string[] lines = failed_ids.Split('\n');
                    List<int> mail_ids = new List<int>();
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string line = lines[i].Trim();
                        if (line == "")
                            continue;
                        mail_ids.Add(int.Parse(line));
                    }
                    mail_parser.parse_specific_mails(mail_ids.ToArray());
                }
                else
                {
                    MyLogger.Info("Invalid parsing type. It must \"all\" or \"by_id\" or \"by_file\".");
                    return;
                }

                while (!g_must_end)
                    Thread.Sleep(500);
            }
            catch (Exception exception)
            {
                g_must_end = true;
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        static void get_log_summary()
        {
            try
            {
                List<string> log_files = new List<string>();

                if (File.Exists("Log.txt"))
                {
                    File.Copy("Log.txt", "Log.txt.0");
                    log_files.Add("Log.txt.0");
                }

                int i = 1;
                while (true)
                {
                    string fpath = $"Log.txt.{i}";
                    if (!File.Exists(fpath))
                        break;
                    log_files.Insert(0, fpath);
                    i++;
                }

                int card_mail_num = 0;
                int parsing_failed_mail_num = 0;
                int pdf_num = 0;
                int pdf_failed_num = 0;
                List<string> pdf_failed_paths = new List<string>();
                int unchecked_mail_num = 0;
                int skipped_mail_num = 0;
                int add_web_link_num = 0;
                List<string> add_web_links = new List<string>();
                int web_scrap_success_num = 0;
                int web_scrap_unsupported_num = 0;
                int web_scrap_todo_num = 0;
                SortedDictionary<string, int> card_mail_nums = new SortedDictionary<string, int>();
                SortedDictionary<string, List<string>> parse_failed_mails = new SortedDictionary<string, List<string>>();
                List<string> unsupported_web_links = new List<string>();
                List<string> todo_web_links = new List<string>();

                List<UserInfo> user_list = Program.g_db.get_user_list();
                foreach (UserInfo user in user_list)
                {
                    int account_id = Program.g_user.get_account_id(user.mail_address);
                    DataTable dt = Program.g_db.get_mail_num_by_mail_type(account_id);
                    if (dt != null && dt.Rows != null)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string type = row["mail_type"].ToString();
                            int num = int.Parse(row["num"].ToString());

                            if (card_mail_nums.ContainsKey(type))
                                card_mail_nums[type] += num;
                            else
                                card_mail_nums.Add(type, num);

                            card_mail_num += num;
                        }
                    }

                    dt = Program.g_db.get_mails_by_check_state(account_id, ConstEnv.MAIL_PARSING_FAILED);
                    if (dt != null && dt.Rows != null)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string type = row["mail_type"].ToString();
                            int mail_id = int.Parse(row["id"].ToString());

                            if (parse_failed_mails.ContainsKey(type))
                                parse_failed_mails[type].Add(mail_id.ToString());
                            else
                                parse_failed_mails.Add(type, new List<string>() { mail_id.ToString() });

                            parsing_failed_mail_num++;
                        }
                    }

                    dt = Program.g_db.get_mails_by_check_state(account_id, ConstEnv.MAIL_UNCHECKED);
                    if (dt != null && dt.Rows != null)
                        unchecked_mail_num = dt.Rows.Count;

                    skipped_mail_num += Program.g_db.get_skipped_mail_num(account_id);
                }

                foreach (string log_file in log_files)
                {
                    var lines = File.ReadLines(log_file);
                    foreach (string line in lines)
                    {
                        if (line.IndexOf("end eml->pdf exit code = ") != -1)
                        {
                            pdf_num++;
                        }
                        else if (line.IndexOf("Failed eml->pdf : ") != -1)
                        {
                            string temp = line.Substring(line.IndexOf("Failed eml->pdf : ") + "Failed eml->pdf : ".Length).Trim();
                            if (!pdf_failed_paths.Contains(temp))
                                pdf_failed_paths.Add(temp);

                            pdf_failed_num++;
                        }
                        else if (line.IndexOf("*** Add web link ***") != -1)
                        {
                            string temp = line.Substring(line.IndexOf("*** Add web link ***") + "*** Add web link ***".Length).Trim();
                            if (!add_web_links.Contains(temp))
                                add_web_links.Add(temp);

                            add_web_link_num++;
                        }
                        else if (line.IndexOf("Scrap success -") != -1)
                        {
                            web_scrap_success_num++;
                        }
                        else if (line.IndexOf("Scrap url other reason -") != -1)
                        {
                            string temp = line.Substring(line.IndexOf("Scrap url other reason -") + "Scrap url other reason -".Length).Trim();
                            if (!todo_web_links.Contains(temp))
                                todo_web_links.Add(temp);

                           web_scrap_todo_num++;
                        }
                        else if (line.IndexOf("It's not unsupported web link.") != -1)
                        {
                            string temp = line.Substring(line.IndexOf("It's not unsupported web link.") + "It's not unsupported web link.".Length).Trim();
                            if (!unsupported_web_links.Contains(temp))
                                unsupported_web_links.Add(temp);

                            web_scrap_unsupported_num++;
                        }
                    }
                }

                using (var sw = new StreamWriter("Log_Summary.txt"))
                {
                    sw.WriteLine("=================================================================");
                    sw.WriteLine($"Written at {DateTime.Now.ToString()}");
                    sw.WriteLine("=================================================================");
                    sw.WriteLine($"Card mail num                  = {card_mail_num}");
                    sw.WriteLine($"Parsing failed mail num        = {parsing_failed_mail_num}");
                    sw.WriteLine($"Parsing not completed mail num = {unchecked_mail_num}");
                    sw.WriteLine($"converted PDF num              = {pdf_num}");
                    sw.WriteLine($"converting failed PDF num      = {pdf_failed_num}");
                    sw.WriteLine($"Skipped mail num               = {skipped_mail_num}");
                    sw.WriteLine($"Web Link num                   = {add_web_link_num}");
                    sw.WriteLine($"Web scrapped num               = {web_scrap_success_num}");
                    sw.WriteLine($"Web scrapped unsupported num   = {web_scrap_unsupported_num}");
                    sw.WriteLine($"Web scrapped todo num          = {web_scrap_todo_num}");
                    sw.WriteLine("-----------------------------------------------------------------");
                    sw.WriteLine("");

                    sw.WriteLine("-----------------------------------------------------------------");
                    sw.WriteLine("Card mails.");
                    foreach (string key in card_mail_nums.Keys)
                    {
                        sw.WriteLine($"{key}\t\t{card_mail_nums[key]}");
                    }
                    sw.WriteLine("");

                    sw.WriteLine("-----------------------------------------------------------------");
                    sw.WriteLine("Parsing Failed Card mails. (Summary)");
                    foreach (string key in parse_failed_mails.Keys)
                    {
                        sw.WriteLine($"{key}\t\t{parse_failed_mails[key].Count}");
                    }
                    sw.WriteLine("");

                    foreach (string key in parse_failed_mails.Keys)
                    {
                        sw.WriteLine("-----------------------------------------------------------------");
                        sw.WriteLine($"Parsing Failed Card mails : {key}");
                        foreach (string id in parse_failed_mails[key])
                            sw.WriteLine(id);
                        sw.WriteLine("");
                    }
                    sw.WriteLine("");

                    sw.WriteLine("-----------------------------------------------------------------");
                    sw.WriteLine($"Converting Failed PDF Files");
                    foreach (string link in pdf_failed_paths)
                        sw.WriteLine(link);
                    sw.WriteLine("");

                    sw.WriteLine("-----------------------------------------------------------------");
                    sw.WriteLine($"Unsupported Web Links");
                    foreach (string link in unsupported_web_links)
                        sw.WriteLine(link);
                    sw.WriteLine("");

                    sw.WriteLine("-----------------------------------------------------------------");
                    sw.WriteLine($"ToDo (not scrapped by the other reason) Web Links");
                    foreach (string link in todo_web_links)
                        sw.WriteLine(link);
                    sw.WriteLine("");
                    sw.WriteLine("");
                    sw.WriteLine("=================================================================");
                }
            }
            catch (Exception exception)
            {
                g_must_end = true;
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }

            if (File.Exists("Log.txt.0"))
                File.Delete("Log.txt.0");
        }
        static void get_mails_by_type(string mail_type)
        {
            List<string> id_list = new List<string>();

            if (!Directory.Exists("ExtractMails"))
                Directory.CreateDirectory("ExtractMails");

            string mail_dir = Path.Combine("ExtractMails", mail_type);
            if (Directory.Exists(mail_dir))
                Directory.Delete(mail_dir, true);
            Directory.CreateDirectory(mail_dir);

            try
            {
                MyLogger.Info($"Start copy mails by mail type {mail_type}");

                DataTable dt = Program.g_db.get_mail_localPaths_by_mail_type(mail_type);
                if (dt != null && dt.Rows != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string local_eml_folder = row["local_download_path"].ToString();
                        string id = row["id"].ToString();

                        string src = Path.Combine(local_eml_folder, ConstEnv.LOCAL_MAIL_FILE_NAME);
                        if (!File.Exists(src))
                            continue;
                        string dst = Path.Combine(mail_dir, $"{id}.eml");
                        if (File.Exists(dst))
                            File.Delete(dst);
                        File.Copy(src, dst, true);

                        id_list.Add(id);

                        MyLogger.Info($"File Copy {src} -> {dst}");
                    }
                }

                File.WriteAllLines(Path.Combine(mail_dir, "mail_id.txt"), id_list.ToArray());
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }

            MyLogger.Info($"End copy mails by mail type {mail_type} : {id_list.Count} copied.");
        }
        static void get_mails_by_id(string mail_type, string mail_id_file)
        {
            int n = 0;
            List<int> id_list = new List<int>();

            if (!Directory.Exists("ExtractMails"))
                Directory.CreateDirectory("ExtractMails");

            string mail_dir = Path.Combine("ExtractMails", mail_type);
            if (!Directory.Exists(mail_dir))
                Directory.CreateDirectory(mail_dir);

            try
            {
                MyLogger.Info($"Start copy mails by mail type {mail_type}");

                string mail_ids = File.ReadAllText(mail_id_file);
                mail_ids.Replace("\r\n", "\n");
                string[] lines = mail_ids.Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (line == "")
                        continue;
                    id_list.Add(int.Parse(line));
                }

                foreach (int id in id_list)
                {
                    string local_eml_folder = Program.g_db.get_mail_folder_path(id);
                    string src = Path.Combine(local_eml_folder, ConstEnv.LOCAL_MAIL_FILE_NAME);
                    if (!File.Exists(src))
                        continue;
                    string dst = Path.Combine(mail_dir, $"{id}.eml");
                    if (File.Exists(dst))
                        File.Delete(dst);
                    File.Copy(src, dst, true);
                    n++;
                    MyLogger.Info($"File Copy {src} -> {dst}");
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }

            MyLogger.Info($"End copy mails by mail type {mail_type} : {n} of {id_list.Count} copied.");
        }
        static void download_order_cancelled_mails()
        {
            try
            {
                MyLogger.Info("Start downloading order cancellation mails from skipped_mail DB");

                CancellationTokenSource cts = new CancellationTokenSource();
                List<UserInfo> user_list = Program.g_db.get_user_list();

                foreach (UserInfo user in user_list)
                {
                    int account_id = Program.g_user.get_account_id(user.mail_address);

                    DataTable dt = Program.g_db.get_order_cancelled_mails_from_skipped_table_by_account(account_id);
                    if (dt == null)
                        continue;

                    XMailChecker mailChecker = new XMailChecker(user);

                    mailChecker.download_mails_from_skipped_dt(dt, cts);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }

            MyLogger.Info("End downloading order cancellation mails from skipped_mail DB");
        }
        static async void ScrapGiftcardZen(string user, string password)
        {
            try
            {
                MyLogger.Info($"Start scrapping gift card zen. user = {user}, password = {password}");

                Program.g_db.clear_gift_card_zen_tables();

                KWebGCZen scrapper = new KWebGCZen();

                int ret = await scrapper.scrap_link(user, password);
                if (ret == ConstEnv.SCRAP_FAILED)
                {
                    MyLogger.Error("Failed scrapping gift card zen");
                    return;
                }

                List<GiftCardZenSales> card_list = scrapper.m_lst_GCZ_Sales;

                MyLogger.Info($"Scrapped count : {card_list.Count}");

                Program.g_db.fill_gift_card_zen_tables(card_list);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            MyLogger.Info("End scrapping gift card zen");
        }
        static void revise_card_txt()
        {
            List<string> multi_order_vendors = new List<string>();

            string card_txt = ConstEnv.get_output_file_path();
            string card_txt_1 = card_txt + ".1";

            string[] lines = File.ReadAllLines(card_txt);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                string new_line = line;

                if (line.IndexOf("_order_id\":") != -1 || line.IndexOf("_order\":") != -1)
                {
                    string order = line.Substring(line.IndexOf(":") + 1).Trim();
                    order = order.Substring(1, order.Length - 3); // <"2023336",> -> <2023336>
                    string[] order_list = order.Split(',');
                    string new_order = "";
                    string vendor_type = "";
                    int j;
                    if (order_list.Length > 1)
                    {
                        vendor_type = lines[i - 1].Trim();
                        vendor_type = vendor_type.Substring(vendor_type.IndexOf(":") + 1).Trim();
                        vendor_type = vendor_type.Substring(1, vendor_type.Length - 3);
                        if (!multi_order_vendors.Contains(vendor_type))
                            multi_order_vendors.Add(vendor_type);

                        for (j = 0; j < order_list.Length; j++)
                        {
                            string[] new_orders = new_order.Split(',');
                            if (!new_orders.Contains(order_list[j]))
                                new_order += (new_order == "") ? order_list[j] : "," + order_list[j];
                        }
                    }
                    else
                    {
                        new_order = order;
                    }
                    if (order != new_order)
                    {
                        MyLogger.Info($"Order Changed ({vendor_type}) : {order} -> {new_order}");

                        new_line = line.Substring(0, line.IndexOf(":"));
                        new_line += $": \"{new_order}\",";
                    }
                }

                File.AppendAllText(card_txt_1, new_line + "\n"/*Environment.NewLine*/);
            }

            File.Delete(card_txt);
            System.IO.File.Move(card_txt_1, card_txt);

            foreach (string v in multi_order_vendors)
            {
                MyLogger.Info($"Multi order vendors : {v}");
            }
        }
        static void load_result_card_file()
        {
            try
            {
                string card_file = ConstEnv.get_output_file_path();
                string json_text = File.ReadAllText(card_file);
                KReportJsonHelper stat = new KReportJsonHelper();
                stat.load_from_json_text(json_text);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        static void collect_tax_data()
        {
            g_user.load_user_info_from_db(ConstEnv.USER_REPORT_MODE_CRM, new List<string>() { "info@doublegame.net" });
            List<UserInfo> user_list = Program.g_user.user_info_list;
            if (user_list.Count != 1)
            {
                MyLogger.Info($"Invalid user num : {user_list.Count}");
                return;
            }

            UserInfo user = user_list[0];

            ReportSalesTax report_tax = new ReportSalesTax();
            report_tax.collect_tax_data(user.id);
        }
        static void recheck_already_fetched_mails()
        {
            g_user.load_user_info_from_db(ConstEnv.USER_REPORT_MODE_CRM, new List<string>() { "info@doublegame.net" });

            List<string> mail_folers = new List<string>();
            List<uint> uids = new List<uint>();

            string[] lines = File.ReadAllLines("already_fetched.txt");
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line == "")
                    continue;
                line = line.Substring(0, line.Length - 1);
                string mail_folder = "";
                UInt32 validity = 0;
                uint uid = 0;

                string temp = line.Substring(line.LastIndexOf(",") + 1);
                uid = uint.Parse(temp);

                line = line.Substring(0, line.LastIndexOf(","));
                validity = UInt32.Parse(temp);

                line = line.Substring(0, line.LastIndexOf(","));
                mail_folder = line;

                mail_folers.Add(mail_folder);
                uids.Add(uid);
            }

            MyLogger.Info($"START redownload already-fecthed mails : Num = {mail_folers.Count}");

            CancellationTokenSource cts = new CancellationTokenSource();
            List<UserInfo> user_list = Program.g_user.user_info_list;
            if (user_list.Count != 1)
            {
                MyLogger.Info($"Invalid user num : {user_list.Count}");
                return;
            }

            UserInfo user = user_list[0];

            int account_id = Program.g_user.get_account_id(user.mail_address);

            XMailChecker mailChecker = new XMailChecker(user);

            mailChecker.redownload_fecthed_mails(mail_folers, uids, cts);

            MyLogger.Info($"END redownload already-fecthed mails : Remaining Num = {mail_folers.Count}");
        }
        static void update_file_hash()
        {
            g_db.update_file_hash();
        }
        static void delete_all_order_reports()
        {
            g_user.load_user_info_from_db(ConstEnv.USER_REPORT_MODE_CRM, new List<string>() { "info@doublegame.net" });
            List<UserInfo> user_list = Program.g_user.user_info_list;
            if (user_list.Count != 1)
            {
                MyLogger.Info($"Invalid user num : {user_list.Count}");
                return;
            }
            UserInfo user = user_list[0];
            int account_id = Program.g_user.get_account_id(user.mail_address);

            g_db.delete_reports_by_mail_type("OP_%", account_id);
            g_db.delete_reports_by_mail_type("SC_%", account_id);
            g_db.delete_reports_by_mail_type("CC_%", account_id);
        }
        static void reparse_cr2_mails()
        {
            try
            {
                CancellationTokenSource cts = new CancellationTokenSource();

                g_must_end = false;

                InterComm.monitor_interprocess_command(cts);

                //g_user.load_user_info_from_db();
                g_user.load_user_info_from_db(ConstEnv.USER_REPORT_MODE_CRM, new List<string>() { "info@doublegame.net" });
                BaseData.load_proxy_info_from_db();

                XMailParser mail_parser = new XMailParser();

                int[] mail_ids = g_db.get_empty_pin_cr2_mail_ids();
                //int[] mail_ids = g_db.get_parsing_failed_cr2_mail_ids();
                foreach (int id in mail_ids)
                {
                    g_db.delete_report_data_by_mail_id(id);

                    MyLogger.Info($"Delete report by mail id : {id}");
                }

                MyLogger.Info($"mail number to parse : {mail_ids.Length}");

                mail_parser.parse_specific_mails(mail_ids);

                while (!g_must_end)
                    Thread.Sleep(500);
            }
            catch (Exception exception)
            {
                g_must_end = true;
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        static void delete_duplicated_mails()
        {
            string[] hash_list = g_db.get_duplicated_mail_hash_code();
            foreach (string hash_code in hash_list)
            {
                g_db.delete_duplicated_feteched_mail(hash_code);

                MyLogger.Info($"Deleted duplicated mail. hash : {hash_code}");
            }
        }
        static void change_giftcardzen_csv_time_format()
        {
            string csv_file = "E:\\12_josh_work\\Document\\Send\\HomeDepotGC.csv";
            string out_csv_file = "E:\\12_josh_work\\Document\\Send\\HomeDepotGC_out.csv";

            string[] lines = File.ReadAllLines(csv_file);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line == "")
                    continue;

                string temp = line.Substring(0, line.IndexOf(",")).Trim();
                if (temp.IndexOf(" at ") != -1)
                {
                    temp = temp.Replace(" at ", " ");
                    temp = temp.Replace("AM", " AM");
                    temp = temp.Replace("PM", " PM");

                    DateTime time = DateTime.Parse(temp);
                    temp = time.ToString("yyyy-MM-dd hh:mm:ss");
                    temp = temp + line.Substring(line.IndexOf(","));
                }
                else
                {
                    temp = line;
                }
                File.AppendAllText(out_csv_file, temp + "\n");
            }
        }
        static void create_old_crm_table()
        {
            string csv_file = "E:\\12_josh_work\\Document\\old CRM\\giftcards.csv";

            //DataTable dt = CSVUtil.csv2dt(csv_file);

            //g_db.create_old_crm_table(dt);
            g_db.BulkInsertCSV("old_crm", csv_file, "\r\n");
        }
        static void create_old_crm_diff_csv()
        {
            string csv_file = "E:\\12_josh_work\\Document\\old CRM\\giftcards.csv";
            DataTable dt = CSVUtil.csv2dt(csv_file);

            List<string> id_list = new List<string>();
            string diff_csv_file_1 = "E:\\12_josh_work\\Document\\old CRM\\File2.csv";
            string[] lines = File.ReadAllLines(diff_csv_file_1);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (i == 0 || line == "")
                    continue;

                string id = line.Substring(0, line.IndexOf(",")).Trim();
                id_list.Add(id);
            }

            MyLogger.Info($"File 2 data num : {id_list.Count}");

            int num = 0;
            int total = dt.Rows.Count;
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = dt.Rows[i];

                string id = row["id"].ToString();

                if (!id_list.Contains(id))
                    dt.Rows.RemoveAt(i);

                MyLogger.Info($"File 2 data Process : {++num} / {total}");
            }

            string diff_csv_file_2 = "E:\\12_josh_work\\Document\\old CRM\\File2_out.csv";
            CSVUtil.dt2csv(dt, diff_csv_file_2);
        }
        static void make_order_status_table()
        {
            DataTable dt = g_db.get_order_reports_from_db();
            if (dt == null || dt.Rows == null)
                return;

            int num = 0;
            foreach (DataRow row in dt.Rows)
            {
                int account_id = int.Parse(row["account_id"].ToString());
                int report_id = int.Parse(row["report_id"].ToString());
                string order_id = row["order_id"].ToString();
                string mail_type = row["mail_type"].ToString();
                string order_status = row["order_status"].ToString();
                DateTime senttime = DateTime.Parse(row["mail_senttime"].ToString());

                g_db.insert_order_status_to_db(account_id, report_id, order_id, mail_type, order_status, senttime);

                MyLogger.Info($"Insert order status : {++num} / {dt.Rows.Count}");
            }
        }
        static void parse_statement_directory(string direcotry)
        {
            //try
            {
                if (!Directory.Exists(direcotry))
                    throw new Exception($"Directory Not Found : {direcotry}");

                string[] subdirectoryEntries = Directory.GetDirectories(direcotry);
                foreach (string subdirectory in subdirectoryEntries)
                {
                    string sub_dir = System.IO.Path.GetFileName(subdirectory);
                    if (sub_dir.ToLower() == "$recycle.bin" || sub_dir.ToLower() == "system volume information")
                        continue;
                    parse_statement_directory(subdirectory);
                }

                string[] fileEntries = Directory.GetFiles(direcotry, "*.pdf");
                foreach (string fileName in fileEntries)
                {
                    parse_statement_file(fileName);
                }
            }
            //catch (Exception exception)
            //{
            //    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            //}
        }
        static void parse_statement_file(string pdf_file)
        {
            //try
            {
                if (!File.Exists(pdf_file))
                    throw new Exception($"File Not Found : {pdf_file}");

                BSHelper bs_helper = new BSHelper();

                bs_helper.parse_pdf(pdf_file);
            }
            //catch (Exception exception)
            //{
            //    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            //}
        }
    }
}
