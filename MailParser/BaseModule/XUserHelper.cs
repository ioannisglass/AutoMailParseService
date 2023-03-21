using MailParser;
using Logger;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserHelper
{
    public class XUserHelper
    {
        public List<UserInfo> user_info_list;

        public XUserHelper()
        {
            user_info_list = new List<UserInfo>();
        }
        public bool is_report_mode_gs(int work_mode)
        {
            if ((work_mode & ConstEnv.USER_REPORT_MODE_GS) == ConstEnv.USER_REPORT_MODE_GS)
                return true;
            return false;
        }
        public bool is_report_mode_gs(UserInfo user)
        {
            return is_report_mode_gs(user.report_mode);
        }
        public bool is_report_mode_crm(int work_mode)
        {
            if ((work_mode & ConstEnv.USER_REPORT_MODE_CRM) == ConstEnv.USER_REPORT_MODE_CRM)
                return true;
            return false;
        }
        public bool is_report_mode_crm(UserInfo user)
        {
            return is_report_mode_crm(user.report_mode);
        }
        public void save_user_info(string file_name)
        {
            try
            {
                File.WriteAllText(file_name, JsonConvert.SerializeObject(user_info_list));
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void load_account_info(string file_name)
        {
            try
            {
                user_info_list = JsonConvert.DeserializeObject<List<UserInfo>>(File.ReadAllText(file_name));

                foreach (UserInfo user in user_info_list)
                {
                    if (user.report_mode != ConstEnv.USER_REPORT_MODE_CRM && user.report_mode != ConstEnv.USER_REPORT_MODE_GS)
                        throw new Exception($"Invalid user work mode : user = {user.mail_address}");
                    Program.g_db.add_user_info(user);
                }

            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void load_user_info_from_db()
        {
            try
            {
                user_info_list.Clear();

                List<UserInfo> tmp_user_list = Program.g_db.get_user_list();

                foreach (UserInfo user in tmp_user_list)
                {
                    MyLogger.Info($"----- load user info from db");
                    MyLogger.Info($"      user mail             : {user.mail_address}");
                    MyLogger.Info($"      user password         : {user.password}");
                    MyLogger.Info($"      user mail_server      : {user.mail_server}");
                    MyLogger.Info($"      user mail_server_port : {user.mail_server_port}");
                    MyLogger.Info($"      user mail_server_type : {user.server_type}");
                    MyLogger.Info($"      user mail_server_ssl  : {user.ssl}");
                    MyLogger.Info($"      giftspread_user       : {user.giftspread_user}");
                    MyLogger.Info($"      giftspread_password   : {user.giftspread_password}");
                    MyLogger.Info($"      report mode           : {user.report_mode}");

                    user_info_list.Add(user);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void load_user_info_from_db(int work_mode, List<string> user_mail_list)
        {
            try
            {
                user_info_list.Clear();

                List<UserInfo> tmp_user_list = Program.g_db.get_user_list();

                foreach (UserInfo user in tmp_user_list)
                {
                    if (!user_mail_list.Contains(user.mail_address))
                        continue;

                    user.report_mode = work_mode;

                    MyLogger.Info($"----- load user info from db");
                    MyLogger.Info($"      user mail             : {user.mail_address}");
                    MyLogger.Info($"      user password         : {user.password}");
                    MyLogger.Info($"      user mail_server      : {user.mail_server}");
                    MyLogger.Info($"      user mail_server_port : {user.mail_server_port}");
                    MyLogger.Info($"      user mail_server_type : {user.server_type}");
                    MyLogger.Info($"      user mail_server_ssl  : {user.ssl}");
                    MyLogger.Info($"      giftspread_user       : {user.giftspread_user}");
                    MyLogger.Info($"      giftspread_password   : {user.giftspread_password}");
                    MyLogger.Info($"      work mode             : {user.report_mode}");

                    user_info_list.Add(user);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public int get_user_num()
        {
            return user_info_list.Count;
        }
        public UserInfo get_user_by_id(int id)
        {
            foreach (UserInfo user in user_info_list)
            {
                if (user.id == id)
                    return user;
            }
            return null;
        }
        public int get_account_id(string mail_address)
        {
            foreach (UserInfo user in user_info_list)
            {
                if (user.mail_address == mail_address)
                    return user.id;
            }
            return -1;
        }
        public string get_mailaddress_by_account_id(int account_id)
        {
            foreach (UserInfo user in user_info_list)
            {
                if (user.id == account_id)
                    return user.mail_address;
            }
            return "";
        }
        public void get_giftspread_account_info(int id, out string gc_user, out string gc_password)
        {
            gc_user = "";
            gc_password = "";
            foreach (UserInfo user in user_info_list)
            {
                if (user.id == id)
                {
                    gc_user = user.giftspread_user;
                    gc_password = user.giftspread_password;
                    return;
                }
            }
        }
    }
}
