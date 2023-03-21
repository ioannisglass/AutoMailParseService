using Logger;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace MailParser
{
    public class AppSettings<T> where T : new()
    {
        private const string DEFAULT_FILENAME = "settings.ini";

        public void Save(string fileName = DEFAULT_FILENAME)
        {
            try
            {
                File.WriteAllText(fileName, JsonConvert.SerializeObject(this, Newtonsoft.Json.Formatting.Indented));
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }

        public static void Save(T pSettings, string fileName = DEFAULT_FILENAME)
        {
            File.WriteAllText(fileName, JsonConvert.SerializeObject(pSettings, Newtonsoft.Json.Formatting.Indented));
        }

        public static T Load(string fileName = DEFAULT_FILENAME)
        {
            try
            {
                T t = new T();
                if (File.Exists(fileName))
                    t = JsonConvert.DeserializeObject<T>(File.ReadAllText(fileName));
                else
                    return default(T);
                return t;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return default(T);
            }
        }
    }

    public class UserSetting : AppSettings<UserSetting>
    {
        public int delay_time = 60;
        public int mail_download_retry_second = 900;
        public int pdf_converter_retry_msec = 1000;
        public int web_scrap_retry_msec = 1000;
        public int web_scrap_retry_max_num = 5;
        
        public string first_url = "";

        public string gsheets_credential_json_file = "Google Spreadsheet 13140-b5d60babc482.json";

        public string gsheets_id_4_retailers = "1CUOxqaN_Gei3W7H6VWHxtnE1fYaTyrT8WfVKA-TMaAM";
        public string gsheets_sheet_name_4_retailers = "Sheet1";

        public string gsheets_id_tax_step2 = "18XXLJECGr9Pi2OBxD5koes4maHpFLilvIxmPxeap9yk";
        public string gsheets_sheet_name_tax_step2 = "Sheet1";

        public int status_manual_check_hours = 36;

        public string result_file_name = "result.txt";
        public string[] error_res_list = new string[]
        { "ERROR_ZERO_CAPTCHA_FILESIZE", "ERROR_WRONG_USER_KEY", "ERROR_KEY_DOES_NOT_EXIST",
            "ERROR_ZERO_BALANCE", "ERROR_PAGEURL", "ERROR_NO_SLOT_AVAILABLE", "ERROR_TOO_BIG_CAPTCHA_FILESIZE",
            "ERROR_WRONG_FILE_EXTENSION", "ERROR_IMAGE_TYPE_NOT_SUPPORTED", "CAPCHA_NOT_READY" };
        public Dictionary<string, string> MailServer_Info_IMAP = new Dictionary<string, string>
        {
            {"@gmail.com", "imap.gmail.com|993"},
            {"@outlook.com", "imap-mail.outlook.com|993"},
        };
        public Dictionary<string, string> MailServer_Info_POP = new Dictionary<string, string>
        {
            {"@outlook.com", "pop-mail.outlook.com|995"},
            {"@gmail.com", "pop.gmail.com|995"}
        };
        public Dictionary<string, string> MailServer_Info_SMTP = new Dictionary<string, string>
        {
            {"@outlook.com", "smtp-mail.outlook.com"},
            {"@gmail.com", "smtp.gmail.com|TLS|587|SSL|465"}
        };

        public string database_name = "mailparser_db";

        public string db_hostname = "localhost";
        public int db_port = 3306;
        public string db_username = "root";
        public string db_password = "";
        public bool db_use_ssh = false;
        public string db_ssh_hostname = "";
        public int db_ssh_port = 22;
        public string db_ssh_username = "";
        public string db_ssh_password = "";
        public string db_ssh_keyfile = "";

        public string captcha_api = "39a5bfc01408ffcc26290069dfc65548";

    }
}
