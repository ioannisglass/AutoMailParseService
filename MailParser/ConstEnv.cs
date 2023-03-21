using MailHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailParser
{
    public static class ConstEnv
    {
        public static readonly int ACCOUNT_MAIL_CONNECT_FAILED = 0;

        public static readonly int PROXY_TYPE_HTTPS = 0;
        public static readonly int PROXY_TYPE_SOCKS4 = 1;
        public static readonly int PROXY_TYPE_SOCKS5 = 2;

        public static readonly string MAIL_ATTACH_TYPE_PDF = ".pdf";
        public static readonly string MAIL_ATTACH_TYPE_CSV = ".csv";

        public static readonly string LOCAL_MAIL_FILE_NAME = "message.eml";

        public static readonly int MAIL_VENDOR_TYPE_1 = 0;
        public static readonly int MAIL_VENDOR_TYPE_2 = 1;
        public static readonly int MAIL_VENDOR_TYPE_3 = 2;
        public static readonly int MAIL_VENDOR_TYPE_4 = 3;
        public static readonly int MAIL_VENDOR_TYPE_5 = 4;

        public static readonly int LOGIN_STARTED = 0;
        public static readonly int LOGIN_SUCCESS = 1;

        public static readonly int MAIL_SERVER_IMAP = 0;
        public static readonly int MAIL_SERVER_POP3 = 1;
        public static readonly int MAIL_SERVER_SMTP = 2;

        public static readonly int SCRAP_READY = 0;
        public static readonly int SCRAP_FAILED = 1;
        public static readonly int SCRAP_UNSUPPORTED = 2;
        public static readonly int SCRAP_OTHER = 3;
        public static readonly int SCRAP_CANCELED = 4;
        public static readonly int SCRAP_SUCCESS = 100;

        public static readonly string DB_MAIL_ACCOUNT_TABLE_NAME = "mail_account";
        public static readonly string DB_PROXY_TABLE_NAME = "proxy_server";
        public static readonly string DB_FETCH_MAIL_TABLE_NAME = "fetched_mail";
        public static readonly string DB_MAIL_LAST_WORK_TABLE_NAME = "mail_last_work";
        public static readonly string DB_SKIPPED_MAIL_TABLE_NAME = "skipped_mail";
        public static readonly string DB_GIFTCARDZEN_TABLE_NAME = "gift_card_zen";
        public static readonly string DB_GIFTCARDZEN_DETAILS_TABLE_NAME = "gift_card_zen_details";
        public static readonly string DB_REPORT_MAIN_TABLE_NAME = "report_main";
        public static readonly string DB_REPORT_OP_TABLE_NAME = "report_op";
        public static readonly string DB_REPORT_SC_TABLE_NAME = "report_sc";
        public static readonly string DB_REPORT_CR_TABLE_NAME = "report_cr";
        public static readonly string DB_REPORTDATA_PI_TABLE_NAME = "reportdata_cardpayment";
        public static readonly string DB_REPORTDATA_DETAILS_TABLE_NAME = "reportdata_details";
        public static readonly string DB_REPORTDATA_CASHBACK_TABLE_NAME = "reportdata_instant_cashback";
        public static readonly string DB_REPORTDATA_PRODUCT_TABLE_NAME = "reportdata_product";
        public static readonly string DB_ORDER_STATUS_TABLE_NAME = "order_status";
        public static readonly string DB_STATEMENT_FILES_TABLE_NAME = "statement_files";
        public static readonly string DB_BANK_TRANSACTIONS_TABLE_NAME = "bank_transactions";

        public static readonly int CARD_DETAILS_ALL = 0;
        public static readonly int CARD_DETAILS_V1 = 1;
        public static readonly int CARD_DETAILS_V2 = 2;

        public static readonly int MAIL_UNCHECKED = 0;
        public static readonly int MAIL_PARSING_SUCCEED = 1;
        public static readonly int MAIL_PARSING_FAILED = 2;
        public static readonly int MAIL_SCRAP_PARTIAL_FAILED = 3;

        public static readonly int OS_UNKNOWN = 1;
        public static readonly int OS_WINDOWS = 1;
        public static readonly int OS_UNIX = 2;
        public static readonly int OS_OSX = 3;

        public static int OS_TYPE = OS_WINDOWS;

        public static uint APP_WORK_MODE_FETCH_MAILS = (1 << 0);
        public static uint APP_WORK_MODE_PARSE_MAILS = (1 << 1);
        public static uint APP_WORK_MODE_CREATE_PDF_IN_PARSING = (1 << 2);
        public static uint APP_WORK_MODE_WEB_SCRAP = (1 << 3);
        public static uint APP_WORK_MODE_UPDATE_DB = (1 << 4);
        public static uint APP_WORK_MODE_REPORT = (1 << 5);
        public static uint APP_WORK_MODE_DELETE_LOCAL_MAIL = (1 << 6);

        public static uint APP_WORK_MODE_ALL = 0xFFFFFFFF;

        public static uint app_work_mode = APP_WORK_MODE_ALL;

        public static int USER_REPORT_MODE_GS = 1; // fetch just only 4 retailers (Bass, Target, Cabelas, Lowes) and update the order status to google sheet.
        public static int USER_REPORT_MODE_CRM = 2; // fetch all mails and send to crm

        public static string output_all_root_dir = "card";
        public static string output_cancel_root_dir = "cancel";

        public static string output_all_file = "cards.txt";
        public static string output_cancel_file = "canceled_cards.txt";

        public static string local_mail_all_root_dir = "Messages";
        public static string local_mail_cancel_root_dir = "Messages";


        public static string RETAILER_BASS = "Bass";
        public static string RETAILER_TARGET = "Target";
        public static string RETAILER_BESTBUY = "BestBuy";
        public static string RETAILER_AMAZON = "Amazon";
        public static string RETAILER_HOMEDEPOT = "HomeDepot";
        public static string RETAILER_SAMSCLUB = "SamsClub";
        public static string RETAILER_WALMART = "Walmart";
        public static string RETAILER_CABELAS = "Cabelas";
        public static string RETAILER_SEARS = "Sears";
        public static string RETAILER_LOWES = "Lowes";
        public static string RETAILER_EBAY = "eBay";
        public static string RETAILER_STAPLES = "Staples";
        public static string RETAILER_HP = "HP";
        public static string RETAILER_OFFICEDEPOT = "OfficeDepot";
        public static string RETAILER_KOHLS = "Kohls";
        public static string RETAILER_DELL = "DELL";
        public static string RETAILER_GOOGLEEXPRESS = "Google Express";

        public static string REPORT_ORDER_STATUS_PURCHAESD = "Order Purchased";
        public static string REPORT_ORDER_STATUS_SHIPPED = "Shipped";
        public static string REPORT_ORDER_STATUS_CANCELED = "Canceled";
        public static string REPORT_ORDER_STATUS_NOT_SHIPPED = "Not Shipped";
        public static string REPORT_ORDER_STATUS_MANUAL_CHECK = "Manual Check";
        public static string REPORT_ORDER_STATUS_PARTIAL_NOT_SHIPPED = "Partial Not Shipped";
        public static string REPORT_ORDER_STATUS_PARTIAL_CANCELED = "Partial Canceled";

        public static string REPORT_STATUS_ALL_RECEIVED = "All mails are received";
        public static string REPORT_STATUS_CR4_CASHBACK = "RetailMeNot just paid Instant Cashback";

        public static string PDF_CONVERT_FAILED = "PDF Converting Failed";
        public static string PDF_CONVERT_CANCELED = "PDF Converting Canceled";

        public static StateName[] STATE_NAMES = new StateName[]
        {
            new StateName("AL", " Alabama"),
            new StateName("AK", "Alaska"),
            new StateName("AZ", " Arizona"),
            new StateName("AR", "Arkansas"),
            new StateName("CA", "California"),
            new StateName("CO", "Colorado"),
            new StateName("CT", "Connecticut"),
            new StateName("DE", "Delaware"),
            new StateName("FL", "Florida"),
            new StateName("GA", "Georgia"),
            new StateName("HI", "Hawaii"),
            new StateName("ID", "Idaho"),
            new StateName("IL", "Illinois"),
            new StateName("IN", "Indiana"),
            new StateName("IA", "Iowa"),
            new StateName("KS", "Kansas"),
            new StateName("KY", "Kentucky"),
            new StateName("LA", "Louisiana"),
            new StateName("ME", "Maine"),
            new StateName("MD", "Maryland"),
            new StateName("MA", "Massachusetts"),
            new StateName("MI", "Michigan"),
            new StateName("MN", "Minnesota"),
            new StateName("MS", "Mississippi"),
            new StateName("MO", "Missouri"),
            new StateName("MT", "Montana"),
            new StateName("NE", "Nebraska"),
            new StateName("NV", "Nevada"),
            new StateName("NH", "New Hampshire"),
            new StateName("NJ", "New Jersey"),
            new StateName("NM", "New Mexico"),
            new StateName("NY", "New York"),
            new StateName("NC", "North Carolina"),
            new StateName("ND", "North Dakota"),
            new StateName("OH", "Ohio"),
            new StateName("OK", "Oklahoma"),
            new StateName("OR", "Oregon"),
            new StateName("PA", "Pennsylvania"),
            new StateName("RI", "Rhode Island["),
            new StateName("SC", "South Carolina"),
            new StateName("SD", "South Dakota"),
            new StateName("TN", "Tennessee"),
            new StateName("TX", "Texas"),
            new StateName("UT", "Utah"),
            new StateName("VT", "Vermont"),
            new StateName("VA", "Virginia"),
            new StateName("WA", "Washington"),
            new StateName("WV", "West Virginia"),
            new StateName("WI", "Wisconsin"),
            new StateName("WY", "Wyoming"),
            new StateName("AS", "American Samoa"),
            new StateName("GU", "Guam"),
            new StateName("MP", "Northern Mariana Islands"),
            new StateName("PR", "Puerto Rico"),
            new StateName("VI", "U.S. Virgin Islands"),
        };

        public static string[] TAX_STATES_ABBR_NAMES = new string[]
        {
            "NJ", "NY"
        };

        public static string GIFT_CARD = "Gift Card";
        public static string CREDIT_CARD = "Credit Card";

        public static string BANK_NAME_CHASE = "Chase";
        public static string BANK_NAME_WELLSFARGO = "Wells Fargo";
        public static string BANK_NAME_COB = "Capital One";
        public static string BANK_NAME_CUSTOMERS = "Customers Bank";

        public static readonly int STATEMENT_FILE_PARSE_SUCCEED = 1;
        public static readonly int STATEMENT_FILE_PARSE_FAILED = 2;
        public static readonly int STATEMENT_FILE_PARSE_IMAGE = 3;


        static public bool check_handle_flag(uint flag)
        {
            return ((app_work_mode & flag) == flag);
        }

        static public string get_output_file_path()
        {
            string out_file_path = "";
            string dir_path = "";
            string file_name = "";

            dir_path = output_all_root_dir;
            file_name = output_all_file;

            if (!Directory.Exists(dir_path))
                Directory.CreateDirectory(dir_path);
            out_file_path = Path.Combine(dir_path, file_name);

            return out_file_path;
        }
        static public string get_canceled_output_file_path()
        {
            string out_file_path = "";
            string dir_path = "";
            string file_name = "";

            dir_path = output_cancel_root_dir;
            file_name = output_cancel_file;

            if (!Directory.Exists(dir_path))
                Directory.CreateDirectory(dir_path);
            out_file_path = Path.Combine(dir_path, file_name);

            return out_file_path;
        }
        static public string get_local_mail_root_path()
        {
            string dir_path = "";

            dir_path = local_mail_all_root_dir;

            if (!Directory.Exists(dir_path))
                Directory.CreateDirectory(dir_path);

            return dir_path;
        }
        static public void get_os_type()
        {
            OperatingSystem os = Environment.OSVersion;
            PlatformID pid = os.Platform;
            switch (pid)
            {
                case PlatformID.Win32NT:
                case PlatformID.Win32S:
                case PlatformID.Win32Windows:
                case PlatformID.WinCE:
                    OS_TYPE = OS_WINDOWS;
                    break;
                case PlatformID.Unix:
                    OS_TYPE = OS_UNIX;
                    break;
                case PlatformID.MacOSX:
                    OS_TYPE = OS_OSX;
                    break;
                default:
                    OS_TYPE = OS_UNKNOWN;
                    break;
            }
        }
        static public bool is_4_retailers(string retailer)
        {
            if (retailer == ConstEnv.RETAILER_BASS)
                return true;
            if (retailer == ConstEnv.RETAILER_TARGET)
                return true;
            if (retailer == ConstEnv.RETAILER_CABELAS)
                return true;
            if (retailer == ConstEnv.RETAILER_LOWES)
                return true;
            return false;
        }
    }
    public class StateName
    {
        public readonly string full_name;
        public readonly string abbr_name;

        public StateName(string _abbr_name, string _full_name)
        {
            full_name = _full_name;
            abbr_name = _abbr_name;
        }
    }
}
