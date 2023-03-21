using MailParser;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserHelper
{
    [JsonObject(MemberSerialization.OptIn)]
    public class UserInfo
    {
        public int id;
        [JsonProperty] public string mail_address;
        [JsonProperty] public string password;

        [JsonProperty] public string mail_server;
        [JsonProperty] public int mail_server_port;
        [JsonProperty] public bool ssl;

        [JsonConverter(typeof(MailServerTypeConverter))]
        public int server_type;

        [JsonProperty] public string giftspread_user;
        [JsonProperty] public string giftspread_password;

        [JsonProperty] public int report_mode;

        public UserInfo()
        {
            id = -1;
            mail_address = "";
            password = "";
            mail_server = "";
            mail_server_port = 0;
            ssl = false;
            server_type = ConstEnv.MAIL_SERVER_IMAP;
            giftspread_user = "";
            giftspread_password = "";
            report_mode = ConstEnv.USER_REPORT_MODE_CRM;
        }
    }
    public class MailServerTypeConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            int type = (int)value;
            string conv_value = "";
            if (type == ConstEnv.MAIL_SERVER_IMAP)
                conv_value = "imap";
            else if (type == ConstEnv.MAIL_SERVER_POP3)
                conv_value = "pop3";
            else if (type == ConstEnv.MAIL_SERVER_SMTP)
                conv_value = "smtp";
            else
                throw new Exception($"Unknown mail server type : ({type})");
            writer.WriteValue(conv_value);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            int type = ConstEnv.MAIL_SERVER_IMAP;

            if (reader.TokenType == JsonToken.Null)
            {
                throw new Exception($"Unknown mail server type : ({reader.TokenType})");
            }
            else if (reader.TokenType == JsonToken.String)
            {
                string value = serializer.Deserialize(reader, Type.GetType("string")) as string;
                value = value.ToLower();
                if (value == "imap")
                    type = ConstEnv.MAIL_SERVER_IMAP;
                else if (value == "pop3")
                    type = ConstEnv.MAIL_SERVER_POP3;
                else if (value == "smtp")
                    type = ConstEnv.MAIL_SERVER_SMTP;
                else
                    throw new Exception($"Unknown mail server type : ({value})");
                return type;
            }
            else
            {
                throw new Exception($"Unknown mail server type : ({reader.TokenType})");
            }
        }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(int);
        }
    }
}
