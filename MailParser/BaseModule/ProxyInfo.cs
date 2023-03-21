using MailParser;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseModule
{
    public class ProxyInfo
    {
        [JsonConverter(typeof(ProxyServerTypeConverter))]
        public int type;
        public string host;
        public int port;
        public string username;
        public string password;        
    }
    public class ProxyServerTypeConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            int type = (int)value;
            string conv_value = "";
            if (type == ConstEnv.PROXY_TYPE_HTTPS)
                conv_value = "https";
            else if (type == ConstEnv.PROXY_TYPE_SOCKS4)
                conv_value = "socks4";
            else if (type == ConstEnv.PROXY_TYPE_SOCKS5)
                conv_value = "socsk5";
            else
                throw new Exception($"Unknown proxy server type : ({type})");
            writer.WriteValue(conv_value);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            int type = ConstEnv.PROXY_TYPE_HTTPS;

            if (reader.TokenType == JsonToken.Null)
            {
                throw new Exception($"Unknown proxy server type : ({reader.TokenType})");
            }
            else if (reader.TokenType == JsonToken.String)
            {
                string value = serializer.Deserialize(reader, Type.GetType("string")) as string;
                value = value.ToLower();
                if (value == "https")
                    type = ConstEnv.PROXY_TYPE_HTTPS;
                else if (value == "socks4")
                    type = ConstEnv.PROXY_TYPE_SOCKS4;
                else if (value == "socsk5")
                    type = ConstEnv.PROXY_TYPE_SOCKS5;
                else
                    throw new Exception($"Unknown proxy server type : ({value})");
                return type;
            }
            else
            {
                throw new Exception($"Unknown proxy server type : ({reader.TokenType})");
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
