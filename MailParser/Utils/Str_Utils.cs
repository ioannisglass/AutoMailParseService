using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    class Str_Utils
    {
        public static string Decode_base64(string base64_encoded_str)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64_encoded_str);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static string GetRandomUserAgent()
        {
            string[] strArray = new string[]
            {
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36",
                //"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36",
                "Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36",
                //"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36 OPR/43.0.2442.991",
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 YaBrowser/17.10.0.2052 Yowser/2.5 Safari/537.36",
                //"Mozilla/5.0 (Windows NT 5.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36 OPR/43.0.2442.991",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36",
                //"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.89 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36",
                "Mozilla/5.0 (compatible; U; ABrowse 0.6; Syllable) AppleWebKit/420+ (KHTML, like Gecko)",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.75 Safari/537.36 OPR/36.0.2130.32",
                "Opera/9.80 (Windows NT 6.1; WOW64) Presto/2.12.388 Version/12.18",
                "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36 OPR/43.0.2442.991",
                "Opera/9.80 (Windows NT 6.0) Presto/2.12.388 Version/12.14",
                "Opera/9.80 (Windows NT 5.1; WOW64) Presto/2.12.388 Version/12.17"
            };
            return strArray[new Random().Next(0, strArray.Length - 1)];
        }
        public static string CleanPath(string path)
        {
            string clean_path = path;
            clean_path = clean_path.Replace('/', Path.DirectorySeparatorChar);
            clean_path = clean_path.Replace('\\', Path.DirectorySeparatorChar);
            return string.Join("_", clean_path.Split(Path.GetInvalidPathChars()));
        }

        public static float string_to_currency(string _src)
        {
            string src = _src.Trim();
            float f = 0;
            int is_minus = 1;

            try
            {
                if (src[0] == 'C')
                    src = src.Substring(1).Trim();
                if (src[0] == '-')
                {
                    src = src.Substring(1).Trim();
                    is_minus = -1;
                }
                if (src[0] == '$')
                    src = src.Substring(1).Trim();
                src = src.Trim();
                f = string_to_float(src);
                f *= is_minus;
            }
            catch (Exception)
            {

            }

            return f;
        }
        public static float string_to_float(string src)
        {
            float f = 0;
            if (src == "")
                return 0;

            try
            {
                f = float.Parse(src);
            }
            catch (Exception)
            {
                string float_charset = "-.,0123456789";
                int i;
                int dot_pos = -1;
                for (i = 0; i < src.Length; i++)
                {
                    if (src[i] == '-' && i != 0)
                        break;
                    if (src[i] == '.')
                    {
                        if (dot_pos != -1)
                            break;
                        dot_pos = i;
                    }
                    if (!float_charset.Contains(src[i]))
                        break;
                }
                string src1 = src.Substring(0, i);
                f = float.Parse(src1);
            }
            return f;
        }
        public static int string_to_int(string src)
        {
            int n = 0;
            if (src == "")
                return 0;
            try
            {
                n = int.Parse(src);
            }
            catch (Exception)
            {
                string integer_charset = "-,0123456789";
                int i;
                for (i = 0; i < src.Length; i++)
                {
                    if (src[i] == '-' && i != 0)
                        break;
                    if (!integer_charset.Contains(src[i]))
                        break;
                }
                string src1 = src.Substring(0, i);
                n = int.Parse(src1);
            }
            return n;
        }
    }
}
