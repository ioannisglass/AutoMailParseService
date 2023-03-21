using Newtonsoft.Json;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Utils;
using GiftCard;
using Logger;
using MailParser;

namespace WebHelper
{
    partial class IWebHelper
    {
        public async Task<string> Get_site_key()
        {
            string site_key = "";
            try
            {
//                 foreach (IWebElement elem in WebDriver.FindElementsByTagName("div"))
//                 {
//                     site_key = elem.GetAttribute("data-sitekey");
//                     if (site_key != null)
//                     {
//                         MyLogger.Info($"Site key is found from div tags : {site_key}");
//                         return site_key;
//                     }                        
//                 }

                /*string frame_name = Get_self_name();
                MyLogger.Info($"In Get site key, the frame name is : {frame_name}");
                if (frame_name != "recaptcha")
                    WebDriver.SwitchTo().Frame("recaptcha");*/

                int retry_num = 0;
                while (site_key == "" && retry_num <= 30)
                {
                    retry_num += 1;
                    foreach (IWebElement elem in WebDriver.FindElementsByTagName("iframe"))
                    {
                        try
                        {
                            MyLogger.Info($"Checking iframe tags.");
                            int pos = elem.GetAttribute("src").IndexOf("https://www.google.com/recaptcha/api2/anchor?ar=1&k=");
                            if (pos > -1)
                            {                                
                                site_key = elem.GetAttribute("src").Substring("https://www.google.com/recaptcha/api2/anchor?ar=1&k=".Length);
                                int pos_1 = site_key.IndexOf('&');
                                site_key = site_key.Substring(0, pos_1);
                                MyLogger.Info($"site key : {site_key}");
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    await TaskDelay(1000);
                }
            }
            catch (Exception e)
            {
                MyLogger.Error($"Error catched : {e.Message}");
            }
            return site_key;
        }
        public async Task<string> Get_ID_from_site_key(string site_key)
        {
            try
            {
                string page_url = WebDriver.Url;
            
                string in_url = $"https://2captcha.com/in.php?key={Program.g_setting.captcha_api}&method=userrecaptcha&googlekey={site_key}&pageurl={page_url}";
                var w = new WebClient();
                string response_string = w.DownloadString(in_url);

                string[] fields = response_string.Split('|');
                string id = "";
                id = fields[1];
                return id;
            }
            catch (Exception e)
            {
                MyLogger.Error($"Error catched : {e.Message}");
            }
            return null;
        }
        public async Task<string> Get_Id(Dictionary<string, string> dict)
        {
            try
            {
                string in_url = "https://2captcha.com/in.php";

                HttpWebRequest request;
                request = (HttpWebRequest)HttpWebRequest.Create(in_url);

                var json_data = JsonConvert.SerializeObject(dict);

                request.Method = "POST";
                request.ContentType = "application/json";
                request.ContentLength = json_data.Length;                
                request.UserAgent = Str_Utils.GetRandomUserAgent();
                request.Accept = "*/*";

                var bytes_data = Encoding.ASCII.GetBytes(json_data);

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(bytes_data, 0, bytes_data.Length);
                }

                HttpWebResponse response = null;
                response = (HttpWebResponse)request.GetResponse();
                string response_string = new StreamReader(response.GetResponseStream()).ReadToEnd();

                string[] fields = response_string.Split('|');
                string id = "";
                id = fields[1];
                return id;
            }
            catch (Exception ex)
            {
                MyLogger.Info($"Error catched : {ex.Message}");
                return null;
            }
        }
        public async Task<string> Get_captcha_string_from_Id(string id)
        {
            try
            {
                string str_data_res = "";
                var w = new WebClient();
                w.Headers.Add(HttpRequestHeader.UserAgent, Str_Utils.GetRandomUserAgent());
                
                string res_url = $"https://2captcha.com/res.php?key={Program.g_setting.captcha_api}&action=get&id=";

                Stopwatch sub_wt = new Stopwatch();
                sub_wt.Start();
                bool flag = true;
                do
                {
                    MyLogger.Info($"Thread #{m_ID} - response started.");
                    if (flag == true)
                    {
                        await TaskDelay(15000);
                        flag = false;
                    }
                    str_data_res = w.DownloadString(res_url + id);
                    MyLogger.Info($"Captcha response - {str_data_res}");

                    int pos = Array.IndexOf(Program.g_setting.error_res_list, str_data_res);
                    if (pos > -1)
                    {
                        await TaskDelay(5000);
                    }
                    else
                        break;
                } while (sub_wt.ElapsedMilliseconds < 120000);

                string[] res_fields = str_data_res.Split('|');
                string captcha_string = res_fields[1];

                return captcha_string;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched in Get_captcha_string_from_Id : {ex.Message}");
                return null;
            }
        }
        public async Task<bool> Try_google_captcha()
        {
            try
            {
                string site_key = await Get_site_key();

                if (site_key == "")
                    return false;

                string id = await Get_ID_from_site_key(site_key);
                MyLogger.Info($"Captcha ID : {id}");

                string captcha_Str = await Get_captcha_string_from_Id(id);
                MyLogger.Info($"Captcha string : {captcha_Str}");

                /*string frame_name = Get_self_name();
                if (frame_name != "recaptcha")
                    WebDriver.SwitchTo().Frame("recaptcha");*/

                string script = $"document.getElementById('g-recaptcha-response-3').innerHTML='{captcha_Str}'";
                m_js.ExecuteScript(script);
                
                m_js.ExecuteScript(string.Format("{0}(\"{1}\");", "onCardCaptchaCallback", captcha_Str));
                
                return true;
            }
            catch (Exception e)
            {
                MyLogger.Error($"Error catched in Try_google_captcha: {e.Message}");
            }
            return false;
        }


        public async Task<bool> TryCaptcha(int mode, By by, int TimeOut = 120000)
        {
            string xpath;
            string str_base64 = "";

            try
            {
                Dictionary<string, string> dict = new Dictionary<string, string>();

                if (mode == 1)
                {
                    string src_temp = WebDriver.FindElement(by).GetAttribute("src");
                    string[] srcs = src_temp.Split(',');
                    str_base64 = srcs[1];
                }
                else
                {
                    MyLogger.Info($"Thread #{m_ID} - started img founding.");

                    IWebElement element = WebDriver.FindElement(by);
                    string cap_img_url = element.GetAttribute("src");

                    MyLogger.Info($"Thread #{m_ID} - img found.");

                    str_base64 = await Captcha_to_Base64(cap_img_url);

                    MyLogger.Info($"Thread #{m_ID} - {str_base64}");
                }

                if (Program.g_must_end)
                {
                    return false;
                }

                dict.Add("method", "base64");
                dict.Add("key", "86f75378dfe8e6bfad50e5ea37c61182");
                dict.Add("body", str_base64);

                MyLogger.Info($"Thread #{m_ID} - dict added.");

                string id = await Get_Id_from_2captcha(dict);
                MyLogger.Info($"Thread #{m_ID} - id = {id}.");

                if (Program.g_must_end)
                {
                    return false;
                }
                string captcha_string = await Get_captcha_string_from_Id(id);
                MyLogger.Info($"Thread #{m_ID} - captcha = {captcha_string}.");

                if (Program.g_must_end)
                {
                    return false;
                }

                if (mode == 1)
                {
                    xpath = "//input[@data-test='captcha-input' and @id='captcha']";
                    await TryEnterText_by_xpath(xpath, captcha_string, "value", Program.g_setting.delay_time * 1000, true);
                    return true;
                }
                else
                {
                    await TryEnterText("captcha_input", captcha_string, "value", Program.g_setting.delay_time * 1000, true);
                    xpath = "//button[@name='chapter:chapter_body:bottomButtons:container:bottomButtons_body:ok']";
                    await TaskDelay(2000);
                    await TryClickByPath(xpath, 1);
                    return true;
                }
            }
            catch (Exception e)
            {
                MyLogger.Error($"#{m_ID} - {e.Message} - error occured in captcha.");
            }
            return false;
        }

        public async Task<string> Captcha_to_Base64(string url)
        {
            OpenNewTab(url);
            WebDriver.SwitchTo().Window(WebDriver.WindowHandles.Last());

            Screenshot sh = WebDriver.GetScreenshot();
            IWebElement element = WebDriver.FindElementByTagName("img");

            Bitmap captcha;
            using (var ms = new MemoryStream(sh.AsByteArray))
            {
                Bitmap screenBitmap;
                screenBitmap = new Bitmap(ms);
                captcha = screenBitmap.Clone(
                    new Rectangle(
                        element.Location.X,
                        element.Location.Y,
                        element.Size.Width,
                        element.Size.Height
                        ), screenBitmap.PixelFormat);
            }
            WebDriver.Close();
            WebDriver.SwitchTo().Window(WebDriver.WindowHandles.First());
            /*if (frame_name != "")
                WebDriver.SwitchTo().Frame(frame_name);*/
            //captcha.Save("cap.png");

            System.IO.MemoryStream ms1 = new MemoryStream();
            captcha.Save(ms1, ImageFormat.Png);
            byte[] byteImage = ms1.ToArray();
            string str_base64 = Convert.ToBase64String(byteImage);

            return str_base64;
        }

        public async Task<string> Get_Id_from_2captcha(Dictionary<string, string> dict)
        {
            string id = "";
            try
            {
                Stopwatch sub_wt = new Stopwatch();
                sub_wt.Start();
                do
                {
                    string in_url = "https://2captcha.com/in.php";

                    HttpWebRequest request;
                    request = (HttpWebRequest)HttpWebRequest.Create(in_url);

                    var json_data = JsonConvert.SerializeObject(dict);

                    request.Method = "POST";
                    request.ContentType = "application/json";
                    request.ContentLength = json_data.Length;
                    request.UserAgent = Str_Utils.GetRandomUserAgent();
                    request.Accept = "*/*";

                    var bytes_data = Encoding.ASCII.GetBytes(json_data);

                    using (var stream = request.GetRequestStream())
                    {
                        stream.Write(bytes_data, 0, bytes_data.Length);
                    }

                    HttpWebResponse response = null;
                    response = (HttpWebResponse)request.GetResponse();
                    string response_string = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    MyLogger.Info($"ID Response - {response_string}");

                    try
                    {
                        string[] fields = response_string.Split('|');
                        id = fields[1];
                        MyLogger.Info($"ID - {id}");
                        break;
                    }
                    catch(Exception ex)
                    {
                        await TaskDelay(5000);
                    }                 
                } while (sub_wt.ElapsedMilliseconds < 60000);                
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched : {ex.Message}");
            }
            return id;
        }
    }
}
