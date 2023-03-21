using MailParser;
using Logger;
using MailHelper;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace WebAuto
{
    public class KWebGCSpread : KWebBase
    {
        public KWebGCSpread() : base()
        {

        }

        protected override bool check_valid_link(string link)
        {
            if (link.StartsWith("https://www.giftcardspread.com/"))
                return true;

            return false;
        }
        private async Task<bool> BypassImgCaptcha()
        {
            try
            {
                string strCaptchaImgXpath = "//img[@alt='Captcha']";
                string strCaptchaInputXpath = "//input[@id='vcaptcha' and @type='text']";

                Screenshot screenshot = ((ITakesScreenshot)WebDriver).GetScreenshot();

                //screenshot.SaveAsFile("captcha.png", ScreenshotImageFormat.Png);

                Image img;
                byte[] imageBytes = screenshot.AsByteArray;
                using (var ms2 = new MemoryStream(imageBytes, 0, imageBytes.Length))
                {
                    img = Image.FromStream(ms2, true);
                }

                Rectangle rect = new Rectangle();

                IWebElement webelemCaptchaImg = WebDriver.FindElementsByXPath(strCaptchaImgXpath)[1];
                Point p = webelemCaptchaImg.Location;
                rect = new Rectangle(p.X, p.Y, webelemCaptchaImg.Size.Width, webelemCaptchaImg.Size.Height);

                Bitmap bmpImage = new Bitmap(img);
                Bitmap cropedImag = bmpImage.Clone(rect, bmpImage.PixelFormat);
                //MyLogger.Info($"String of croped image - {cropedImag.ToString()}");
                //cropedImag.Save("cropped.png", ImageFormat.Png);

                System.IO.MemoryStream ms1 = new MemoryStream();
                cropedImag.Save(ms1, ImageFormat.Png);
                byte[] byteImage = ms1.ToArray();
                string str_base64 = Convert.ToBase64String(byteImage);
                MyLogger.Info($"Base64 string of img - {str_base64}");

                Dictionary<string, string> dict = new Dictionary<string, string>();
                dict.Add("method", "base64");
                dict.Add("key", Program.g_setting.captcha_api);
                dict.Add("body", str_base64);
                string id = await Get_Id_from_2captcha(dict);
                MyLogger.Info($"Captcha id - {id}");

                if (Program.g_must_end)
                {
                    return false;
                }

                string strCaptcha = await Get_captcha_string_from_Id(id);
                MyLogger.Info($"Captcha string - {strCaptcha}");

                //Driver.FindElementByXPath(strCaptchaInputXpath).SendKeys(strCaptcha);
                await TryEnterText_by_xpath(strCaptchaInputXpath, strCaptcha, "value", 3000, false);

                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
            }
            return false;
        }

        protected async Task<int> scrap(string link, string order_id, string site_user, string site_password, List<ZGiftCardDetails> card_details)
        {
            int scrap_status = ConstEnv.SCRAP_FAILED;
            bool is_browser_opened = false;

            string strFirstUrl = "https://www.giftcardspread.com";
            string strOrderUrl = "https://www.giftcardspread.com/Voucher/PrintAllVouchers?OrderNo=" + order_id;

            try
            {
                if (!check_valid_link(link))
                    return ConstEnv.SCRAP_UNSUPPORTED;

                if (!await open_browser())
                    return ConstEnv.SCRAP_FAILED;
                MyLogger.Info("GiftCardSpread browser started successfully.");
                is_browser_opened = true;

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                if (!await Navigate(strFirstUrl))
                    throw new KScrapException($"Navi to first page url of GC spread failed - {strFirstUrl}");

                if (!await TryClick(By.PartialLinkText("Login"), 0))
                {
                    MyLogger.Error("Login button click failed in mode 0.");
                    if (!await TryClick(By.PartialLinkText("Login"), 1))
                    {
                        MyLogger.Error("Login button click failed in mode 1.");
                        if (!await TryClick(By.PartialLinkText("Login"), 2))
                        {
                            throw new KScrapException("Login button click failed in mode 2.");
                        }
                    }
                }

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                await TaskDelay(5000);

                string strAddrXpath = "//input[@id='exampleInputName' and @name='UserName']";
                string strPassXpath = "//input[@id='InputPassword1' and @name='Password']";

                if (!await WaitToPresentByPath(strAddrXpath, 5000))
                    throw new KScrapException($"Input on login page is not appeared.");

                IWebElement webelemAddr = WebDriver.FindElementsByXPath(strAddrXpath)[1];
                IWebElement webelemPass = WebDriver.FindElementsByXPath(strPassXpath)[1];

                webelemAddr.SendKeys(site_user);
                webelemPass.SendKeys(site_password);

                int retry_num = 0;
                while (true)
                {
                    bool bCaptchaResult = await BypassImgCaptcha();
                    if (!bCaptchaResult)
                    {
                        if (!await BypassImgCaptcha())
                            throw new KScrapException("Bypass captcha failed.");
                    }

                    string strSubmitButtonXpath = "//button[@id='btnSubmit' and @type='submit']";
                    await TryClick_All(strSubmitButtonXpath);
                    //await TryClickByPath(strSubmitButtonXpath, 1);

                    string acc_url = "https://www.giftcardspread.com/MyAccount";

                    if (!await WaitUrlSame(acc_url))
                        MyLogger.Error($"After submit, correct page is not appeared.");
                    else
                        break;
                    retry_num++;
                    if (retry_num > 5)
                        throw new KScrapException("Bypass captcha failed.");
                }

                if (!await Navigate(strOrderUrl))
                    throw new KScrapException($"Navi to order url failed - {strOrderUrl}");

                string strXpathItem = "//div[@class='pageBreak']";

                if (!await WaitToPresentByPath(strXpathItem, 7000))
                {
                    throw new KScrapException($"Card infos are not appeared. URL - {strOrderUrl}");
                }

                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> elems = WebDriver.FindElementsByXPath(strXpathItem);

                string strPin = string.Empty;
                string strCardNum = string.Empty;
                string strMechant = string.Empty;
                string strValue = string.Empty;

                foreach (IWebElement elem in elems)
                {
                    string strXpathPin = ".//span[@data-bind='text: Pin']";
                    strPin = get_elem_value(elem, strXpathPin);
                    MyLogger.Info($"Pin - {strPin}");

                    string strXpathCardNum = ".//span[@data-bind='text: CardNo']";
                    strCardNum = get_elem_value(elem, strXpathCardNum);
                    MyLogger.Info($"Card Num - {strCardNum}");

                    string strXpathValue = ".//div[@class='barcode-box']//h1//span";
                    string strValueTemp = get_elem_value(elem, strXpathValue);
                    strValueTemp = strValueTemp.Trim();
                    MyLogger.Info($"ValueTemp - {strValueTemp}");

                    strMechant = strValueTemp.Substring(0, strValueTemp.IndexOf("$"));
                    MyLogger.Info($"Merchant - {strMechant}");
                    strValue = strValueTemp.Substring(strValueTemp.IndexOf("$"), strValueTemp.Length - strValueTemp.IndexOf("$"));
                    MyLogger.Info($"Value - {strValue}");

                    card_details.Add(new ZGiftCardDetails(strMechant, Str_Utils.string_to_currency(strValue), 0, strCardNum, strPin));
                }

                scrap_status = ConstEnv.SCRAP_SUCCESS;
            }
            catch (KScrapException exception)
            {
                MyLogger.Error($"Scrap Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
            }
            finally
            {
                if (is_browser_opened)
                {
                    await Quit_browser();
                }
            }

            if (scrap_status == ConstEnv.SCRAP_FAILED && Program.g_must_end)
                scrap_status = ConstEnv.SCRAP_CANCELED;

            return scrap_status;
        }
        public override async Task<int> scrap_link(ZScrapParam param)
        {
            List<ZGiftCardDetails> card_details = new List<ZGiftCardDetails>();
            KReportCR report = param.report as KReportCR;

            string site_user = "";
            string site_password = "";
            Program.g_user.get_giftspread_account_info(param.report.m_mail_account_id, out site_user, out site_password);

            int scrap_status = await scrap(param.link, report.m_order_id, site_user, site_password, card_details);
            if (scrap_status == ConstEnv.SCRAP_SUCCESS)
            {
                lock (param.report.m_giftcard_details)
                {
                    foreach (ZGiftCardDetails data in card_details)
                    {
                        report.add_giftcard_details(data);
                    }
                }
            }
            return scrap_status;
        }
    }
}
