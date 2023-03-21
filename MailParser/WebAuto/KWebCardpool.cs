using MailParser;
using Logger;
using MailHelper;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebAuto;

namespace WebAuto
{
    public class KWebCardpool : KWebBase
    {

        public KWebCardpool() : base()
        {    
        }

        protected override bool check_valid_link(string link)
        {
            if (link.StartsWith("http://email.cardpool.com/") ||
                link.StartsWith("https://www.cardpool.com/"))
                return true;

            return false;
        }
        private bool get_cardnum_and_pin_from_htmltext(string htmltext, string retailer, float value, List<ZGiftCardDetails> card_details)
        {
            string strCardNum = "";
            string strPin = "";
            string temp = "";
            int pos;
            string[] cardnum_key = new string[] { "Card #:", "Card Number:" };
            string[] pin_key = new string[] { "Pin:", "Security Code (PIN):" };

            foreach (string k in cardnum_key)
            {
                pos = htmltext.IndexOf(k, StringComparison.CurrentCultureIgnoreCase);
                if (pos == -1)
                    continue;
                temp = htmltext.Substring(pos + k.Length);

                pos = temp.IndexOf("\n");
                if (pos != -1)
                    pos = temp.IndexOf("</div>");
                if (pos != -1)
                {
                    strCardNum = temp.Substring(0, pos).Trim();
                    strCardNum = XMailHelper.html2text(strCardNum);

                    foreach (string k1 in pin_key)
                    {
                        pos = strCardNum.IndexOf(k1, StringComparison.CurrentCultureIgnoreCase);
                        if (pos != -1)
                        {
                            strCardNum = strCardNum.Substring(0, pos).Trim();
                            break;
                        }
                    }

                    strCardNum = strCardNum.Replace("\r", "");
                    strCardNum = strCardNum.Replace("\n", "");
                    MyLogger.Info($"From Web text : Card Number - {strCardNum}");
                }
            }

            foreach (string k in pin_key)
            {
                pos = htmltext.IndexOf(k, StringComparison.CurrentCultureIgnoreCase);
                if (pos == -1)
                    continue;
                temp = htmltext.Substring(pos + k.Length);

                pos = temp.IndexOf("\n");
                if (pos != -1)
                    pos = temp.IndexOf("</div>");
                if (pos != -1)
                {
                    strPin = temp.Substring(0, pos).Trim();
                    strPin = XMailHelper.html2text(strPin);
                    strPin = strPin.Replace("\r", "");
                    strPin = strPin.Replace("\n", "");
                    MyLogger.Info($"From Web text : Pin - {strPin}");
                }
            }

            if (strCardNum != "" || strPin != "")
            {
                card_details.Add(new ZGiftCardDetails(retailer, value, 0, strCardNum, strPin));
                return true;
            }
            return false;
        }
        protected async Task<int> scrap(string link, string retailer, float value, List<ZGiftCardDetails> card_details)
        {
            int scrap_status = ConstEnv.SCRAP_FAILED;
            bool is_browser_opened = false;

            string strXpathCardNum = "//span[@id='cardNumber2']";
            string strXpathPin = "//span[@id='Span2']";

            try
            {
                if (!check_valid_link(link))
                    return ConstEnv.SCRAP_UNSUPPORTED;
                MyLogger.Info($"Cardpool link is valid - {link}");

                if (!await open_browser())
                    return ConstEnv.SCRAP_FAILED;
                MyLogger.Info("Cardpool browser started successfully.");
                is_browser_opened = true;

                if (!await Navigate(link))
                {
                    throw new KScrapException($"Navi to {link} failed.");
                }
                MyLogger.Info($"Cardpool link navigation success - {link}");

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strXpathCaptcha = "//div[@id='captcha_window']";
                if (await WaitToPresentByPath(strXpathCaptcha, 5000))
                {
                    MyLogger.Info($"Captcha appeared in cardpool link.");
                    await Try_google_captcha();
                    string strXpathSubmit = "//span[@class='button_text']";
                    if (!await TryClickByPath(strXpathSubmit, 1))
                    {
                        MyLogger.Info($"Submit button click failed.");
                        await TryClick_All(strXpathSubmit);
                    }
                }

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                if (!await WaitToPresentByPath(strXpathCardNum, 5000))
                {
                    string curr_url = WebDriver.Url.Trim();
                    if (curr_url[curr_url.Length - 1] == '/')
                        curr_url = curr_url.Substring(0, curr_url.Length - 1);
                    if (curr_url == "https://www.cardpool.com/sell-gift-cards" ||
                        curr_url == "https://www.cardpool.com")
                    {
                        scrap_status = ConstEnv.SCRAP_OTHER;
                        throw new KScrapException($"URL not working.");
                    }
                    else
                    {
                        MyLogger.Info("Card Number is not appeared.");

                        if (get_cardnum_and_pin_from_htmltext(WebDriver.PageSource, retailer, value, card_details))
                            scrap_status = ConstEnv.SCRAP_SUCCESS;
                        else
                            scrap_status = ConstEnv.SCRAP_FAILED;
                        return scrap_status;
                    }
                }

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                if (get_cardnum_and_pin_from_htmltext(WebDriver.PageSource, retailer, value, card_details))
                {
                    scrap_status = ConstEnv.SCRAP_SUCCESS;
                }
                else
                {
                    string strCardNum = "";
                    string strPin = "";

                    strCardNum = await get_value(strXpathCardNum);
                    MyLogger.Info($"Card Number - {strCardNum}");

                    strPin = await get_value(strXpathPin);
                    if (strPin == "")
                    {
                        try
                        {
                            strPin = WebDriver.FindElementsByXPath("//div[@id='cardInfo']//span")[1].Text;
                        }
                        catch (Exception ex1)
                        {
                            MyLogger.Info($"Exception Error to locate pin : message = {ex1.Message} ");
                        }
                    }
                    MyLogger.Info($"Pin - {strPin}");

                    card_details.Add(new ZGiftCardDetails(retailer, value, 0, strCardNum, strPin));

                    scrap_status = ConstEnv.SCRAP_SUCCESS;
                }
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

            int scrap_status = await scrap(param.link, param.retailer, param.value, card_details);
            if (scrap_status == ConstEnv.SCRAP_SUCCESS)
            {
                KReportCR report = param.report as KReportCR;
                lock (param.report.m_giftcard_details_v2)
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
