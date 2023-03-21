using MailParser;
using Logger;
using MailHelper;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebAuto
{
    public class KWebHomedepot : KWebBase
    {
        public KWebHomedepot() : base()
        {

        }

        protected override bool check_valid_link(string link)
        {
            if (link.StartsWith("https://homedepot.cashstar.com/"))
                return true;

            return false;
        }

        protected async Task<int> scrap(string link, string mailto, List<ZGiftCardDetails_V2> card_details)
        {
            int scrap_status = ConstEnv.SCRAP_FAILED;
            bool is_browser_opened = false;

            string strGC = "";
            string strPin = "";

            try
            {
                if (!check_valid_link(link))
                    return ConstEnv.SCRAP_UNSUPPORTED;

                if (!await open_browser())
                    return ConstEnv.SCRAP_FAILED;

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                MyLogger.Info("Homedepot browser started successfully.");
                is_browser_opened = true;

                if (!await Navigate(link))
                    throw new KScrapException($"Navi to link failed - {link}");

                string strInputXpath = "//input[@name='value' and @id='id_value']";

                if (!await WaitToPresentByPath(strInputXpath, 5000))
                    throw new KScrapException("Input is not appeared.");

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strMail = mailto;

                await TryEnterText_by_xpath(strInputXpath, strMail, "value", 5000, true);

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strButton = "//button[@id='btn-challenge-screen']";
                await TryClickByPath(strButton, 1);

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strInfoItemXpath = "//p[@id='barcode-num']//span";

                if (!await WaitToPresentByPath(strInfoItemXpath, 7000))
                    throw new KScrapException($"Span tag is not appeared.");

                MyLogger.Info($"Span tags count - {WebDriver.FindElementsByXPath(strInfoItemXpath).Count}");

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                foreach (IWebElement sub_span in WebDriver.FindElementsByXPath(strInfoItemXpath))
                {
                    if (sub_span.GetAttribute("class") == "" || sub_span.GetAttribute("class") == "N/A")
                    {
                        strGC += sub_span.Text;
                    }
                }
                MyLogger.Info($"Gift card - {strGC}");

                string strInfoXpath = "//p[@id='barcode-num']";
                string strInfo = await get_value(strInfoXpath);

                string temp = strInfo.Substring(strInfo.IndexOf("PIN:") + "PIN:".Length);
                bool flag = false;
                for (int i = 0; i < temp.Length; i++)
                {
                    if (Char.IsDigit(temp[i]))
                    {
                        strPin += temp[i];
                        flag = true;
                    }
                    else
                    {
                        if (flag)
                            break;
                    }
                }
                MyLogger.Info($"Pin - {strPin}");

                card_details.Add(new ZGiftCardDetails_V2(strGC, strPin));

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
            List<ZGiftCardDetails_V2> card_details = new List<ZGiftCardDetails_V2>();

            int scrap_status = await scrap(param.link, param.mailto, card_details);
            if (scrap_status == ConstEnv.SCRAP_SUCCESS)
            {
                KReportCR report = param.report as KReportCR;
                lock (param.report.m_giftcard_details_v2)
                {
                    foreach (ZGiftCardDetails_V2 data in card_details)
                    {
                        report.add_giftcard_details_v2(data);
                    }
                }
            }
            return scrap_status;
        }
    }
}
