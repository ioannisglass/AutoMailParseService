using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebAuto
{
    public class KWebValueSears : KWebBase
    {
        public KWebValueSears() : base()
        {

        }

        protected override bool check_valid_link(string link)
        {
            if (link.StartsWith("https://value.sears.com/"))
                return true;

            return false;
        }

        protected async Task<int> scrap(string link, List<ZGiftCardDetails_V2> card_details)
        {
            int scrap_status = ConstEnv.SCRAP_FAILED;
            bool is_browser_opened = false;

            try
            {
                if (!check_valid_link(link))
                    return ConstEnv.SCRAP_UNSUPPORTED;

                if (!await open_browser())
                    return ConstEnv.SCRAP_FAILED;

                MyLogger.Info("ValueSears browser started successfully.");
                is_browser_opened = true;

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                if (!await Navigate(link))
                    throw new KScrapException($"Navi to blue button link failed - {link}");

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strInfoXpath = "//div[@class='gc_print_account_info']";

                if (!await WaitToPresentByPath(strInfoXpath, 5000))
                    throw new KScrapException("Info is not appeared.");

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strInfo = await get_value(strInfoXpath);
                MyLogger.Info($"Info - {strInfo}");

                string strGC = "";
                string strPin = "";

                int nIdxGC = strInfo.IndexOf("Account number:") + "Account number:".Length;
                string temp = strInfo.Substring(nIdxGC);
                for (int i = 0; i < temp.Length; i++)
                {
                    if (Char.IsDigit(temp[i]))
                        strGC += temp[i];
                    else
                        break;
                }
                MyLogger.Info($"Gift card - {strGC}");

                int nIdxPin = strInfo.IndexOf("Pin:") + "Pin:".Length;
                temp = strInfo.Substring(nIdxPin);
                for (int i = 0; i < temp.Length; i++)
                {
                    if (Char.IsDigit(temp[i]))
                        strPin += temp[i];
                    else
                        break;
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

            int scrap_status = await scrap(param.link, card_details);
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
