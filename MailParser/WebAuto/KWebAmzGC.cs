using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebAuto;

namespace WebAuto
{
    public class KWebAmzGC : KWebBase
    {

        public KWebAmzGC() : base()
        {
        }

        protected override bool check_valid_link(string link)
        {
            if (link.StartsWith("https://www.amazon.com/gp/") ||
                link.StartsWith("https://www.amazon.com/gp/"))
                return true;

            return false;
        }
        protected async Task<int> scrap(string link, List<ZGiftCardDetails_V2> card_details)
        {
            int scrap_status = ConstEnv.SCRAP_FAILED;
            bool is_browser_opened = false;

            string strXpathCardNum = "//span[@id='cardNumber2']";
            string strXpathPin = "//span[@id='Span2']";

            try
            {
                if (!check_valid_link(link))
                    return ConstEnv.SCRAP_UNSUPPORTED;
                MyLogger.Info($"AMZ-GC link is valid - {link}");

                if (!await open_browser())
                    return ConstEnv.SCRAP_FAILED;
                MyLogger.Info("AMZ-GC browser started successfully.");
                is_browser_opened = true;

                if (!await Navigate(link))
                {
                    throw new KScrapException($"Navi to {link} failed.");
                }
                MyLogger.Info($"AMZ-GC link navigation success - {link}");

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                if (!await WaitToPresentByPath(strXpathCardNum, 5000))
                {
                    throw new KScrapException($"Card Number is not appeared.");
                }

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                string strCardNum = await get_value(strXpathCardNum);
                MyLogger.Info($"Card Number - {strCardNum}");

                string strPin = await get_value(strXpathPin);
                MyLogger.Info($"Pin - {strPin}");

                card_details.Add(new ZGiftCardDetails_V2(strCardNum, strPin));

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
