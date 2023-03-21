using MailParser;
using Logger;
using MailHelper;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace WebAuto
{
    public class KWebRMN : KWebBase
    {
        public KWebRMN() : base()
        {

        }

        protected override bool check_valid_link(string link)
        {
            if (link.StartsWith("https://www.retailmenot.com/"))
                return true;

            return false;
        }

        protected async Task<int> scrap(string link, KReportCR report)
        {
            int scrap_status = ConstEnv.SCRAP_FAILED;
            bool is_browser_opened = false;

            string strXpathCardItem = "//div[@class='claim-card-container']//a[@class='claim-card-button button-primary']";

            try
            {
                if (!check_valid_link(link))
                    return ConstEnv.SCRAP_UNSUPPORTED;

                if (!await open_browser())
                    return ConstEnv.SCRAP_FAILED;

                MyLogger.Info("RMN browser started successfully.");
                is_browser_opened = true;

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                if (!await Navigate(link))
                    throw new KScrapException("Navi to first link button url failed.");
                MyLogger.Info("Navi to first url success.");

                if (!await WaitToPresentByPath(strXpathCardItem, 15000))
                {
                    string strXpathRule = "//div[@class='special-rules display-special-rules']//p";
                    string strInfo = WebDriver.FindElementByXPath(strXpathRule).Text.Trim();
                    string strKey = "Card numbers and pins are not shown because this order is over 100 days old.";

                    MyLogger.Info($"Rule info - {strInfo}.");
                    if (strInfo.Contains(strKey))
                    {
                        scrap_status = ConstEnv.SCRAP_OTHER;

                        string strOrderXpath = "//span[@class='claim-order-number']";
                        string orderID = await get_value(strOrderXpath);
                        report.set_order_id(orderID);
                        throw new KScrapException($"{orderID}:100 days old.");
                    }
                    else
                    {
                        throw new KScrapException("Card items are not appeared.");
                    }
                }

                if (Program.g_must_end)
                    throw new KScrapException($"Must end.");

                MyLogger.Info($"Card numbers - {WebDriver.FindElementsByXPath(strXpathCardItem).Count}");
                int count = WebDriver.FindElementsByXPath(strXpathCardItem).Count;

                for (int i = 0; i < count; i++)
                {
                    string xpath = $"//div[@class='claim-card-container']//a[@class='claim-card-button button-primary' and @data-index='{i}']";

                    if (Program.g_must_end)
                        throw new KScrapException($"Must end.");

                    if (!await WaitToPresentByPath(xpath, 5000))
                        throw new KScrapException("Click button is not appeared.");

                    if (!await TryClickByPath(xpath, 2))
                        if (!await TryClick_All(xpath))
                            throw new KScrapException("Click items failed.");

                    if (Program.g_must_end)
                        throw new KScrapException($"Must end.");

                    if (!await WaitToPresentByPath("//img[@id='barcode' and @class='claim-barcode']", 5000))
                        throw new Exception("Bar code is not appeared.");

                    int counts = 0;
                    string strOrder = "";
                    string strValue = "";
                    string strRetailer = "";
                    string strCardNumber = "";
                    string strPin = "";

                    while (true)
                    {
                        counts++;
                        if (counts > 2)
                            break;

                        strOrder = await get_value("//span[@class='claim-order-number']");
                        MyLogger.Info($"Order Number - {strOrder}");
                        if (strOrder == "" || strOrder == "N/A")
                        {
                            if (Program.g_must_end)
                                throw new KScrapException($"Must end.");

                            await TaskDelay(5000);
                            //System.Threading.Thread.Sleep(5000);
                            continue;
                        }

                        report.set_order_id(strOrder);
                        strValue = await get_value("//div[@class='claim-card-title-amount']");
                        strValue = strValue.Trim();
                        if (strValue.StartsWith("$"))
                            strValue = strValue.Substring(1);
                        MyLogger.Info($"Value - {strValue}");

                        strRetailer = await get_value("//span[@class='claim-card-title-merchant']");
                        MyLogger.Info($"Retailer - {strRetailer}");

                        strCardNumber = await get_value("//div[@class='card-number']");
                        MyLogger.Info($"Card Number - {strCardNumber}");

                        strPin = await get_value("//div[@class='pin']");
                        MyLogger.Info($"Pin - {strPin}");

                        report.add_giftcard_details(new ZGiftCardDetails(strRetailer, Str_Utils.string_to_currency(strValue), 0, strCardNumber, strPin));

                        break;
                    }

                    WebDriver.Navigate().Back();
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

            string order;
            int scrap_status = await scrap(param.link, param.report as KReportCR);
            return scrap_status;
        }
    }
}
