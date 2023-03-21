using BaseModule;
using MailParser;
using Logger;
using MailHelper;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebHelper;

namespace WebAuto
{
    public class KWebGCZen : IWebHelper
    {
        public List<GiftCardZenSales> m_lst_GCZ_Sales;
        public KWebGCZen()
        {
            m_lst_GCZ_Sales = new List<GiftCardZenSales>();
        }

        public async Task<int> scrap_link(string _mail, string _pass)
        {
            bool is_browser_opened = false;
            bool return_value = false;

            try
            {
                if (!await OpenBrowser())
                    return ConstEnv.SCRAP_FAILED;

                MyLogger.Info("GiftCardZen browser started successfully.");
                is_browser_opened = true;

                if (Program.g_must_end)
                    throw new Exception($"Must end.");

                int cap_retry_num = 0;
                while (true)
                {
                    if (cap_retry_num >= 3)
                        throw new Exception("Captcha 3 times failed.");

                    string strUrlSign = "https://giftcardzen.com/users/sign_in";
                    if (!await Navigate(strUrlSign))
                        throw new Exception($"# navi to sign in url {strUrlSign} failed.");

                    string strAddrXpath = "//input[@type='email' and @id='user_email']";
                    if (!await WaitToPresentByPath(strAddrXpath, 15000))
                        throw new Exception($"# Mail address input in sign in page is not appeared.");

                    if (!await TryEnterText_by_xpath(strAddrXpath, _mail, "value", 3000, true))
                    {
                        MyLogger.Error($"# Mail address input failed in sign in page in case of true.");
                        if (!await TryEnterText_by_xpath(strAddrXpath, _mail, "value", 3000, false))
                            throw new Exception($"# Mail address input failed in sign in page in case of false.");
                    }

                    string strPassXpath = "//input[@type='password' and @id='user_password']";
                    if (!await WaitToPresentByPath(strPassXpath, 15000))
                    {
                        throw new Exception($"# Password input sign in page is not appeared.");
                    }
                    if (!await TryEnterText_by_xpath(strPassXpath, _pass, "value", 3000, true))
                    {
                        MyLogger.Error($"# Mail password input failed in sign in page in case of true.");
                        if (!await TryEnterText_by_xpath(strPassXpath, _pass, "value", 3000, false))
                            throw new Exception($"# Mail password  input failed in sign in page in case of false.");
                    }

                    await BypassInsideCaptcha();

                    string strUrlFirst = "https://giftcardzen.com/a/sales?page=1";

                    if (!await Navigate(strUrlFirst))
                        MyLogger.Error($"# navi to first url {strUrlFirst} failed.");
                    else
                        break;
                    cap_retry_num++;
                }

                if (Program.g_must_end)
                    throw new Exception($"Must end.");

                int page_no = 1;
                while (true)
                {
                    MyLogger.Info($"Page No - {page_no}");
                    MyLogger.Info($"Page url - {WebDriver.Url}");
                    await ScrapPage();

                    //string str_xpath_last = "//*[@id='sales']/main/div/ul/li[9]/a";
                    string str_xpath_btn_next = "//a[@rel='next']";

                    if (await WaitToPresentByPath(str_xpath_btn_next, 3000))
                    {
                        await TryClickByPath(str_xpath_btn_next, 1);
                        MyLogger.Info("Next button is clicked.");
                    }
                    else
                    {
                        MyLogger.Info("Next button is not existed. End.");
                        break;
                    }
                    page_no++;
                }

                MyLogger.Info("Scrap link in GCZen finished successfully.");
                return_value = true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
            }
            finally
            {
                if(is_browser_opened)
                {
                    await Quit_In_Exception();
                }
            }
            return return_value ? ConstEnv.SCRAP_SUCCESS : ConstEnv.SCRAP_FAILED;
        }

        private async Task<bool> BypassInsideCaptcha()
        {
            try
            {
                string strLoginBtnXpath = "//button[@type='submit']";

                string strCaptchaSiteKey = await get_value(strLoginBtnXpath, "", "data-sitekey");
                MyLogger.Info($"Site key - {strCaptchaSiteKey}");

                string strCaptchaId = await Get_ID_from_site_key(strCaptchaSiteKey);
                MyLogger.Info($"Captcha ID - {strCaptchaId}");

                string strCaptcha = await Get_captcha_string_from_Id(strCaptchaId);
                MyLogger.Info($"Captcha response - {strCaptcha}");

                string script = $"document.getElementById('g-recaptcha-response').innerHTML='{strCaptcha}'";
                m_js.ExecuteScript(script);
                m_js.ExecuteScript(string.Format("{0}(\"{1}\");", "invisibleRecaptchaSubmit", strCaptcha));

                return true;
            }
            catch(Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
            }
            return false;
        }

        private async Task<bool> OpenBrowser()
        {
            try
            {
                if (ConstEnv.OS_TYPE == ConstEnv.OS_WINDOWS)
                {
                    if (!await Start())
                    {
                        MyLogger.Error("Chrome starting failed on Windows.");
                        return false;
                    }
                    MyLogger.Error("Chrome starting success on Windows.");
                }
                else if (ConstEnv.OS_TYPE == ConstEnv.OS_UNIX)
                {
                    if (!await start_headless())
                    {
                        MyLogger.Error("Chrome starting failed on Ubuntu");
                        return false;
                    }
                    MyLogger.Error("Chrome starting success on Ubuntu.");
                }
                else
                {
                    MyLogger.Error("Unknown OS type.");
                    return false;
                }

                MyLogger.Info($"Browser started successfully.");
                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
                return false;
            }
        }

        private async Task ScrapPage()
        {
            MyLogger.Info("Scrap page started.");
            string xpath_row = "//table[@class='table table-striped']//tbody//tr";

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> row_elems = WebDriver.FindElementsByXPath(xpath_row);
            if (row_elems == null)
                return;
            int row_count = row_elems.Count;

            for (int i = 0; i < row_count; i++)
            {
                await WaitToPresentByPath(xpath_row, 3000);
                row_elems = WebDriver.FindElementsByXPath(xpath_row);
                IWebElement row_elem = row_elems[i];
                GiftCardZenSales sales_item = await ScrapPageRow(row_elem);
                m_lst_GCZ_Sales.Add(sales_item);
            }
            MyLogger.Info("Scrap page finished.");
        }

        private async Task<GiftCardZenSales> ScrapPageRow(IWebElement _elem_row)
        {
            GiftCardZenSales sales_item = new GiftCardZenSales();

            string sales_id = string.Empty;
            string detail_url = string.Empty;
            string created = string.Empty;
            string paid = string.Empty;
            //string worth = string.Empty;
            //string contains = string.Empty;

            string xpath_id = ".//td[@class='id']//a";
            sales_id = get_elem_value(_elem_row, xpath_id).Trim();
            detail_url = get_elem_value(_elem_row, xpath_id, "", "href").Trim();

            string xpath_created = ".//td[@class='hide-small']";
            created = get_elem_value(_elem_row, xpath_created).Trim();

            string xpath_paid = ".//td[@class='amount']";
            paid = get_elem_value(_elem_row, xpath_paid).Trim();

            //string xpath_worth = ".//td[@class='value']";
            //worth = get_elem_value(_elem_row, xpath_worth).Trim();

            //string xpath_contains = ".//td[5]";
            //contains = get_elem_value(_elem_row, xpath_contains).Trim();

            await Navigate(detail_url);
            ScrapDetailPage(sales_item.lst_digi_cards);
            WebDriver.Navigate().Back();

            MyLogger.Info(sales_id + "\n" + created + "\n" + paid + "\n" /*+ worth + "\n" + contains + "\n"*/ + detail_url);
            MyLogger.Info("\n" + "######################################" + "\n");

            sales_item.order_ID = sales_id;
            sales_item.date_of_purchase = created;
            sales_item.total = paid;

            return sales_item;
        }

        private void ScrapDetailPage(List<GiftCardZenDigiCard> _lst_digi_card)
        {
            MyLogger.Info("Scrap detail page started.");
            string xpath_row = "//table[@class='cards table table-striped']//tbody//tr[@class='card-row']";

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> row_elems = WebDriver.FindElementsByXPath(xpath_row);

            int idx_card_num = 0;
            for (int i = 0; i < row_elems.Count; i++)
            {
                IWebElement row_elem = row_elems[i];
                GiftCardZenDigiCard sales_item = ScrapDetailPageRow(row_elem, idx_card_num);
                if (sales_item.status == string.Empty)
                    idx_card_num++;
                _lst_digi_card.Add(sales_item);
            }
            MyLogger.Info("Scrap detail page finished.");
        }

        private GiftCardZenDigiCard ScrapDetailPageRow(IWebElement _elem_row, int idx)
        {
            GiftCardZenDigiCard digi_card = new GiftCardZenDigiCard();

            string str_merchant = string.Empty;
            string str_card_value = string.Empty;
            string str_you_paid = string.Empty;
            string str_you_saved = string.Empty;
            string str_card_number = string.Empty;
            string str_pin = string.Empty;
            string str_status = string.Empty;

            string xpath_merchant = ".//td[@class='merchant-col']//span[@class='name']";
            str_merchant = get_elem_value(_elem_row, xpath_merchant).Trim();

            string xpath_card_value = ".//td[@class='value']";
            str_card_value = get_elem_value(_elem_row, xpath_card_value).Trim();

            string xpath_you_paid = ".//td[@class='pay']";
            str_you_paid = get_elem_value(_elem_row, xpath_you_paid).Trim();

            string xpath_you_saved = ".//td[@class='save']";
            str_you_saved = get_elem_value(_elem_row, xpath_you_saved).Trim();

            string xpath_status = ".//td[@class='status']//span[@class='success']";
            string xpath_action = ".//td[@class='status']//div[@class='actions']";
            if (get_elem_value(_elem_row, xpath_action, "ERROR") == "ERROR")
            {
                str_status = get_elem_value(_elem_row, xpath_status).Trim();
            }
            else
            {
                string xpath_barcode = "//table[@class='cards table table-striped']//tbody//tr[@class='barcode']";
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> barcode_elems = WebDriver.FindElementsByXPath(xpath_barcode);

                string xpath_card_number = ".//h3";
                string xpath_pin = ".//h4";
                str_card_number = get_elem_value(barcode_elems[idx], xpath_card_number);
                str_pin = get_elem_value(barcode_elems[idx], xpath_pin);

                if (str_pin.IndexOf("PIN:") != -1)
                {
                    string temp = str_pin.Substring(str_pin.IndexOf("PIN:") + "PIN:".Length);
                    str_pin = temp.Trim();
                }
            }

            MyLogger.Info(str_merchant + "\n" + str_card_value + "\n" + str_you_paid + "\n" + str_you_saved + "\n" + str_card_number + "\n" + str_pin);
            MyLogger.Info("\n" + "######################################" + "\n");

            digi_card.retailer = str_merchant;
            digi_card.value = str_card_value;
            digi_card.cost = str_you_paid;
            digi_card.discount = str_you_saved;
            digi_card.card_number = str_card_number;
            digi_card.pin = str_pin;
            digi_card.status = str_status;

            return digi_card;
        }
    }
}
