using GiftCard;
using Logger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebHelper
{
    partial class IWebHelper
    {
        public string Get_self_name()
        {
            string name = (string)m_js.ExecuteScript("return self.name");
            return name;
        }
        public void close_last_window(string msg = "")
        {
            WebDriver.SwitchTo().Window(WebDriver.WindowHandles.Last());
            WebDriver.Close();
            WebDriver.SwitchTo().Window(WebDriver.WindowHandles.First());
            if (msg != "")
            {
                m_js.ExecuteScript($"alert({msg});");
            }
        }
        public async Task<bool> Navigate(string target)
        {
            try
            {
                string url = WebDriver.Url;
                WebDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
                //WebDriver.Manage().Timeouts().ImplicitWait(TimeSpan.FromSeconds(120));
                WebDriver.Navigate().GoToUrl(target);
                return await WaitUrlChange(url);
            }
            catch (TimeoutException ex)
            {
                MyLogger.Error($"TimeoutException in Navigate - {ex.Message}");
                //ForceStopChrDriver();
                return false;
            }
            catch (OpenQA.Selenium.WebDriverException ex)
            {
                MyLogger.Error($"WebDriverException in Navigate - {ex.Message}");
                //ForceStopChrDriver();
                return false;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched in Navigate - {ex.Message}");
                //await Quit();
                return false;
            }
        }
        public void OpenNewTab(string url)
        {
            m_js.ExecuteScript(string.Format("window.open('{0}', '_blank');", url));
        }
        public void NewTab(string tabUrl)
        {
            lock (m_locker)
            {
                string newTabScript = "var d=document,a=d.createElement('a');"
                                + "a.target='_blank';a.href='{0}';"
                                + "a.innerHTML='new tab';"
                                + "d.body.appendChild(a);"
                                + "a.click();"
                                + "a.parentNode.removeChild(a);";
                if (String.IsNullOrEmpty(tabUrl))
                    tabUrl = "about:blank";

                m_js.ExecuteScript(String.Format(newTabScript, tabUrl));
            }
        }
    }
}
