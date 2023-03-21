using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GiftCard;
using Logger;

namespace WebHelper
{
    partial class IWebHelper
    {
        public async Task<bool> TryClick(string classname, int mode = 0, int delay = 100)
        {
            try
            {
                await TryClick(By.ClassName(classname), mode);
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error in TryClick - {ex.Message}");
                return false;
            }
        }
        public async Task<bool> TryClickByPath(string xpath, int mode = 0, int delay = 100)
        {
            try
            {
                return await TryClick(By.XPath(xpath), mode);
                //return true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error in TryClickByPath - {ex.Message}");
                return false;
            }
        }
        public async Task<bool> TryClick(By by, int mode)
        {
            try
            {
                if (mode == 0)
                {
                    WebDriver.ExecuteScript("arguments[0].click('');", ((RemoteWebDriver)WebDriver).FindElement(by));
                }
                else if (mode == 1)
                {
                    WebDriver.FindElement(by).Click();
                }
                else if (mode == 2)
                {
                    Actions action = new Actions(WebDriver);
                    action.MoveToElement(WebDriver.FindElement(by)).Perform();
                    action.Click(WebDriver.FindElement(by)).Perform();
                }

                return true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched in TryClick of mode {mode.ToString()} - {ex.Message}");
            }
            return false;
        }
        public async Task<bool> TryClick_All(string xpath)
        {
            if (await TryClickByPath(xpath, 0))
                return true;
            if (await TryClickByPath(xpath, 1))
                return true;
            if (await TryClickByPath(xpath, 2))
                return true;
            bool ret = false;
            try
            {
                m_js.ExecuteAsyncScript($"document.evaluate('{xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");
                ret = true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched in TryClick_All when excute async script after trying all - {ex.Message}");
                ret = false;
            }

            if (ret == false)
                MyLogger.Error($"{m_ID} : Clicking all ways failed. {xpath}");
            return ret;
        }

        public async Task<bool> TryClickAndWait(string toClick, string toWait, int mode = 0, int TimeOut = 10000)
        {
            return await TryClickAndWait(By.XPath(toClick), By.XPath(toWait), mode, TimeOut);
        }
        public async Task<bool> TryClickAndWait(By toClick, By toWait, int mode = 0, int TimeOut = 10000)
        {
            if (!await WaitToPresent(toClick, 3000))
            {
                MyLogger.Error($"#{m_ID} - Element to be clicked is not present! mode:{mode} By: {toClick}");
                return false;
            }

            Stopwatch wt = new Stopwatch();
            wt.Start();
            while (wt.ElapsedMilliseconds < TimeOut)
            {
                try
                {
                    if (mode == 1)
                    {
                        string script = @"(function(x) {
                            var el = document.evaluate('" + toClick + @"', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                            let hoverEvent = document.createEvent ('MouseEvents');
                            hoverEvent.initEvent ('mouseover', true, true);
                            el.dispatchEvent (hoverEvent);

                            let downEvent = document.createEvent ('MouseEvents');
                            downEvent.initEvent ('mousedown', true, true);
                            el.dispatchEvent (downEvent);

                            let upEvent = document.createEvent ('MouseEvents');
                            upEvent.initEvent ('mouseup', true, true);
                            el.dispatchEvent (upEvent);

                            let clickEvent = document.createEvent ('MouseEvents');
                            clickEvent.initEvent ('click', true, true);
                            el.dispatchEvent (clickEvent);
                            })();";
                        WebDriver.ExecuteScript(script);
                        if (!await WaitToPresent(toWait, TimeOut))
                        {
                            MyLogger.Error($"#{m_ID} - Click failed for waiting! mode:{mode} By: {toClick}");
                            return false;
                        }
                        MyLogger.Error($"#{m_ID} - Click success! mode:{mode} By: {toClick}");
                        return true;
                    }
                    else if (mode == 0)
                    {
                        WebDriver.ExecuteScript("arguments[0].click('');", WebDriver.FindElement(toClick));
                        if (!await WaitToPresent(toWait, TimeOut))
                        {
                            MyLogger.Error($"#{m_ID} - Click failed for waiting! mode:{mode} By: {toClick}");
                            return false;
                        }

                        MyLogger.Error($"#{m_ID} - Click success! mode:{mode} By: {toClick}");
                        return true;
                    }
                    else if (mode == 2)
                    {
                        WebDriver.FindElement(toClick).Click();
                        if (!await WaitToPresent(toWait, TimeOut))
                        {
                            MyLogger.Error($"#{m_ID} - Click failed for waiting! mode:{mode} By: {toClick}");
                            return false;
                        }

                        MyLogger.Error($"#{m_ID} - Click success! mode:{mode} By: {toClick}");
                        return true;
                    }
                    else if (mode == 3)
                    {
                        Actions action = new Actions(WebDriver);
                        action.MoveToElement(WebDriver.FindElement(toClick)).Perform();
                        action.Click(WebDriver.FindElement(toClick)).Perform();
                        if (!await WaitToPresent(toWait, TimeOut))
                        {
                            MyLogger.Error($"#{m_ID} - Click failed for waiting! mode:{mode} By: {toClick}");
                            return false;
                        }
                        MyLogger.Error($"#{m_ID} - Click success! mode:{mode} By: {toClick}");
                        return true;
                    }
                }
                catch
                {

                }
            }
            MyLogger.Error($"#{m_ID} - Click failed for waiting! mode:{mode} By: {toClick}");
            return false;
        }
    }
}
