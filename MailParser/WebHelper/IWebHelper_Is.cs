using GiftCard;
using Logger;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebHelper
{
    partial class IWebHelper
    {
        public bool IsElementPresent(By by)
        {
            try
            {
                WebDriver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException ex)
            {
                MyLogger.Error($"Error catched in IsElementPresent - {by.ToString()} - {ex.Message}");
                return false;
            }
        }
        public async Task<bool> IsElementVisible(By by, int timeout = 0)
        {
            try
            {
                Stopwatch wt = new Stopwatch();
                wt.Start();
                do
                {
                    if (IsElementVisible(WebDriver.FindElement(by)))
                        return true;
                    await TaskDelay(100);
                } while (wt.ElapsedMilliseconds < timeout);
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool IsElementVisible(IWebElement element)
        {
            return element.Displayed && element.Enabled;
        }
    }
}
