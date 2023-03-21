using OpenQA.Selenium;
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
        public async Task<bool> TryEnterText_by_xpath(string xpath, string textToEnter, string atributeToEdit = "value", int TimeOut = 10000, bool manualyEnter = false)
        {
            return await TryEnterText(By.XPath(xpath), textToEnter, atributeToEdit, TimeOut, manualyEnter);
        }
        public async Task<bool> TryEnterText(string classname, string textToEnter, string atributeToEdit = "value", int TimeOut = 10000, bool manualyEnter = false)
        {
            return await TryEnterText(By.ClassName(classname), textToEnter, atributeToEdit, TimeOut, manualyEnter);
        }

        public async Task<bool> TryEnterText(By by, string textToEnter, string atributeToEdit = "value", int TimeOut = 10000, bool manualyEnter = false)
        {
            Stopwatch wt = new Stopwatch();
            wt.Start();
            while (wt.ElapsedMilliseconds < TimeOut)
            {
                try
                {
                    if (IsElementPresent(by) && await IsElementVisible(by))
                    {
                        WebDriver.FindElement(by).SendKeys((string)Keys.Control + "a");
                        if (manualyEnter)
                            WebDriver.FindElement(by).SendKeys(textToEnter);
                        else
                            WebDriver.ExecuteScript($"arguments[0].value = '{textToEnter}';", ((RemoteWebDriver)WebDriver).FindElement(by));

                        for (int index = 0; index < 11; ++index)
                        {
                            if ((string)WebDriver.ExecuteScript("return arguments[0].value;", WebDriver.FindElement(by)) == textToEnter)
                            {
                                return true;
                            }
                            await TaskDelay(100);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MyLogger.Error($"#{m_ID} - Failed to enter text. Exception:{ex.Message}");
                    return false;
                }
                await TaskDelay(100);
            }
            return false;
        }
        public async Task<bool> TryEnterText_All(By by, string textToEnter, string atributeToEdit = "value", int TimeOut = 10000)
        {
            bool flag = await TryEnterText(by, textToEnter, atributeToEdit = "value", TimeOut, false);
            if (flag)
                return true;
            else
                flag = await TryEnterText(by, textToEnter, atributeToEdit = "value", TimeOut, true);

            if (flag)
                return true;

            Actions action = new Actions(WebDriver);
            action.MoveToElement(WebDriver.FindElement(by)).Perform();

            flag = await TryEnterText(by, textToEnter, atributeToEdit = "value", TimeOut, false);
            if (flag)
                return true;
            else
                flag = await TryEnterText(by, textToEnter, atributeToEdit = "value", TimeOut, true);

            if (flag)
                return true;

            return false;
        }
    }
}
