using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using GiftCard;
using Logger;

namespace WebHelper
{
    partial class IWebHelper
    {
        public async Task<bool> WaitUrlChange(string url, int timeout = 7000)
        {
            try
            {
                Stopwatch wt = new Stopwatch();
                wt.Start();
                while (wt.ElapsedMilliseconds < timeout)
                {
                    if (WebDriver.Url != url)
                        return true;
                    await TaskDelay(100);
                }
            }
            catch (Exception ex)
            {
                MyLogger.Error($"#{m_ID} - Failed to wait for url change. Exception:{ex.Message}");
            }
            return false;
        }
        public async Task<bool> WaitUrlSame(string url, int timeout = 7000)
        {
            try
            {
                Stopwatch wt = new Stopwatch();
                wt.Start();
                while (wt.ElapsedMilliseconds < timeout)
                {
                    if (WebDriver.Url == url)
                        return true;
                    await TaskDelay(100);
                }
            }
            catch (Exception ex)
            {
                MyLogger.Error($"#{m_ID} - Failed to wait for url same. Exception:{ex.Message}");
            }
            return false;
        }
        public async Task<bool> WaitToVisible(string classname, int TimeOut = 1000)
        {
            return await WaitToVisible(By.ClassName(classname), TimeOut);
        }

        public async Task<bool> WaitToVisibleByPath(string xpath, int TimeOut = 1000)
        {
            return await WaitToVisible(By.XPath(xpath), TimeOut);
        }
        public async Task<bool> WaitToVisible(By by, int TimeOut = 1000)
        {
            Stopwatch wt = new Stopwatch();
            wt.Start();
            while (wt.ElapsedMilliseconds < TimeOut)
            {
                if (await IsElementVisible(by))
                    return true;
                Thread.Sleep(100);
            }
            return false;
        }

        public async Task<bool> WaitToUnvisable(string classname, int TimeOut = 1000)
        {
            return await WaitToUnvisable(By.ClassName(classname), TimeOut);
        }
        public async Task<bool> WaitToUnvisable(By by, int TimeOut = 1000)
        {
            Stopwatch wt = new Stopwatch();
            wt.Start();
            while (wt.ElapsedMilliseconds < TimeOut)
            {
                try
                {
                    if (!await IsElementVisible(by))
                        return true;
                }
                catch
                {
                    return false;
                }
                await TaskDelay(100);
            }
            return false;
        }

        public async Task<bool> WaitToPresent(string classname, int TimeOut = 2000)
        {
            return await WaitToPresent(By.ClassName(classname), TimeOut);
        }
        public async Task<bool> WaitToPresentByPath(string xpath, int TimeOut = 2000)
        {
            return await WaitToPresent(By.XPath(xpath), TimeOut);
        }

        public async Task<bool> WaitToPresent(By by, int TimeOut = 5000)
        {
            Stopwatch wt = new Stopwatch();
            wt.Start();
            do
            {
                if (IsElementPresent(by))
                    return true;
                await TaskDelay(1000);
            }
            while (wt.ElapsedMilliseconds < TimeOut);
            return false;
        }
        public async Task<By> WaitToPresent(List<By> by, int TimeOut = 1000)
        {
            Stopwatch wt = new Stopwatch();
            wt.Start();
            while (wt.ElapsedMilliseconds < TimeOut)
            {
                using (List<By>.Enumerator enumerator = by.GetEnumerator())
                {
                    while (enumerator.MoveNext())
                    {
                        By current = enumerator.Current;
                        if (IsElementPresent(current))
                            return current;
                    }
                }
                await TaskDelay(100);
            }
            return null;
        }

        public async Task<bool> Wait_for_Nextstep_By_Path(string xpath, int timeout = 30000)
        {
            return await Wait_for_Nextstep(By.XPath(xpath), timeout);
        }
        public async Task<bool> Wait_for_Nextstep_By_ClsName(string classname, int timeout = 30000)
        {
            return await Wait_for_Nextstep(By.ClassName(classname), timeout);
        }
        public async Task<bool> Wait_for_Nextstep(By by, int timeout = 30000)
        {
            try
            {
                if (!await WaitToPresent(by, timeout) || !await WaitToVisible(by, timeout))
                {
                    MyLogger.Error($"#{m_ID} - Next step not shown");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.Info(string.Format("Error wait for next step: {0}", (object)ex.Message));
                return false;
            }
        }
        public async Task<bool> Navigate_Find_by_XPath(string url, string xpath, int timeout = 30000)
        {
            return await Navigate_Find(url, By.XPath(xpath), timeout);
        }
        public async Task<bool> Navigate_Find(string url, By by, int timeout = 30000)
        {
            try
            {
                if (!await Navigate(url) || !await Wait_for_Nextstep(by, timeout))
                    return false;
                return true;
            }
            catch(Exception /*ex*/)
            {
                //MyLogger.Error($"{m_param.user} - Error catched in Navigate_Find : {ex.Message}");
                return true;
            }
        }

        public async Task RandomWait()
        {
            int delay = (new Random()).Next(100, 1000);
            //await TaskDelay(delay);
            System.Threading.Thread.Sleep(delay);
        }

        public async Task TaskDelay(int delay)
        {
            //await TaskDelay(delay);
            System.Threading.Thread.Sleep(delay);
        }
    }
}
