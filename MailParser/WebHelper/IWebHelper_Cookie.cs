using OpenQA.Selenium;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using GiftCard;
using Cookie = System.Net.Cookie;
using Logger;

namespace WebHelper
{
    partial class IWebHelper
    {
        public async Task<bool> remove_all_cookies()
        {
            try
            {
                WebDriver.Manage().Cookies.DeleteAllCookies();
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Removing history failed. {ex.Message}");
            }
            return false;
        }
        public CookieContainer ConvertSeleniumCookieToCookieContainer(ICookieJar seleniumCookie)
        {
            CookieContainer cookieContainer = new CookieContainer();
            using (IEnumerator<OpenQA.Selenium.Cookie> enumerator = seleniumCookie.AllCookies.GetEnumerator())
            {
                while (((IEnumerator)enumerator).MoveNext())
                {
                    OpenQA.Selenium.Cookie current = enumerator.Current;
                    cookieContainer.Add(new Cookie(current.Name, current.Value, current.Path, current.Domain));
                }
            }
            return cookieContainer;
        }
        public List<string> Convert_SeleniumCookie_To_StringList(ICookieJar seleniumCookie)
        {
            List<string> lstCookieString;
            lstCookieString = new List<string>();
            using (IEnumerator<OpenQA.Selenium.Cookie> enumerator = seleniumCookie.AllCookies.GetEnumerator())
            {
                while (((IEnumerator)enumerator).MoveNext())
                {
                    OpenQA.Selenium.Cookie current = enumerator.Current;
                    lstCookieString.Add(current.ToString());
                }
            }
            return lstCookieString;
        }

        public List<string> Convert_CookieContainer_To_StringList(CookieContainer cookieContainer, Uri uri)
        {
            List<string> lstCookieString;
            lstCookieString = new List<string>();
            try
            {
                foreach (Cookie cook in cookieContainer.GetCookies(uri))
                {
                    lstCookieString.Add(cook.ToString());
                }
                return lstCookieString;
            }
            catch (Exception e)
            {
                MyLogger.Error($"Error catched in cookieContainer2String[] : {e.Message}");
            }
            return null;
        }
        public CookieContainer Convert_ListString_To_CookieContainer(List<string> strCookies)
        {
            CookieContainer cookiecontainer = new CookieContainer();
            try
            {
                foreach (string cookie in strCookies)
                    cookiecontainer.SetCookies(new Uri("http://.com"), cookie);
                return cookiecontainer;
            }
            catch(Exception ex)
            {
                MyLogger.Error($"Error catched in Convert_ListString_To_CookieContainer[] : {ex.Message}");
            }
            return null;
        }
    }
}
