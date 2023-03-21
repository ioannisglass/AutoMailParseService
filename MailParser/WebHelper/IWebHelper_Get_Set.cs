using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebHelper
{
    partial class IWebHelper
    {
        public async Task<string> get_value(string xpath, string err = "", string field = "innerText")
        {
            try
            {
                var elem = WebDriver.FindElementByXPath(xpath);
                if (elem == null)
                    return err;
                return elem.GetAttribute(field);
            }
            catch (Exception )
            {
                return err;
            }
        }
        public string get_elem_value(IWebElement _elem, string xpath, string err = "", string field = "innerText")
        {
            try
            {
                var elem = _elem.FindElement(By.XPath(xpath));
                if (elem == null)
                    return err;
                return elem.GetAttribute(field);
            }
            catch (Exception ex)
            {
                return err;
            }
        }
        public async Task<string> set_value(string xpath, string val, string field = "value")
        {
            Object node = null;
            string script = "(function()" +
                                "{" +
                                    "node = document.evaluate(\"" + xpath + "\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;" +
                                    "if (node==null) return '" + m_err_str + "';" +
                                    "node." + field + "=\"" + val + "\";" +
                                    "return 'ok';" +
                            "})()";
            node = m_js.ExecuteScript(script);
            if (node != null)
                return node.ToString();
            return m_err_str;
        }

        public async Task<string> set_attribute(string xpath, string val, string field = "innerText")
        {
            Object node = null;
            string script = "(function()" +
                                "{" +
                                    "node = document.evaluate(\"" + xpath + "\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;" +
                                    "if (node==null) return '" + m_err_str + "';" +
                                    "node.setAttribute('" + field + "',\"" + val + "\");" +
                                    "return 'ok';" +
                            "})()";
            node = m_js.ExecuteScript(script);
            if (node != null)
                return node.ToString();
            return m_err_str;
        }
    }
}
