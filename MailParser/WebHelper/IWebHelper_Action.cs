using GiftCard;
using Logger;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace WebHelper
{
    partial class IWebHelper
    {
        public async Task<bool> refresh()
        {
            WebDriver.Navigate().Refresh();
            return true;
        }
        public void DeleteCurrentChromeData()
        {
            try
            {
                Directory.Delete(m_chr_user_data_dir, true);
                return;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"#{m_ID} - Deleting chrome data dir failed. {ex.Message}");
            }
        }

        public void DeleteCurrentChromeData_except_cookie()
        {
            try
            {
                string[] subfolders = Directory.GetDirectories(m_chr_user_data_dir);

                foreach (string subfolder in subfolders)
                    if (subfolder != m_chr_user_data_dir + "\\Default")
                        Directory.Delete(subfolder, true);

                string[] fileEntries = Directory.GetFiles(m_chr_user_data_dir);
                foreach (string fileName in fileEntries)
                    File.Delete(fileName);

                subfolders = Directory.GetDirectories(m_chr_user_data_dir + "\\Default");
                foreach (string subfolder in subfolders)
                    Directory.Delete(subfolder, true);

                return;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"#{m_ID} - Deleting chrome data dir failed. {ex.Message}");
            }
        }
        public async Task<bool> Quit()
        {
            try
            {
                foreach (var hnd in WebDriver.WindowHandles)
                {
                    WebDriver.SwitchTo().Window(hnd);
                    WebDriver.Close();
                }
                WebDriver.Quit();
                WebDriver.Dispose();
                DeleteCurrentChromeData();
                //DeleteCurrentChromeData_except_cookie();
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched in Quit - {ex.Message}");
                return false;
            }
        }

        public async Task<bool> Quit_undelete_data()
        {
            try
            {
                foreach (var hnd in WebDriver.WindowHandles)
                {
                    WebDriver.SwitchTo().Window(hnd);
                    WebDriver.Close();
                }
                WebDriver.Quit();
                WebDriver.Dispose();
                return true;
            }

            catch (Exception ex)
            {
                //MainApp.log_error($"Thread #{m_ID} - Error catched in Quit_undelete_data() - {ex.Message}");
                return false;
            }
        }
        public void ClearChromeData()
        {
            try
            {
                string path = "ChromeData";
                Directory.Delete(path, true);
                Directory.CreateDirectory(path);
            }
            catch
            {
            }
        }
        public static void KillAllChromeDriverProcess()
        {
            MyLogger.Error("Killing all chrome drivers");
            int id = Process.GetCurrentProcess().Id;
            var list1 = new List<Process>((IEnumerable<Process>)Process.GetProcessesByName("chromedriver"));
            var list2 = new List<Process>((IEnumerable<Process>)Process.GetProcessesByName("chrome"));
            foreach (Process proc in list1)
            {
                if (proc.GetParentID() == id)
                {
                    foreach (Process proc2 in list2)
                    {
                        if (proc2.GetParentID() == proc.Id)
                        {
                            new Thread((ThreadStart)(() =>
                            {
                                try
                                {
                                    proc2.Kill();
                                }
                                catch
                                {
                                }
                            })).Start();
                        }
                    }
                    new Thread((ThreadStart)(() =>
                    {
                        try
                        {
                            proc.Kill();
                        }
                        catch
                        {
                        }
                    })).Start();
                }
            }
        }

        public static void ForceStopChrDriver()
        {
            KillAllChromeDriverProcess();
            Directory.Delete("ChromeData", true);
        }

        public async Task<bool> Quit_In_Exception()
        {
            try
            {
                foreach (var hnd in WebDriver.WindowHandles)
                {
                    WebDriver.SwitchTo().Window(hnd);
                    WebDriver.Close();
                }
                WebDriver.Quit();
                WebDriver.Dispose();
                DeleteCurrentChromeData();
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.Error($"Error catched in Quit - {ex.Message}");
                ForceStopChrDriver();
                return false;
            }
        }
    }
}
