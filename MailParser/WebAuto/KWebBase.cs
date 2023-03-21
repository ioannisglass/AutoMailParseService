using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebHelper;

namespace WebAuto
{
    public class KWebBase : IWebHelper
    {
        protected List<string> m_lstrLinks;

        public int timeout;

        public int m_process_result;

        public KWebBase()
        {
            timeout = Program.g_setting.delay_time * 1000;
            Program.g_must_end = false;

            m_process_result = ConstEnv.LOGIN_STARTED;

            MyLogger.Info($"X = #{m_location.X}, Y = #{m_location.Y}");

            m_incognito = false;
            m_dis_js = false;
            m_dis_webrtc = false;
        }

        public async Task<bool> open_browser()
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

        public async Task<bool> Quit_browser()
        {
            try
            {
                /*if (ConstEnv.OS_TYPE == ConstEnv.OS_WINDOWS)
                {
                    if (!await Quit())
                    {
                        MyLogger.Error("Quit Browser failed on Windows.");
                        return false;
                    }
                    MyLogger.Error("Quit Browser success on Windows.");
                }
                else if (ConstEnv.OS_TYPE == ConstEnv.OS_UNIX)
                {
                    //if (!await Quit_undelete_data())
                    if (!await Quit())
                        {
                        MyLogger.Error("Quit Browser failed on Ubuntu");
                        return false;
                    }
                    MyLogger.Error("Quit Browser success on Ubuntu.");
                }
                else
                {
                    MyLogger.Error("Unknown OS type.");
                    return false;
                }
                MyLogger.Info($"Browser quited successfully.");
                return true;*/

                return await Quit_In_Exception();
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
                return false;
            }
        }

        #region virtual functions
        public virtual async Task<int> scrap_link(ZScrapParam param)
        {
            return ConstEnv.SCRAP_FAILED;
        }

        protected virtual bool check_valid_link(string link)
        {
            return false;
        }

        #endregion

    }
    public class KScrapException : Exception
    {
        public KScrapException(string message) : base(message)
        {
        }
    }
}