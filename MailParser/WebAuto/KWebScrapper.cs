using MailParser;
using WebAuto;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Data;

namespace WebAuto
{
    public class KWebScrapper
    {
        private List<KWebBase> m_web_handlers = new List<KWebBase>();
        public KWebScrapper()
        {
            m_web_handlers.Add(new KWebCardpool());
            m_web_handlers.Add(new KWebGCSpread());
            m_web_handlers.Add(new KWebRMN());
            m_web_handlers.Add(new KWebGetegiftcard());
            m_web_handlers.Add(new KWebHomedepot());
            m_web_handlers.Add(new KWebValueSears());
            m_web_handlers.Add(new KWebAmzGC());
        }
        static public async void scrap(KReportBase report)
        {
            KWebScrapper scrapper = new KWebScrapper();
            await scrapper.scrap_async(report);
        }
        public async Task scrap_async(KReportBase report)
        {
            int k;

            if (report.m_scrap_params == null)
                return;

            foreach (ZScrapParam param in report.m_scrap_params)
            {
                int ret = ConstEnv.SCRAP_FAILED;
                for (k = 0; k < m_web_handlers.Count; k++)
                {
                    ret = await m_web_handlers[k].scrap_link(param);
                    if (ret != ConstEnv.SCRAP_UNSUPPORTED)
                        break;
                }

                if (k == m_web_handlers.Count)
                {
                    MyLogger.Info($"It's not unsupported web link. link = {param.link}, vendor = {param.report.m_mail_type}");
                    param.status = ConstEnv.SCRAP_UNSUPPORTED;
                }
                else
                {
                    if (ret != ConstEnv.SCRAP_SUCCESS)
                    {
                        MyLogger.Info($"Scrap NOT Success : status = {ret}, link = {param.link}, vendor = {param.report.m_mail_type}");
                    }
                    else
                    {
                        MyLogger.Info($"Scrap Success : link = {param.link}, vendor = {param.report.m_mail_type}");
                    }
                }
                param.status = ret;
            }
        }
    }
}
