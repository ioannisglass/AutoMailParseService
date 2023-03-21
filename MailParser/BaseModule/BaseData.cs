using DbHelper;
using MailParser;
using Logger;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseModule
{
    class BaseData
    {

        static public object lock_proxy_info = new object();
        static public List<ProxyInfo> proxy_info_list = new List<ProxyInfo>();

        static public void save_proxy_info(string file_name)
        {
            lock (lock_proxy_info)
            {
                try
                {
                    File.WriteAllText(file_name, JsonConvert.SerializeObject(proxy_info_list));
                }
                catch (Exception exception)
                {
                    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                }
            }
        }
        static public void load_proxy_info(string file_name)
        {
            lock (lock_proxy_info)
            {
                try
                {
                    proxy_info_list = JsonConvert.DeserializeObject<List<ProxyInfo>>(File.ReadAllText(file_name));

                    foreach (ProxyInfo proxy in proxy_info_list)
                    {
                        Program.g_db.add_proxy_info(proxy);
                    }
                }
                catch (Exception exception)
                {
                    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                }
            }
        }
        static public void load_proxy_info_from_db()
        {
            lock (lock_proxy_info)
            {
                try
                {
                    proxy_info_list.Clear();

                    DataTable dt = Program.g_db.get_proxy_dt();
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow row = dt.Rows[i];

                            ProxyInfo proxy = new ProxyInfo();
                            proxy.host = row["server_url"].ToString();
                            proxy.port = int.Parse(row["server_port"].ToString());
                            proxy.username = row["user_name"].ToString();
                            proxy.password = row["password"].ToString();
                            proxy.type = int.Parse(row["server_type"].ToString());

                            proxy_info_list.Add(proxy);

                            MyLogger.Info($"----- load proxy info from db");
                            MyLogger.Info($"      proxy server_url     : {proxy.host}");
                            MyLogger.Info($"      proxy server_port    : {proxy.port}");
                            MyLogger.Info($"      proxy user_name      : {proxy.username}");
                            MyLogger.Info($"      proxy password       : {proxy.password}");
                            MyLogger.Info($"      proxy server_type    : {proxy.type}");
                        }
                    }
                }
                catch (Exception exception)
                {
                    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                }
            }
        }
    }
}
