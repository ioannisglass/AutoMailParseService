using BaseModule;
using MailParser;
using Logger;
using MailHelper;
using MySql.Data.MySqlClient;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using UserHelper;

namespace DbHelper
{
    public class DbMgr
    {
        protected readonly string _database;
        protected readonly string _server;
        protected readonly int _port;
        protected readonly string _uid;
        protected readonly string _password;
        protected readonly bool _ssh;
        protected readonly string _ssh_server;
        protected readonly int _ssh_port;
        protected readonly string _ssh_uid;
        protected readonly string _ssh_password;
        protected readonly string _ssh_keyfile;

        protected ForwardedPortLocal portFwld = null;
        protected SshClient _ssh_client = null;


        public DbMgr(string database, string server, int port, string uid, string password, bool ssh = false, string ssh_server = "", int ssh_port = 22, string ssh_uid = "", string ssh_password = "", string ssh_keyfile = "")
        {
            _database = database;
            _server = server;
            _port = port;
            _uid = uid;
            _password = password;
            _ssh = ssh;
            _ssh_server = ssh_server;
            _ssh_port = ssh_port;
            _ssh_uid = ssh_uid;
            _ssh_password = ssh_password;
            _ssh_keyfile = ssh_keyfile;
        }
        protected bool connect_ssh()
        {
            if (!_ssh)
                return true;

            List<AuthenticationMethod> methods = new List<AuthenticationMethod>();

            if (_ssh_password != "")
                methods.Add(new PasswordAuthenticationMethod(_ssh_uid, _ssh_password));

            if (_ssh_keyfile != "")
            {
                var keyFile = new PrivateKeyFile(_ssh_keyfile);
                var keyFiles = new[] { keyFile };
                methods.Add(new PrivateKeyAuthenticationMethod(_ssh_uid, keyFiles));
            }

            ConnectionInfo connectionInfo = new ConnectionInfo(_ssh_server, _ssh_port, _ssh_uid, methods.ToArray());
            connectionInfo.Timeout = TimeSpan.FromSeconds(1000);

            /*
            * It works fine for SSH user/password.
            * 
                    PasswordConnectionInfo connectionInfo = new PasswordConnectionInfo(_ssh_server, _ssh_uid, _ssh_password);
                    connectionInfo.Timeout = TimeSpan.FromSeconds(5);
            */
            _ssh_client = new SshClient(connectionInfo);
            _ssh_client.Connect();
            if (!_ssh_client.IsConnected)
                throw new Exception("SSH connection is inactive");
            //portFwld = new ForwardedPortLocal("127.0.0.1"/*your computer ip*/, _server /*server ip*/, 3306 /*server mysql port*/);
            portFwld = new ForwardedPortLocal(IPAddress.Loopback.ToString(), "localhost", 3306);
            _ssh_client.AddForwardedPort(portFwld);
            portFwld.Start();
            if (!portFwld.IsStarted)
                return false;

            return true;
        }
        public bool connectable()
        {
            bool ret = false;
            using (MySqlConnection connection = connect())
            {
                try
                {
                    connection.Open();
                    ret = (connection.State == ConnectionState.Open);
                }
                catch (Exception exception)
                {
                    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                    return false;
                }
            }
            return ret;
        }
        public MySqlConnection connect()
        {
            if (_ssh)
            {
                if (portFwld == null || !portFwld.IsStarted)
                    connect_ssh();
            }

            string connection_string;
            if (!_ssh)
                connection_string = String.Format("server={0};port={1};database={2};uid={3};password={4}", _server, _port, _database, _uid, _password);
            else
                connection_string = String.Format("server={0};database={1};uid={2};password={3};port={4}", portFwld.BoundHost, _database, _uid, _password, portFwld.BoundPort);
            return new MySqlConnection(connection_string);
        }
        public void execute_sql(string sql, Dictionary<string, object> cmd_params)
        {
            //MyLogger.Info($"... DB execute_sql : {sql}");
            using (MySqlConnection connection = connect())
            {
                connection.Open();
                using (MySqlTransaction tr = connection.BeginTransaction())
                {
                    try
                    {
                        using (MySqlCommand cmd = new MySqlCommand())
                        {
                            cmd.Connection = connection;
                            cmd.CommandText = sql;
                            cmd.CommandTimeout = 600;
                            if (cmd_params != null)
                            {
                                foreach (KeyValuePair<string, object> p in cmd_params)
                                    cmd.Parameters.AddWithValue(p.Key, p.Value);
                            }
                            cmd.ExecuteNonQuery();

                            tr.Commit();
                        }
                    }
                    catch (MySqlException exception)
                    {
                        MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name})\n    Message : {exception.Message}\n    Number : {exception.Number}\n    {exception.StackTrace}");

                        try
                        {
                            tr.Rollback();
                        }
                        catch (MySqlException exception1)
                        {
                            MyLogger.Error($"Rollback Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name})\n    Message : {exception1.Message}\n    Number : {exception1.Number}\n    {exception1.StackTrace}");
                        }

                        throw exception;
                    }
                }
            }
        }
        public int insert_sql(string sql, Dictionary<string, object> cmd_params)
        {
            int last_inserted_id = -1;

            //MyLogger.Info($"... DB insert_sql : {sql}");
            using (MySqlConnection connection = connect())
            {
                connection.Open();
                using (MySqlTransaction tr = connection.BeginTransaction())
                {
                    try
                    {
                        using (MySqlCommand cmd = new MySqlCommand())
                        {
                            cmd.Connection = connection;
                            cmd.CommandText = sql;
                            cmd.CommandTimeout = 600;
                            if (cmd_params != null)
                            {
                                foreach (KeyValuePair<string, object> p in cmd_params)
                                    cmd.Parameters.AddWithValue(p.Key, p.Value);
                            }
                            cmd.ExecuteNonQuery();

                            last_inserted_id = (int)cmd.LastInsertedId;

                            tr.Commit();
                        }
                    }
                    catch (MySqlException exception)
                    {
                        MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name})\n    Message : {exception.Message}\n    Number : {exception.Number}\n    {exception.StackTrace}");

                        try
                        {
                            tr.Rollback();
                        }
                        catch (MySqlException exception1)
                        {
                            MyLogger.Error($"Rollback Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name})\n    Message : {exception1.Message}\n    Number : {exception1.Number}\n    {exception1.StackTrace}");
                        }
                        last_inserted_id = -1;

                        throw exception;
                    }
                }
            }
            return last_inserted_id;
        }
        public DataTable select(string sql, Dictionary<string, object> cmd_params)
        {
            //MyLogger.Info($"... DB SELECT : {sql}");
            var dt = new DataTable();
            using (MySqlConnection connection = connect())
            {
                using (var da = new MySqlDataAdapter(sql, connection))
                {
                    da.SelectCommand.CommandTimeout = 600;
                    if (cmd_params != null)
                    {
                        foreach (KeyValuePair<string, object> p in cmd_params)
                            da.SelectCommand.Parameters.AddWithValue(p.Key, p.Value);
                    }
                    try
                    {
                        connection.Open(); // not necessarily needed in this case because DataAdapter.Fill does it otherwise 
                        da.Fill(dt);
                    }
                    catch (MySqlException exception)
                    {
                        MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name})\n    Message : {exception.Message}\n    Number : {exception.Number}\n    {exception.StackTrace}");

                        throw exception;
                    }
                }
            }
            return dt;
        }
        public void BulkInsertCSV(string table_name, string csv_file_path, string line_terminater = "\n")
        {
            using (MySqlConnection connection = connect())
            {
                connection.Open();
                var msbl = new MySqlBulkLoader(connection);
                msbl.TableName = table_name;
                msbl.FileName = csv_file_path;
                msbl.FieldTerminator = ",";
                msbl.FieldQuotationCharacter = '"';
                msbl.LineTerminator = line_terminater;
                msbl.Load();
            }
        }
        public void add_user_info(UserInfo user)
        {
            string sql = $"SELECT id FROM {ConstEnv.DB_MAIL_ACCOUNT_TABLE_NAME} WHERE mail = @mail ;";
            DataTable dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail"] = user.mail_address
                    });
            if (dt.Rows != null && dt.Rows.Count > 0)
            {
                int id = int.Parse(dt.Rows[0]["id"].ToString());
                sql = $"UPDATE {ConstEnv.DB_MAIL_ACCOUNT_TABLE_NAME} SET mail = @mail, password = @password, ";
                sql += $"mail_server_url = @mail_server_url, mail_server_port = @mail_server_port, mail_server_type = @mail_server_type, ";
                sql += $"mail_server_ssl = @mail_server_ssl, giftspread_user = @giftspread_user, giftspread_password = @giftspread_password, work_mode = @work_mode WHERE id = @id ;";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail"] = user.mail_address,
                        ["@password"] = user.password,
                        ["@mail_server_url"] = user.mail_server,
                        ["@mail_server_port"] = user.mail_server_port,
                        ["@mail_server_type"] = user.server_type,
                        ["@mail_server_ssl"] = user.ssl ? 1 : 0,
                        ["@giftspread_user"] = user.giftspread_user,
                        ["@giftspread_password"] = user.giftspread_password,
                        ["@work_mode"] = user.report_mode,
                        ["@id"] = id,
                    });

                user.id = id;

                MyLogger.Info($"----- load user info from json (update)");
                MyLogger.Info($"      user id               : {user.id}");
                MyLogger.Info($"      user mail             : {user.mail_address}");
                MyLogger.Info($"      user password         : {user.password}");
                MyLogger.Info($"      user mail_server      : {user.mail_server}");
                MyLogger.Info($"      user mail_server_port : {user.mail_server_port}");
                MyLogger.Info($"      user mail_server_type : {user.server_type}");
                MyLogger.Info($"      user mail_server_ssl  : {user.ssl}");
                MyLogger.Info($"      giftspread_user       : {user.giftspread_user}");
                MyLogger.Info($"      giftspread_password   : {user.giftspread_password}");
                MyLogger.Info($"      work mode             : {user.report_mode}");
            }
            else
            {
                sql = $"INSERT INTO {ConstEnv.DB_MAIL_ACCOUNT_TABLE_NAME} (mail, password, mail_server_url, mail_server_port, mail_server_type, mail_server_ssl, giftspread_user, giftspread_password, work_mode) VALUES ";
                sql += $"( @mail, @password, @mail_server_url, @mail_server_port, @mail_server_type, @mail_server_ssl, @giftspread_user, @giftspread_password, @work_mode ) ;";
                int id = insert_sql(sql,
                            new Dictionary<string, object>()
                            {
                                ["@mail"] = user.mail_address,
                                ["@password"] = user.password,
                                ["@mail_server_url"] = user.mail_server,
                                ["@mail_server_port"] = user.mail_server_port,
                                ["@mail_server_type"] = user.server_type,
                                ["@mail_server_ssl"] = user.ssl ? 1 : 0,
                                ["@giftspread_user"] = user.giftspread_user,
                                ["@giftspread_password"] = user.giftspread_password,
                                ["@work_mode"] = user.report_mode
                            });

                user.id = id;

                MyLogger.Info($"----- load user info from json (insert)");
                MyLogger.Info($"      user id               : {user.id}");
                MyLogger.Info($"      user mail             : {user.mail_address}");
                MyLogger.Info($"      user password         : {user.password}");
                MyLogger.Info($"      user mail_server      : {user.mail_server}");
                MyLogger.Info($"      user mail_server_port : {user.mail_server_port}");
                MyLogger.Info($"      user mail_server_type : {user.server_type}");
                MyLogger.Info($"      user mail_server_ssl  : {user.ssl}");
                MyLogger.Info($"      giftspread_user       : {user.giftspread_user}");
                MyLogger.Info($"      giftspread_password   : {user.giftspread_password}");
                MyLogger.Info($"      work mode             : {user.report_mode}");
            }
        }
        public void add_proxy_info(ProxyInfo proxy)
        {
            string sql = $"SELECT id FROM {ConstEnv.DB_PROXY_TABLE_NAME} WHERE server_url = @server_url ;";
            DataTable dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@server_url"] = proxy.host
                    });
            if (dt.Rows != null && dt.Rows.Count > 0)
            {
                int id = int.Parse(dt.Rows[0]["id"].ToString());
                sql = $"UPDATE {ConstEnv.DB_PROXY_TABLE_NAME} SET server_url = @server_url, server_port = @server_port, ";
                sql += $"user_name = @user_name, password = @password, server_type = @server_type ";
                sql += $"WHERE id = @id ;";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@server_url"] = proxy.host,
                        ["@server_port"] = proxy.port,
                        ["@user_name"] = proxy.username,
                        ["@password"] = proxy.password,
                        ["@server_type"] = proxy.type,
                        ["@id"] = id
                    });

                MyLogger.Info($"----- load proxy info from json (update)");
                MyLogger.Info($"      proxy server_url     : {proxy.host}");
                MyLogger.Info($"      proxy server_port    : {proxy.port}");
                MyLogger.Info($"      proxy user_name      : {proxy.username}");
                MyLogger.Info($"      proxy password       : {proxy.password}");
                MyLogger.Info($"      proxy server_type    : {proxy.type}");
            }
            else
            {
                sql = $"INSERT INTO {ConstEnv.DB_PROXY_TABLE_NAME} (server_url, server_port, user_name, password, server_type) VALUES ";
                sql += $"( @server_url, @server_port, @user_name, @password, @server_type ) ;";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@server_url"] = proxy.host,
                        ["@server_port"] = proxy.port,
                        ["@user_name"] = proxy.username,
                        ["@password"] = proxy.password,
                        ["@server_type"] = proxy.type
                    });

                MyLogger.Info($"----- load proxy info from json (insert)");
                MyLogger.Info($"      proxy server_url     : {proxy.host}");
                MyLogger.Info($"      proxy server_port    : {proxy.port}");
                MyLogger.Info($"      proxy user_name      : {proxy.username}");
                MyLogger.Info($"      proxy password       : {proxy.password}");
                MyLogger.Info($"      proxy server_type    : {proxy.type}");
            }
        }
        public DataTable get_proxy_dt()
        {
            string sql = $"SELECT * FROM {ConstEnv.DB_PROXY_TABLE_NAME}";
            DataTable dt = select(sql, null);
            return dt;
        }
        public List<UserInfo> get_user_list()
        {
            List<UserInfo> user_list = new List<UserInfo>();
            try
            {
                string query = $"SELECT * FROM {ConstEnv.DB_MAIL_ACCOUNT_TABLE_NAME} ;";
                DataTable dt = select(query, null);
                if (dt.Rows == null)
                    return user_list;

                foreach (DataRow row in dt.Rows)
                {
                    UserInfo user = new UserInfo();

                    user.id = int.Parse(row["id"].ToString());
                    user.mail_address = row["mail"].ToString();
                    user.password = row["password"].ToString();
                    user.mail_server = row["mail_server_url"].ToString();
                    user.mail_server_port = int.Parse(row["mail_server_port"].ToString());
                    user.server_type = int.Parse(row["mail_server_type"].ToString());
                    user.ssl = (int.Parse(row["mail_server_ssl"].ToString()) == 1);
                    user.giftspread_user = row["giftspread_user"].ToString();
                    user.giftspread_password = row["giftspread_password"].ToString();
                    user.report_mode = int.Parse(row["work_mode"].ToString());

                    user_list.Add(user);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return user_list;
        }
        public DataTable get_last_checked_mail_info(int account_id)
        {
            try
            {
                string query = $"SELECT * FROM {ConstEnv.DB_MAIL_LAST_WORK_TABLE_NAME} WHERE account_id = @account_id ;";
                DataTable dt = select(query,
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id
                    });
                return dt;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return null;
            }
        }
        public void update_last_checked_mail_info(int account_id, string mail_folder, uint validity, uint uid, DateTime time)
        {
            try
            {
                DataTable dt = select($"SELECT id FROM {ConstEnv.DB_MAIL_LAST_WORK_TABLE_NAME} WHERE account_id = @account_id AND mail_folder = @mail_folder ;",
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id,
                        ["@mail_folder"] = mail_folder
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                {
                    string cmd_text = $"INSERT INTO {ConstEnv.DB_MAIL_LAST_WORK_TABLE_NAME} (account_id, mail_folder, validity, unique_id, time) VALUES ";
                    cmd_text += $"( @account_id, @mail_folder, @validity, @unique_id, @time) ;";
                    execute_sql(cmd_text,
                        new Dictionary<string, object>()
                        {
                            ["@account_id"] = account_id,
                            ["@mail_folder"] = mail_folder,
                            ["@validity"] = validity,
                            ["@unique_id"] = uid,
                            ["@time"] = time.ToString("yyyy-MM-dd HH:mm:ss")
                        });
                }
                else
                {
                    int id = int.Parse(dt.Rows[0]["id"].ToString());

                    string cmd_text = $"UPDATE {ConstEnv.DB_MAIL_LAST_WORK_TABLE_NAME} SET validity = @validity, ";
                    cmd_text += $"unique_id = @unique_id, time = @time WHERE id = @id ;";
                    execute_sql(cmd_text,
                       new Dictionary<string, object>()
                       {
                           ["@id"] = id,
                           ["@validity"] = validity,
                           ["@unique_id"] = uid,
                           ["@time"] = time.ToString("yyyy-MM-dd HH:mm:ss")
                       });
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public int get_account_id_from_mail_id(int id)
        {
            try
            {
                DataTable dt = select($"SELECT account_id FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE id = @id ;",
                    new Dictionary<string, object>()
                    {
                        ["@id"] = id
                    });
                if (dt.Rows == null || dt.Rows.Count != 1)
                    return -1;

                int account_id = int.Parse(dt.Rows[0]["account_id"].ToString());
                return account_id;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return -1;
            }
        }
        public bool is_already_fetched_mail(int account_id, string mail_folder, uint validity, uint uid)
        {
            try
            {
                string cmd_text = $"SELECT * FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE account_id = @account_id AND ";
                cmd_text += $"mail_folder = @mail_folder AND validity = @validity AND unique_id = @unique_id ;";
                DataTable dt = select(cmd_text,
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id,
                        ["@mail_folder"] = mail_folder,
                        ["@validity"] = validity,
                        ["@unique_id"] = uid
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                {
                    return false;
                }
                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return false;
            }
        }
        public bool is_already_fetched_mail_by_hash(int account_id, string hash)
        {
            try
            {
                string cmd_text = $"SELECT * FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE account_id = @account_id AND ";
                cmd_text += $"hash_code = @hash_code ;";
                DataTable dt = select(cmd_text,
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id,
                        ["@hash_code"] = hash
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                {
                    return false;
                }
                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return false;
            }
        }
        public int get_fetched_mail_from_db(int account_id, string mail_folder, uint validity, uint uid)
        {
            try
            {
                string cmd_text = $"SELECT id FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE account_id = @account_id AND mail_folder = @mail_folder AND validity = @validity AND unique_id = @unique_id ;";
                DataTable dt = select(cmd_text,
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id,
                        ["@mail_folder"] = mail_folder,
                        ["@validity"] = validity,
                        ["@unique_id"] = uid
                    });
                if (dt.Rows == null || dt.Rows.Count != 1)
                    return -1;
                int id = int.Parse(dt.Rows[0]["id"].ToString());
                return id;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return -1;
        }
        public int add_fetched_mail_to_db(int account_id, string mail_folder, uint validity, uint uid, string hash, string subject, string sender, DateTime time, string store_path, string mail_type)
        {
            int new_id = -1;
            try
            {
                string cmd_text = $"INSERT INTO {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} (account_id, mail_folder, validity, unique_id, hash_code, local_download_path, checked, subject, sender, time, mail_type) VALUES ";
                cmd_text += $"( @account_id, @mail_folder, @validity, @unique_id, @hash_code, @local_download_path, @checked, @subject, @sender, @time, @mail_type) ;";
                new_id = insert_sql(cmd_text,
                            new Dictionary<string, object>()
                            {
                                ["@account_id"] = account_id,
                                ["@mail_folder"] = mail_folder,
                                ["@validity"] = validity,
                                ["@unique_id"] = uid,
                                ["@hash_code"] = hash,
                                ["@local_download_path"] = store_path,
                                ["@checked"] = ConstEnv.MAIL_UNCHECKED,
                                ["@subject"] = subject,
                                ["@sender"] = sender,
                                ["@time"] = time.ToString("yyyy-MM-dd HH:mm:ss"),
                                ["@mail_type"] = mail_type
                            });
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                new_id = -1;
            }
            return new_id;
        }

        public DataTable get_unchecked_fetched_mail_ids()
        {
            DataTable dt = null;

            try
            {
                string sql;
                sql = $"SELECT id from {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE checked = @checked AND ";

                for (int i = 0; i < Program.g_user.user_info_list.Count; i++)
                    sql += $" account_id = @account_id{i+ 1}";
                sql += " ORDER BY time ;";

                Dictionary<string, object> select_param = new Dictionary<string, object>();
                for (int i = 0; i < Program.g_user.user_info_list.Count; i++)
                    select_param.Add($"@account_id{i + 1}", Program.g_user.user_info_list[i].id);

                select_param.Add("@checked", ConstEnv.MAIL_UNCHECKED);

                dt = select(sql, select_param);
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;

                return dt;
            }
            catch(Exception ex)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {ex.Message}");
            }
            return dt;
        }
        public string get_mail_folder_path(int id)
        {
            DataTable dt = null;
            string path = "";

            try
            {
                dt = select($"SELECT local_download_path FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE id = @id ;",
                    new Dictionary<string, object>()
                    {
                        ["@id"] = id
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return path;

                path = dt.Rows[0]["local_download_path"].ToString();

                path = path.Replace('/', Path.DirectorySeparatorChar);
                path = path.Replace('\\', Path.DirectorySeparatorChar);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return path;
        }
        public void set_mail_checked_flag(int id, int flag)
        {
            // Set checked flag.

            try
            {
                string sql = $"UPDATE {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} SET checked = @checked WHERE id = @id ;";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@checked"] = flag,
                        ["@id"] = id
                    });
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void add_skipped_mail_to_db(int account_id, uint uid, string mail_folder, string subject, string sender, DateTime sent_time)
        {
            try
            {
                string cmd_text = $"INSERT INTO {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} (account_id, uid, mail_folder, subject, sender, time) VALUES ";
                cmd_text += $"( @account_id, @uid, @mail_folder, @subject, @sender, @time )";
                execute_sql(cmd_text,
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id,
                        ["@uid"] = uid,
                        ["@mail_folder"] = mail_folder,
                        ["@subject"] = subject,
                        ["@sender"] = sender,
                        ["@time"] = sent_time
                    });
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public DataTable get_order_cancelled_mails_from_skipped_table()
        {
            DataTable dt = null;

            try
            {
                dt = select($"SELECT * FROM {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} WHERE subject LIKE '%Cancel%' ORDER BY sender, subject ;", null);
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return dt;
        }
        public DataTable get_order_cancelled_mails_from_skipped_table_by_account(int account_id)
        {
            DataTable dt = null;

            try
            {
                dt = select($"SELECT * FROM {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} WHERE subject LIKE '%Cancel%' AND account_id = @account_id ORDER BY sender, subject ;",
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return dt;
        }
        public DataTable get_mail_num_by_mail_type(int account_id)
        {
            DataTable dt = null;

            try
            {
                dt = select($"SELECT mail_type, COUNT(id) as num FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE account_id = @account_id GROUP BY mail_type ORDER BY mail_type ;",
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return dt;
        }
        public DataTable get_mails_by_check_state(int account_id, int check_state)
        {
            DataTable dt = null;

            try
            {
                dt = select($"SELECT id, mail_type FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE account_id = @account_id AND checked = @checked ORDER BY mail_type ;",
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id,
                        ["@checked"] = check_state
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return dt;
        }
        public int get_skipped_mail_num(int account_id)
        {
            DataTable dt = null;

            try
            {
                dt = select($"SELECT COUNT(id) as num FROM {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} WHERE account_id = @account_id ;",
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return 0;

                DataRow row = dt.Rows[0];
                int num = int.Parse(row["num"].ToString());
                return num;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return 0;
        }
        public DataTable get_skipped_mails(int account_id)
        {
            DataTable dt = null;

            try
            {
                //dt = select($"SELECT * FROM {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} WHERE account_id = @account_id AND subject LIKE '%cancel%' ORDER BY mail_folder ;",
                dt = select($"SELECT * FROM {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} WHERE account_id = @account_id ORDER BY mail_folder ;",
                    new Dictionary<string, object>()
                    {
                        ["@account_id"] = account_id
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return dt;
        }
        public void delete_skipped_mails(int id)
        {
            try
            {
                string sql = $"DELETE FROM {ConstEnv.DB_SKIPPED_MAIL_TABLE_NAME} WHERE id = @id ;";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = id
                    });
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public DataTable get_mail_localPaths_by_mail_type(string mail_type)
        {
            DataTable dt = null;

            try
            {
                dt = select($"SELECT id, local_download_path FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE mail_type = @mail_type ;",
                    new Dictionary<string, object>()
                    {
                        ["@mail_type"] = mail_type
                    });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return null;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                dt = null;
            }
            return dt;
        }
        public void clear_gift_card_zen_tables()
        {
            try
            {
                string sql = $"DELETE FROM {ConstEnv.DB_GIFTCARDZEN_TABLE_NAME} ;";
                execute_sql(sql, null);

                sql = $"DELETE FROM {ConstEnv.DB_GIFTCARDZEN_DETAILS_TABLE_NAME} ;";
                execute_sql(sql, null);

                MyLogger.Info("Clear giftcard zen tables.");
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void fill_gift_card_zen_tables(List<GiftCardZenSales> card_list)
        {
            foreach (GiftCardZenSales card in card_list)
            {
                string sql = $"INSERT INTO {ConstEnv.DB_GIFTCARDZEN_TABLE_NAME} ( order_id, date_of_purchase, total ) VALUES ( @order_id, @date_of_purchase, @total ) ;";
                int new_id = insert_sql(sql,
                                new Dictionary<string, object>()
                                {
                                    ["@order_id"] = card.order_ID,
                                    ["@date_of_purchase"] = card.date_of_purchase,
                                    ["@total"] = card.total
                                });
                if (new_id == -1)
                {
                    continue;
                }

                MyLogger.Info("-------------------------------------------------------------------------------");
                MyLogger.Info("Add giftcard zen Card.");
                MyLogger.Info("-------------------------------------------------------------------------------");
                MyLogger.Info($".....order_id         = {card.order_ID}");
                MyLogger.Info($".....date_of_purchase = {card.date_of_purchase}");
                MyLogger.Info($".....total            = {card.total}");

                foreach (GiftCardZenDigiCard detail in card.lst_digi_cards)
                {
                    sql = $"INSERT INTO {ConstEnv.DB_GIFTCARDZEN_DETAILS_TABLE_NAME} ( parent_id, retailer, value, cost, discount, card_number, pin, status ) VALUES ( @parent_id, @retailer, @value, @cost, @discount, @card_number, @pin, @status ) ;";
                    execute_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@parent_id"] = new_id,
                            ["@retailer"] = detail.retailer,
                            ["@value"] = detail.value,
                            ["@cost"] = detail.cost,
                            ["@discount"] = detail.discount,
                            ["@card_number"] = detail.card_number,
                            ["@pin"] = detail.pin,
                            ["@status"] = detail.status,
                        });

                    MyLogger.Info($"..........parent_id            = {new_id}");
                    MyLogger.Info($"..........retailer             = {detail.retailer}");
                    MyLogger.Info($"..........value                = {detail.value}");
                    MyLogger.Info($"..........cost                 = {detail.cost}");
                    MyLogger.Info($"..........discount             = {detail.discount}");
                    MyLogger.Info($"..........card_number          = {detail.card_number}");
                    MyLogger.Info($"..........pin                  = {detail.pin}");
                    MyLogger.Info($"..........status               = {detail.status}");
                }
            }
        }
        public void insert_order_status_to_db(KReportBase report)
        {
            string[] order_ids = report.m_order_id.Split(',');
            foreach (string order_id in order_ids)
            {
                if (order_id.Trim() == "")
                    continue;

                string sql = $"INSERT INTO {ConstEnv.DB_ORDER_STATUS_TABLE_NAME} ( order_id, mail_type, status, report_id, mail_senttime ) ";
                sql += $" VALUES ( @order_id, @mail_type, @status, @report_id, @mail_senttime ) ;";
                insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@order_id"] = order_id,
                            ["@mail_type"] = report.m_mail_type.ToString(),
                            ["@status"] = report.m_order_status,
                            ["@report_id"] = report.m_report_id,
                            ["@mail_senttime"] = report.m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")
                        });
            }
        }
        public void insert_order_status_to_db(int account_id, int report_id, string order_id, string mail_type, string order_status, DateTime mail_sent_date)
        {
            string[] order_ids = order_id.Split(',');
            foreach (string ord in order_ids)
            {
                if (ord.Trim() == "")
                    continue;

                string sql = $"INSERT INTO {ConstEnv.DB_ORDER_STATUS_TABLE_NAME} ( account_id, order_id, mail_type, status, report_id, mail_senttime ) ";
                sql += $" VALUES ( @account_id, @order_id, @mail_type, @status, @report_id, @mail_senttime ) ;";
                insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@account_id"] = account_id,
                            ["@order_id"] = ord,
                            ["@mail_type"] = mail_type,
                            ["@status"] = order_status,
                            ["@report_id"] = report_id,
                            ["@mail_senttime"] = mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss")
                        });
            }
        }
        public int insert_report_to_main_table_db(int mail_id, KReportBase report)
        {
            int report_id = -1;

            string sql = "";
            sql += $"INSERT INTO {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} ( mail_type, mail_id, mail_senttime, order_id, retailer, receiver, status, total, tax ) ";
            sql += $" VALUES ( @mail_type, @mail_id, @mail_senttime, @order_id, @retailer, @receiver, @status, @total, @tax ) ;";
            report_id = insert_sql(sql,
                            new Dictionary<string, object>()
                            {
                                ["@mail_type"] = report.m_mail_type.ToString(),
                                ["@mail_id"] = mail_id,
                                ["@mail_senttime"] = report.m_mail_sent_date.ToString("yyyy-MM-dd HH:mm:ss"),
                                ["@order_id"] = report.m_order_id,
                                ["@retailer"] = report.m_retailer,
                                ["@receiver"] = report.m_receiver,
                                ["@status"] = report.m_order_status,
                                ["@total"] = report.m_total,
                                ["@tax"] = report.m_tax
                            });
            return report_id;
        }
        public void insert_cr_report_to_db(int report_id, KReportCR report)
        {
            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_CR_TABLE_NAME} WHERE id = @id ;";
            DataTable dt = select(sql,
                                new Dictionary<string, object>()
                                {
                                    ["@id"] = report_id
                                });
            if (dt != null && dt.Rows.Count == 1)
            {
                delete_report_cr_by_report_id(report_id);
                delete_reportdata_details_by_report_id(report_id);
                delete_reportdata_cashback_by_report_id(report_id);
            }

            sql = $"INSERT INTO {ConstEnv.DB_REPORT_CR_TABLE_NAME} ( id, purchase_date, discount, payment_type, payment_id ) ";
            sql += $" VALUES ( @id, @purchase_date, @discount, @payment_type, @payment_id ) ;";
            insert_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id,
                        ["@purchase_date"] = report.m_purchase_date.ToString("yyyy-MM-dd HH:mm:ss"),
                        ["@discount"] = report.m_discount,
                        ["@payment_type"] = report.m_cr_payment_type,
                        ["@payment_id"] = report.m_cr_payment_id
                    });

            insert_giftcard_details_info_to_db(report_id, report.m_giftcard_details);
            insert_giftcard_details_info_to_db(report_id, report.m_giftcard_details_v1);
            insert_giftcard_details_info_to_db(report_id, report.m_giftcard_details_v2);

            insert_instant_cashback_to_db(report_id, report.m_instant_cashback);
        }
        public void insert_op_report_to_db(int report_id, KReportOP report)
        {
            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_OP_TABLE_NAME} WHERE id = @id ;";
            DataTable dt = select(sql,
                                new Dictionary<string, object>()
                                {
                                    ["@id"] = report_id
                                });
            if (dt != null && dt.Rows.Count == 1)
            {
                delete_report_op_by_report_id(report_id);
                delete_reportdata_product_by_report_id(report_id);
                delete_reportdata_payment_by_report_id(report_id);
            }

            sql = $"INSERT INTO {ConstEnv.DB_REPORT_OP_TABLE_NAME} ( id, purchase_date, ship_address, ship_address_state ) ";
            sql += $" VALUES ( @id, @purchase_date, @ship_address, @ship_address_state ) ;";
            insert_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id,
                        ["@ship_address"] = report.m_op_ship_address,
                        ["@ship_address_state"] = report.m_op_ship_address_state,
                        ["@purchase_date"] = report.m_op_purchase_date.ToString("yyyy-MM-dd HH:mm:ss"),
                    });

            insert_products_to_db(report_id, report.m_product_items);
            insert_payment_card_info_to_db(report_id, report.m_payment_card_list);
        }
        public void insert_sc_report_to_db(int report_id, KReportSC report)
        {
            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_SC_TABLE_NAME} WHERE id = @id ;";
            DataTable dt = select(sql,
                                new Dictionary<string, object>()
                                {
                                    ["@id"] = report_id
                                });
            if (dt != null && dt.Rows.Count == 1)
            {
                delete_report_sc_by_report_id(report_id);
                delete_reportdata_product_by_report_id(report_id);
            }

            sql = $"INSERT INTO {ConstEnv.DB_REPORT_SC_TABLE_NAME} ( id, ship_date, post_type, tracking, expected_delivery_date ) ";
            sql += $" VALUES ( @id, @ship_date, @post_type, @tracking, @expected_delivery_date ) ;";
            insert_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id,
                        ["@ship_date"] = report.m_sc_ship_date.ToString("yyyy-MM-dd HH:mm:ss"),
                        ["@post_type"] = report.m_sc_post_type,
                        ["@tracking"] = report.m_sc_tracking,
                        ["@expected_delivery_date"] = report.m_sc_expected_deliver_date.ToString("yyyy-MM-dd HH:mm:ss")
                    });

            insert_products_to_db(report_id, report.m_product_items);
        }
        public void insert_cc_report_to_db(int report_id, KReportCC report)
        {
            delete_reportdata_product_by_report_id(report_id);
            delete_reportdata_payment_by_report_id(report_id);

            insert_products_to_db(report_id, report.m_product_items);
            insert_payment_card_info_to_db(report_id, report.m_payment_card_list);
        }
        public void insert_payment_card_info_to_db(int report_id, List<ZPaymentCard> pi)
        {
            foreach (ZPaymentCard p in pi)
            {
                string sql = $"INSERT INTO {ConstEnv.DB_REPORTDATA_PI_TABLE_NAME} ( report_id, payment_type, last_4_digit, price ) ";
                sql += $" VALUES ( @report_id, @payment_type, @last_4_digit, @price ) ;";
                insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@report_id"] = report_id,
                            ["@payment_type"] = p.payment_type,
                            ["@last_4_digit"] = p.last_4_digit,
                            ["@price"] = p.price
                        });
            }
        }
        public void insert_giftcard_details_info_to_db<T>(int report_id, List<T> details)
        {
            foreach (T t in details)
            {
                string retailer = "";
                float value = 0;
                float cost = 0;
                string gift_card = "";
                string pin = "";
                int data_type = ConstEnv.CARD_DETAILS_ALL;

                if (t.GetType() == typeof(ZGiftCardDetails))
                {
                    ZGiftCardDetails d = t as ZGiftCardDetails;
                    retailer = d.m_retailer;
                    value = d.m_value;
                    cost = d.m_cost;
                    gift_card = d.m_gift_card;
                    pin = d.m_pin;
                    data_type = ConstEnv.CARD_DETAILS_ALL;
                }
                else if (t.GetType() == typeof(ZGiftCardDetails_V1))
                {
                    ZGiftCardDetails_V1 d = t as ZGiftCardDetails_V1;
                    retailer = d.m_retailer;
                    value = d.m_value;
                    cost = d.m_cost;
                    data_type = ConstEnv.CARD_DETAILS_V1;
                }
                else if (t.GetType() == typeof(ZGiftCardDetails_V2))
                {
                    ZGiftCardDetails_V2 d = t as ZGiftCardDetails_V2;
                    gift_card = d.m_gift_card;
                    pin = d.m_pin;
                    data_type = ConstEnv.CARD_DETAILS_V2;
                }
                else
                {
                    throw new Exception($"Invalid type. {details.GetType().ToString()}");
                }

                string sql = $"INSERT INTO {ConstEnv.DB_REPORTDATA_DETAILS_TABLE_NAME} ( report_id, data_type, retailer, value, cost, gift_card, pin ) ";
                sql += $" VALUES ( @report_id, @data_type, @retailer, @value, @cost, @gift_card, @pin ) ;";
                insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@report_id"] = report_id,
                            ["@data_type"] = data_type,
                            ["@retailer"] = retailer,
                            ["@value"] = value,
                            ["@cost"] = cost,
                            ["@gift_card"] = gift_card,
                            ["@pin"] = pin
                        });
            }
        }
        public void insert_instant_cashback_to_db(int report_id, List<float> instant_cashback)
        {
            foreach (float f in instant_cashback)
            {
                string sql = $"INSERT INTO {ConstEnv.DB_REPORTDATA_CASHBACK_TABLE_NAME} ( report_id, instant_cashback ) ";
                sql += $" VALUES ( @report_id, @instant_cashback ) ;";
                insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@report_id"] = report_id,
                            ["@instant_cashback"] = f
                        });
            }
        }
        public void insert_products_to_db(int report_id, List<ZProduct> products)
        {
            foreach (ZProduct p in products)
            {
                string sql = $"INSERT INTO {ConstEnv.DB_REPORTDATA_PRODUCT_TABLE_NAME} ( report_id, title, sku, qty, price, status ) ";
                sql += $" VALUES ( @report_id, @title, @sku, @qty, @price, @status ) ;";
                insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@report_id"] = report_id,
                            ["@title"] = p.title,
                            ["@sku"] = p.sku,
                            ["@qty"] = p.qty,
                            ["@price"] = p.price,
                            ["@status"] = p.status
                        });
            }
        }
        public void update_report_parent_id(int report_id, int parent_id)
        {
            string sql = $"UPDATE {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} SET parent_id = @parent_id WHERE id = @id ";
            execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id,
                        ["@parent_id"] = parent_id
                    });
        }
        public int[] get_child_report_ids(int report_id)
        {
            DataTable dt = null;
            int[] child_ids = null;

            string sql = $"SELECT id FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE parent_id = @parent_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@parent_id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count == 0)
                return child_ids;

            child_ids = new int[dt.Rows.Count];

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                int id = int.Parse(row[0].ToString());
                child_ids[i++] = id;
            }

            return child_ids;
        }
        public bool already_exist_report(string order_id, KReportBase.MailType mail_type)
        {
            DataTable dt = null;

            try
            {
                if (order_id == "")
                    return false;

                string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE order_id = @order_id AND mail_type = @mail_type ;";
                dt = select(sql,
                        new Dictionary<string, object>()
                        {
                            ["@order_id"] = order_id,
                            ["@mail_type"] = mail_type.ToString()
                        });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return false;

                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return false;
        }
        public bool already_exist_report(KReportBase new_report, out bool status_changed)
        {
            DataTable dt = null;

            status_changed = false;

            try
            {
                if (new_report.m_order_id == "")
                    return false;

                string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE order_id = @order_id AND mail_type = @mail_type ;";
                dt = select(sql,
                        new Dictionary<string, object>()
                        {
                            ["@order_id"] = new_report.m_order_id,
                            ["@mail_type"] = new_report.m_mail_type.ToString()
                        });
                if (dt.Rows == null || dt.Rows.Count == 0)
                    return false;

                DataRow row = dt.Rows[0];

                string old_status = row["status"].ToString();
                DateTime old_senttime = DateTime.Parse(row["mail_senttime"].ToString());

                if (old_status != new_report.m_order_status)
                {
                    status_changed = true;
                }

                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return false;
        }
        public KReportBase.MailType get_mail_type_by_report_Id(int report_id)
        {
            KReportBase.MailType mail_type = KReportBase.MailType.Mail_Unknown;
            DataTable dt = null;

            try
            {
                string sql = $"SELECT mail_type FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE id = @id ;";
                dt = select(sql,
                        new Dictionary<string, object>()
                        {
                            ["@id"] = report_id
                        });
                if (dt.Rows == null || dt.Rows.Count != 1)
                    return mail_type;

                string mail_type_str = dt.Rows[0][0].ToString();
                mail_type = (KReportBase.MailType) Enum.Parse(typeof(KReportBase.MailType), mail_type_str, true);

                return mail_type;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return KReportBase.MailType.Mail_Unknown;
            }
        }
        public void get_report_main_info_from_db(KReportBase report, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE id = @id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count != 1)
                throw new Exception($"Card not found : id = {report_id}");

            DataRow row = dt.Rows[0];

            int mail_id = int.Parse(row["mail_id"].ToString());
            if (mail_id != -1)
                report.m_mail_account_id = get_account_id_from_mail_id(mail_id);

            report.set_order_id(row["order_id"].ToString());
            report.set_total(float.Parse(row["total"].ToString()));
            report.m_tax = float.Parse(row["tax"].ToString());
            report.m_mail_sent_date = DateTime.Parse(row["mail_senttime"].ToString());
            report.m_retailer = row["retailer"].ToString();
            report.m_receiver = row["receiver"].ToString();
            report.m_order_status = row["status"].ToString();

            string mail_type_str = row["mail_type"].ToString();
            report.m_mail_type = (KReportBase.MailType)Enum.Parse(typeof(KReportBase.MailType), mail_type_str, true);
        }
        private void get_giftcard_details_from_db(KReportCR report, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORTDATA_DETAILS_TABLE_NAME} WHERE report_id = @report_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@report_id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count == 0)
                return;

            foreach (DataRow row in dt.Rows)
            {
                int data_type = int.Parse(row["data_type"].ToString());
                string retailer = row["retailer"].ToString();
                float value = float.Parse(row["value"].ToString());
                float cost = float.Parse(row["cost"].ToString());
                string gift_card = row["gift_card"].ToString();
                string pin = row["pin"].ToString();

                if (data_type == ConstEnv.CARD_DETAILS_ALL)
                {
                    report.add_giftcard_details(retailer, value, cost, gift_card, pin);
                }
                else if (data_type == ConstEnv.CARD_DETAILS_V1)
                {
                    report.add_giftcard_details_v1(retailer, value, cost);
                }
                else if (data_type == ConstEnv.CARD_DETAILS_V2)
                {
                    report.add_giftcard_details_v2(gift_card, pin);
                }
            }
        }
        public DataTable get_order_reports_from_db()
        {
            DataTable dt = null;

            string sql = "";
            sql += $" SELECT";
            sql += $"    t0.account_id as account_id, t1.id as report_id, t1.order_id as order_id, t1.mail_type as mail_type, t1.status as order_status, t1.mail_senttime as mail_senttime";
            sql += $" FROM";
            sql += $"    {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t0,";
            sql += $"    (SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_type LIKE 'OP_%' OR mail_type LIKE 'SC_%' OR mail_type LIKE 'CC_%') as t1";
            sql += $" WHERE ";
            sql += $"    t0.id = t1.mail_id";
            sql += $" ORDER BY ";
            sql += $"    mail_senttime ";
            sql += $" ;";

            dt = select(sql, null);
            return dt;
        }
        private void get_report_instant_cashback_from_db(KReportCR report, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORTDATA_CASHBACK_TABLE_NAME} WHERE report_id = @report_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@report_id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count == 0)
                return;

            foreach (DataRow row in dt.Rows)
            {
                float f = float.Parse(row["instant_cashback"].ToString());

                if (report.m_instant_cashback == null)
                    report.m_instant_cashback = new List<float>();
                report.m_instant_cashback.Add(f);
            }
        }
        private void get_products_from_db(List<ZProduct> products, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORTDATA_PRODUCT_TABLE_NAME} WHERE report_id = @report_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@report_id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count == 0)
                return;

            foreach (DataRow row in dt.Rows)
            {
                ZProduct item = new ZProduct();

                item.title = row["title"].ToString();
                item.sku = row["sku"].ToString();
                item.qty = int.Parse(row["qty"].ToString());
                item.price = float.Parse(row["price"].ToString());
                item.status = row["status"].ToString();

                products.Add(item);
            }
        }
        private void get_payment_info_from_db(List<ZPaymentCard> payinfos, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORTDATA_PI_TABLE_NAME} WHERE report_id = @report_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@report_id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count == 0)
                return;

            foreach (DataRow row in dt.Rows)
            {
                ZPaymentCard pi = new ZPaymentCard();

                pi.payment_type = row["payment_type"].ToString();
                pi.last_4_digit = row["last_4_digit"].ToString();
                pi.price = float.Parse(row["price"].ToString());

                payinfos.Add(pi);
            }
        }
        public void get_cr_report_from_db(KReportCR report, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_CR_TABLE_NAME} WHERE id = @id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count != 1)
                throw new Exception($"Card not found : id = {report_id}");

            DataRow row = dt.Rows[0];

            report.m_purchase_date = DateTime.Parse(row["purchase_date"].ToString());
            report.m_discount = float.Parse(row["discount"].ToString());
            report.m_cr_payment_type = row["payment_type"].ToString();
            report.m_cr_payment_id = row["payment_id"].ToString();

            get_giftcard_details_from_db(report, report_id);
            get_report_instant_cashback_from_db(report, report_id);
        }
        public void get_cc_report_info_from_db(KReportCC report, int report_id)
        {
            report.m_product_items = new List<ZProduct>();
            get_products_from_db(report.m_product_items, report_id);

            report.m_payment_card_list = new List<ZPaymentCard>();
            get_payment_info_from_db(report.m_payment_card_list, report_id);
        }
        public void get_sc_report_info_from_db(KReportSC report, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_SC_TABLE_NAME} WHERE id = @id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count != 1)
                throw new Exception($"Card not found : id = {report_id}");

            DataRow row = dt.Rows[0];

            report.m_sc_ship_date = DateTime.Parse(row["ship_date"].ToString());
            report.m_sc_post_type = row["post_type"].ToString();
            report.set_tracking(row["tracking"].ToString());
            report.m_sc_expected_deliver_date = DateTime.Parse(row["expected_delivery_date"].ToString());

            report.m_product_items = new List<ZProduct>();
            get_products_from_db(report.m_product_items, report_id);
        }
        public void get_op_report_info_from_db(KReportOP report, int report_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_OP_TABLE_NAME} WHERE id = @id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = report_id
                    });
            if (dt.Rows == null || dt.Rows.Count != 1)
                throw new Exception($"Card not found : id = {report_id}");

            DataRow row = dt.Rows[0];

            report.m_op_purchase_date = DateTime.Parse(row["purchase_date"].ToString());

            report.m_product_items = new List<ZProduct>();
            get_products_from_db(report.m_product_items, report_id);

            report.m_payment_card_list = new List<ZPaymentCard>();
            get_payment_info_from_db(report.m_payment_card_list, report_id);
        }
        public void delete_reports_by_mail_type(string mail_type, int account_id)
        {
            string sql = "";

            sql += $"SELECT ";
            sql += $"    t0.id ";
            sql += $"FROM";
            sql += $"    (SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_type LIKE @mail_type) as t0, ";
            sql += $"    {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t1 ";
            sql += $"WHERE ";
            sql += $"    t0.mail_id = t1.id AND t1.account_id = @account_id ";
            sql += $" ;";

            DataTable dt = select(sql,
                                new Dictionary<string, object>()
                                {
                                    ["@mail_type"] = mail_type,
                                    ["@account_id"] = account_id
                                });
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    int report_id = int.Parse(row["id"].ToString());

                    delete_report_data_by_report_id(report_id);
                }
            }
        }
        public void delete_report_data_by_mail_id(int mail_id)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_id = @mail_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_id"] = mail_id
                    });
            if (dt.Rows != null && dt.Rows.Count > 0)
            {
                List<int> report_ids = new List<int>();
                foreach (DataRow row in dt.Rows)
                {
                    int report_id = int.Parse(row["id"].ToString());
                    report_ids.Add(report_id);
                }

                foreach (int report_id in report_ids)
                    delete_report_data_by_report_id(report_id);
            }

            sql = $"UPDATE {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} SET checked = @checked WHERE id = @id ;";
            execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = mail_id,
                        ["@checked"] = 0
                    });
        }
        public void delete_report_data_by_report_id(int id)
        {
            DataTable dt = null;
            int parent_id = -1;

            string sql = $"SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE id = @id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = id
                    });
            if (dt.Rows != null && dt.Rows.Count == 1)
            {
                parent_id = int.Parse(dt.Rows[0]["parent_id"].ToString());
            }

            sql = $"DELETE FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE id = @id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@id"] = id
                });
            delete_report_cr_by_report_id(id);
            delete_report_op_by_report_id(id);
            delete_report_sc_by_report_id(id);
            delete_reportdata_cashback_by_report_id(id);
            delete_reportdata_details_by_report_id(id);
            delete_reportdata_payment_by_report_id(id);
            delete_reportdata_product_by_report_id(id);

            delete_order_status_by_report_id(id);

            if (parent_id != -1)
            {
                sql = $"UPDATE {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} SET parent_id = '-1' WHERE parent_id = @parent_id ;";
                execute_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@parent_id"] = parent_id
                        });

                delete_report_data_by_report_id(parent_id);
            }
            MyLogger.Info($"Deleted report. id : {id}");
        }
        public void delete_order_status_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_ORDER_STATUS_TABLE_NAME} WHERE report_id = @report_id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@report_id"] = id
                });
            MyLogger.Info($"Deleted order status. report id : {id}");
        }
        public void delete_report_cr_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORT_CR_TABLE_NAME} WHERE id = @id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@id"] = id
                });
        }
        public void delete_report_op_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORT_OP_TABLE_NAME} WHERE id = @id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@id"] = id
                });
        }
        public void delete_report_sc_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORT_SC_TABLE_NAME} WHERE id = @id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@id"] = id
                });
        }
        public void delete_reportdata_cashback_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORTDATA_CASHBACK_TABLE_NAME} WHERE report_id = @report_id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@report_id"] = id
                });
        }
        public void delete_reportdata_details_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORTDATA_DETAILS_TABLE_NAME} WHERE report_id = @report_id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@report_id"] = id
                });
        }
        public void delete_reportdata_payment_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORTDATA_PI_TABLE_NAME} WHERE report_id = @report_id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@report_id"] = id
                });
        }
        public void delete_reportdata_product_by_report_id(int id)
        {
            string sql = $"DELETE FROM {ConstEnv.DB_REPORTDATA_PRODUCT_TABLE_NAME} WHERE report_id = @report_id ;";
            execute_sql(sql,
                new Dictionary<string, object>()
                {
                    ["@report_id"] = id
                });
        }
        public DataTable find_matched_cr1_reports(string order, int mail_account_id)
        {
            DataTable dt = null;

            string sql = $"SELECT t1.* FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} as t1, {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t2 ";
            sql += $" WHERE (t1.mail_type = @mail_type1 OR t1.mail_type = @mail_type2 OR t1.mail_type = @mail_type3 ) AND t1.order_id = @order_id AND t1.parent_id = @parent_id AND t1.mail_id = t2.id AND t2.account_id = @mail_account_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_type1"] = KReportBase.MailType.CR_1_1.ToString(),
                        ["@mail_type2"] = KReportBase.MailType.CR_1_2.ToString(),
                        ["@mail_type3"] = KReportBase.MailType.CR_1_3.ToString(),
                        ["@order_id"] = order,
                        ["@mail_account_id"] = mail_account_id,
                        ["@parent_id"] = -1
                    });
            if (dt == null || dt.Rows == null)
                return null;

            return dt;
        }
        public DataTable find_matched_cr2_reports(string order, int mail_account_id)
        {
            DataTable dt = null;

            string sql = $"SELECT t1.* FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} as t1, {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t2 ";
            sql += $" WHERE (t1.mail_type = @mail_type1 OR t1.mail_type = @mail_type2 ) AND t1.order_id = @order_id AND t1.parent_id = @parent_id AND t1.mail_id = t2.id AND t2.account_id = @mail_account_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_type1"] = KReportBase.MailType.CR_2_1.ToString(),
                        ["@mail_type2"] = KReportBase.MailType.CR_2_2.ToString(),
                        ["@order_id"] = order,
                        ["@mail_account_id"] = mail_account_id,
                        ["@parent_id"] = -1
                    });
            if (dt == null || dt.Rows == null)
                return null;

            return dt;
        }
        public DataTable find_matched_cr3_reports(string order, int mail_account_id)
        {
            DataTable dt = null;

            string sql = $"SELECT t1.* FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} as t1, {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t2 ";
            sql += $" WHERE (t1.mail_type = @mail_type1 OR t1.mail_type = @mail_type2 ) AND t1.order_id = @order_id AND t1.parent_id = @parent_id AND t1.mail_id = t2.id AND t2.account_id = @mail_account_id ;";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_type1"] = KReportBase.MailType.CR_3_1.ToString(),
                        ["@mail_type2"] = KReportBase.MailType.CR_3_2.ToString(),
                        ["@order_id"] = order,
                        ["@mail_account_id"] = mail_account_id,
                        ["@parent_id"] = -1
                    });
            if (dt == null || dt.Rows == null)
                return null;

            return dt;
        }
        public int find_matched_cr4_1_report_by_order(string order, int mail_account_id)
        {
            DataTable dt = null;

            string sql = $"SELECT t1.id FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} as t1, {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t2 ";
            sql += $" WHERE t1.mail_type = @mail_type1 AND t1.order_id = @order_id AND t1.parent_id = @parent_id AND t1.mail_id = t2.id AND t2.account_id = @mail_account_id ;";

            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_type1"] = KReportBase.MailType.CR_4_1.ToString(),
                        ["@order_id"] = order,
                        ["@mail_account_id"] = mail_account_id,
                        ["@parent_id"] = -1
                    });
            if (dt == null || dt.Rows == null || dt.Rows.Count != 1)
                return -1;
            int id = int.Parse(dt.Rows[0][0].ToString());
            return id;
        }
        public int find_matched_cr4_2_report(KReportCR4 report)
        {
            DataTable dt = null;

            string sql = $"SELECT t1.id FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} as t1, {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} as t2, {ConstEnv.DB_REPORT_CR_TABLE_NAME} as t3 ";
            sql += $" WHERE t1.mail_type = @mail_type2 AND t3.payment_type = @payment_type AND t3.payment_id = @payment_id AND t3.id = t1.id AND t1.parent_id = @parent_id AND t1.mail_id = t2.id AND t2.account_id = @mail_account_id ;";

            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_type2"] = KReportBase.MailType.CR_4_2.ToString(),
                        ["@payment_type"] = report.m_cr_payment_type,
                        ["@payment_id"] = report.m_cr_payment_id,
                        ["@mail_account_id"] = report.m_mail_account_id,
                        ["@parent_id"] = -1
                    });
            if (dt == null || dt.Rows == null || dt.Rows.Count != 1)
                return -1;

            int id = int.Parse(dt.Rows[0][0].ToString());
            return id;
        }
        public DataTable collect_sales_tax_order_data(int account_id)
        {
            DataTable dt = null;

            string sql = "";
            sql += $" SELECT t_op.order_id as order_id, t_op.retailer as retailer, t_op.total as total, t_op.tax as tax, t_op.mail_senttime as time, t_op1.ship_address_state as state, t_op.id as op_id, t_sc.id as sc_id ";
            sql += $" FROM ";
            sql += $"    (SELECT * FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} ) as t_fetched, ";
            sql += $"    (SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_type LIKE 'OP_%' AND status = @op_status) as t_op, ";
            sql += $"    (SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_type LIKE 'SC_%' AND status = @sc_status) as t_sc, ";
            sql += $"    (SELECT * FROM {ConstEnv.DB_REPORT_OP_TABLE_NAME} WHERE ship_address_state = @state_1 OR ship_address_state = @state_2 ) as t_op1 ";
            sql += $" WHERE ";
            sql += $"    t_op.order_id = t_sc.order_id AND t_op.order_id != '' AND t_op.tax != '0' AND t_op.id = t_op1.id ";
            sql += $"     AND t_op.mail_id = t_fetched.id AND t_fetched.account_id = @account_id ";
            sql += $" ; ";

            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@op_status"] = ConstEnv.REPORT_ORDER_STATUS_PURCHAESD,
                        ["@sc_status"] = ConstEnv.REPORT_ORDER_STATUS_SHIPPED,
                        ["@state_1"] = "NJ",
                        ["@state_2"] = "NY",
                        ["@account_id"] = account_id
                    });
            if (dt == null || dt.Rows == null)
                return null;

            return dt;
        }
        public DataTable collect_sales_tax_payments(int op_report_id)
        {
            DataTable dt = null;

            string sql = "";
            sql += $" SELECT payment_type, last_4_digit, price ";
            sql += $" FROM ";
            sql += $"   {ConstEnv.DB_REPORTDATA_PI_TABLE_NAME} ";
            sql += $" WHERE ";
            sql += $"    report_id = @report_id ";
            sql += $" ; ";

            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@report_id"] = op_report_id
                    });
            if (dt == null || dt.Rows == null)
                return null;

            return dt;
        }
        public void update_file_hash()
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} ;";
            dt = select(sql, null);
            if (dt == null || dt.Rows == null)
                return;

            int num = 0;
            foreach (DataRow row in dt.Rows)
            {
                int id = int.Parse(row["id"].ToString());

                string eml_file = row["local_download_path"].ToString();
                eml_file = eml_file.Replace('/', Path.DirectorySeparatorChar);
                eml_file = eml_file.Replace('\\', Path.DirectorySeparatorChar);
                eml_file = Path.Combine(eml_file, ConstEnv.LOCAL_MAIL_FILE_NAME);

                string hash = "";
                if (!File.Exists(eml_file))
                    continue;

                hash = XMailHelper.get_file_hash_md5(eml_file);

                sql = $"UPDATE {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} SET hash_code = @hash_code WHERE id = @id ;";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = id,
                        ["@hash_code"] = hash
                    });

                MyLogger.Info($"Update hash {++num} / {dt.Rows.Count}");
            }
        }
        public int[] get_empty_pin_cr2_mail_ids()
        {
            List<int> mail_id_list = new List<int>();

            DataTable dt = null;

            string sql = "";
            sql += $"SELECT t4.id as report_id, t4.mail_id as mail_id ";
            sql += $"FROM ";
            sql += $"   ( ";
            sql += $"   SELECT ";
            sql += $"      t0.id as id ";
            sql += $"   FROM ";
            sql += $"      (SELECT * FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_type = 'CR_2_2') as t0, ";
            sql += $"      {ConstEnv.DB_REPORTDATA_DETAILS_TABLE_NAME} as t1 ";
            sql += $"   WHERE ";
            sql += $"      t0.id = t1.report_id AND t1.gift_card <> '' AND t1.pin = '' ";
            sql += $"   GROUP BY t0.id ";
            sql += $"   ) as t3, ";
            sql += $"   report_main as t4 ";
            sql += $"WHERE";
            sql += $"   t3.id = t4.id ";
            sql += $";   ";

            dt = select(sql, null);
            if (dt == null || dt.Rows == null)
                return mail_id_list.ToArray();

            foreach (DataRow row in dt.Rows)
            {
                int mail_id = int.Parse(row["mail_id"].ToString());
                mail_id_list.Add(mail_id);
            }

            return mail_id_list.ToArray();
        }
        public int[] get_parsing_failed_cr2_mail_ids()
        {
            List<int> mail_id_list = new List<int>();

            DataTable dt = null;

            string sql = $"SELECT id FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE checked = '2' AND mail_type = 'CR_2_2' ;";

            dt = select(sql, null);
            if (dt == null || dt.Rows == null)
                return mail_id_list.ToArray();

            foreach (DataRow row in dt.Rows)
            {
                int mail_id = int.Parse(row["id"].ToString());
                mail_id_list.Add(mail_id);
            }

            return mail_id_list.ToArray();
        }
        public string[] get_duplicated_mail_hash_code()
        {
            List<string> hash_list = new List<string>();

            DataTable dt = null;

            string sql = $"SELECT hash_code, COUNT(*) as c FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME}  GROUP BY hash_code HAVING c > '1' ; ";
            dt = select(sql, null);
            if (dt == null || dt.Rows == null)
                return hash_list.ToArray();

            foreach (DataRow row in dt.Rows)
            {
                string hash_code = row["hash_code"].ToString();
                hash_list.Add(hash_code);
            }

            return hash_list.ToArray();
        }
        public List<int> get_report_id_by_mail_id(int mail_id)
        {
            List<int> report_id_list = new List<int>();

            DataTable dt = null;

            string sql = $"SELECT id FROM {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} WHERE mail_id = @mail_id ;";

            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@mail_id"] = mail_id
                    });
            if (dt == null || dt.Rows == null)
                return report_id_list;

            foreach (DataRow row in dt.Rows)
            {
                int id = int.Parse(row["id"].ToString());
                report_id_list.Add(mail_id);
            }

            return report_id_list;
        }
        public void delete_duplicated_feteched_mail(string hash_code)
        {
            DataTable dt = null;

            string sql = $"SELECT * FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE hash_code = @hash_code ; ";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@hash_code"] = hash_code
                    });
            if (dt == null || dt.Rows == null || dt.Rows.Count <= 1)
                return;

            List<int> mail_id_list = new List<int>();
            List<string> eml_path_list = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                int id = int.Parse(row["id"].ToString());
                mail_id_list.Add(id);

                string local_eml_path = row["local_download_path"].ToString();
                eml_path_list.Add(local_eml_path);

                MyLogger.Info($"Duplicated mail id : {id}, hash = {hash_code}");
            }

            sql = "";
            sql += $"SELECT t1.id as report_id, t1.mail_id as mail_id, t1.order_id  as order_id ";
            sql += $"FROM ";
            sql += $"    (SELECT * FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE hash_code = @hash_code) as t0, ";
            sql += $"    {ConstEnv.DB_REPORT_MAIN_TABLE_NAME} as t1 ";
            sql += $"WHERE ";
            sql += $"    t1.mail_id = t0.id ";
            sql += $" ; ";
            dt = select(sql,
                    new Dictionary<string, object>()
                    {
                        ["@hash_code"] = hash_code
                    });
            if (dt == null || dt.Rows == null || dt.Rows.Count == 0)
                return;

            int mail_id_to_remain = int.Parse(dt.Rows[0]["mail_id"].ToString());
            int report_id_to_remain = int.Parse(dt.Rows[0]["report_id"].ToString());
            MyLogger.Info($"mail_id_to_remain : {mail_id_to_remain}, report_id_to_remain : {report_id_to_remain}");

            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                int report_id = int.Parse(row["report_id"].ToString());
                int mail_id = int.Parse(row["mail_id"].ToString());

                delete_report_data_by_report_id(report_id);
            }

            for (int i = 0; i < mail_id_list.Count; i++)
            {
                int id = mail_id_list[i];
                string local_eml_path = eml_path_list[i];

                if (id == mail_id_to_remain)
                    continue;

                sql = $"DELETE FROM {ConstEnv.DB_FETCH_MAIL_TABLE_NAME} WHERE id = @id ; ";
                execute_sql(sql,
                    new Dictionary<string, object>()
                    {
                        ["@id"] = id
                    });

                if (Directory.Exists(local_eml_path))
                    Directory.Delete(local_eml_path, true);

                MyLogger.Info($"Deleted duplicated mail. id : {id}, hash : {hash_code}");
            }
        }
        public void create_old_crm_table(DataTable dt)
        {
            string sql = "";

            sql += $" CREATE TABLE old_crm (";

            foreach (DataColumn column in dt.Columns)
            {
                sql += $" {column.ColumnName} varchar(255),";
            }
            sql = sql.Substring(0, sql.Length - 1);
            sql += " ) ;";
            execute_sql(sql, null);
        }
        public int insert_statement_files_to_db(string file_name, int parse_state)
        {
            string sql = "";

            sql += $"INSERT INTO {ConstEnv.DB_STATEMENT_FILES_TABLE_NAME} (pdf_file, checked) VALUES ";
            sql += $"( @pdf_file, @checked ) ;";
            int id = insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@pdf_file"] = file_name,
                            ["@checked"] = parse_state
                        });
            return id;
        }
        public int get_statment_file_parsing_state(string file_name)
        {
            DataTable dt;
            string sql = $"SELECT checked FROM {ConstEnv.DB_STATEMENT_FILES_TABLE_NAME} WHERE pdf_file = @pdf_file ;";
            dt = select(sql,
                        new Dictionary<string, object>()
                        {
                            ["@pdf_file"] = file_name
                        });
            if (dt == null || dt.Rows.Count == 0)
                return -1;

            int state = int.Parse(dt.Rows[0]["checked"].ToString());
            return state;
        }
        public int insert_bank_transaction(int pdf_file_id, string bank_name, string account, string trn_type, DateTime date, string description, float amount, string number)
        {
            string sql = "";

            sql += $"INSERT INTO {ConstEnv.DB_BANK_TRANSACTIONS_TABLE_NAME} (pdf_file_id, name, account, trn_type, date, description, amount, number) VALUES ";
            sql += $"( @pdf_file_id, @name, @account, @trn_type, @date, @description, @amount, @number ) ;";
            int id = insert_sql(sql,
                        new Dictionary<string, object>()
                        {
                            ["@pdf_file_id"] = pdf_file_id,
                            ["@name"] = bank_name,
                            ["@account"] = account,
                            ["@trn_type"] = trn_type,
                            ["@date"] = date.ToString("yyyy-MM-dd HH:mm:ss"),
                            ["@description"] = description,
                            ["@amount"] = amount,
                            ["@number"] = number
                        });
            return id;
        }
    }
}
