using BaseModule;
using MailParser;
using Logger;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using UserHelper;
using Utils;

namespace MailHelper
{
    public class XMailChecker
    {
        private UserInfo m_account = new UserInfo();
        private List<KMailBaseParser> m_mail_vendor_handlers = new List<KMailBaseParser>();
        public XMailChecker(UserInfo user)
        {
            m_account = user;

            m_mail_vendor_handlers.Add(new KMailCR1());
            m_mail_vendor_handlers.Add(new KMailCR2());
            m_mail_vendor_handlers.Add(new KMailCR3());
            m_mail_vendor_handlers.Add(new KMailCR4());
            m_mail_vendor_handlers.Add(new KMailCR7());
            m_mail_vendor_handlers.Add(new KMailCR8());
            m_mail_vendor_handlers.Add(new KMailBaseOP());
            m_mail_vendor_handlers.Add(new KMailBaseSC());
            m_mail_vendor_handlers.Add(new KMailBaseCC());
        }
        public ImapClient connect(CancellationTokenSource cts)
        {
            ImapClient client = new ImapClient();

            try
            {
                MyLogger.Info($"connecting .... server = {m_account.mail_server}, port = {m_account.mail_server_port}, login = {m_account.mail_address}, password = {m_account.password}");

                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(m_account.mail_server, m_account.mail_server_port, SecureSocketOptions.Auto, cts.Token);
                MyLogger.Info("server connected");

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate(m_account.mail_address, m_account.password, cts.Token);
                MyLogger.Info("login succeed");

                /*
                 * https://github.com/jstedfast/MailKit/issues/828
                 * using (var client = new ImapClient ()) {
                        var credentials = new NetworkCredential ("username", "password");
                        client.ProxyClient = new Socks5Client ("host", 88, credentials);
                        client.Connect ("imap.example.com", 995, SecureSocketOptions.SslOnConnect);
                        client.Authenticate ("username", "password");
                    }
                 *
                 *
                 */
            }
            catch (OperationCanceledException exception)
            {
                MyLogger.Info($"({System.Reflection.MethodBase.GetCurrentMethod().Name}): {nameof(OperationCanceledException)} thrown with message: {exception.Message}");
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            MyLogger.Info($"connected = {client.IsConnected}, logined = {client.IsAuthenticated}, cancelled = {cts.IsCancellationRequested}");
            return client;
        }
        public void disconnect(ImapClient client)
        {
            try
            {
                client.Disconnect(true);
                MyLogger.Info("disconnected");
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void disconnect(ImapClient client, CancellationTokenSource cts)
        {
            try
            {
                client.Disconnect(true, cts.Token);
                MyLogger.Info("disconnected");
            }
            catch (OperationCanceledException exception)
            {
                MyLogger.Info($"({System.Reflection.MethodBase.GetCurrentMethod().Name}): {nameof(OperationCanceledException)} thrown with message: {exception.Message}");
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        private bool get_last_checked_mail_info(DataTable dt, IMailFolder box, string clean_mail_folder, out UniqueId last_uid, out DateTime last_time)
        {
            last_uid = UniqueId.Invalid;
            last_time = DateTime.MinValue;

            if (dt == null || dt.Rows.Count == 0)
                return false;

            foreach (DataRow row in dt.Rows)
            {
                string folder_name = row["mail_folder"].ToString();
                if (clean_mail_folder == folder_name)
                {
                    uint validity = uint.Parse(row["validity"].ToString());
                    if (box.UidValidity == validity)
                    {
                        uint uid = uint.Parse(row["unique_id"].ToString());
                        last_uid = new UniqueId(uid);
                        last_time = DateTime.Parse(row["time"].ToString());
                        MyLogger.Info($"({m_account.mail_address}) : mail folder = {clean_mail_folder} last uid = {last_uid.Id}, last time = {last_time.ToString()}");
                    }
                    else
                    {
                        last_time = DateTime.Parse(row["time"].ToString());
                        MyLogger.Info($"({m_account.mail_address}) : mail folder = {clean_mail_folder} validity is changed. {validity} -> {box.UidValidity}, last time = {last_time.ToString()}");
                    }
                    return true;
                }
            }
            return false;
        }
        public void check_mails(CancellationTokenSource cts)
        {
            ImapClient client = null;
            DataTable last_checked_dt = null;
            int k;

            try
            {
                client = connect(cts);
                if (cts.IsCancellationRequested)
                    return;
                if (!client.IsConnected || !client.IsAuthenticated)
                    return;

                MyLogger.Info($"({m_account.mail_address}) : Start get folders and UniqueIds....");

                last_checked_dt = Program.g_db.get_last_checked_mail_info(m_account.id);

                foreach (var nss in new FolderNamespaceCollection[] { client.PersonalNamespaces, client.SharedNamespaces, client.OtherNamespaces })
                {
                    if (cts.IsCancellationRequested)
                        return;
                    foreach (var ns in nss)
                    {
                        if (cts.IsCancellationRequested)
                            return;

                        var boxes = client.GetFolders(ns);

                        foreach (IMailFolder box in boxes)
                        {
                            if (cts.IsCancellationRequested)
                                return;

                            box.Subscribe();
                            box.Open(FolderAccess.ReadOnly, cts.Token);

                            MyLogger.Info($"({m_account.mail_address}) : start checking mail folder = {box.FullName} contains {box.Count} mails");

                            if (box.Count == 0)
                                continue;
                            if (box.FullName.ToUpper() == "DELETED ITEMS")
                                continue;
                            if (box.FullName.ToUpper() == "SENT ITEMS")
                                continue;

                            uint start;
                            uint last;
                            string clean_mail_folder = Str_Utils.CleanPath(box.FullName);

                            UniqueId last_fetched_uid;
                            DateTime last_fetched_time;

                            if (get_last_checked_mail_info(last_checked_dt, box, clean_mail_folder, out last_fetched_uid, out last_fetched_time))
                            {
                                if (last_fetched_uid != UniqueId.Invalid)
                                {
                                    start = last_fetched_uid.Id + 1;
                                }
                                else
                                {
                                    DateTime start_time = last_fetched_time - new TimeSpan(1, 0, 0, 0);
                                    start_time = new DateTime(start_time.Year, start_time.Month, start_time.Day);
                                    SearchQuery search_query = SearchQuery.SentSince(start_time);
                                    var uids = box.Search(search_query, cts.Token);
                                    if (uids.Count == 0)
                                        continue;
                                    start = uids[0].Id;
                                }
                            }
                            else
                            {
                                var sub_items = box.Fetch(0, 0, MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope, cts.Token);
                                if (sub_items.Count == 0)
                                {
                                    MyLogger.Info($"({m_account.mail_address}) : failed to get the start uid from {box.FullName}");
                                    continue;
                                }
                                start = sub_items[0].UniqueId.Id;
                            }
                            var last_items = box.Fetch(box.Count - 1, box.Count - 1, MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope, cts.Token);
                            if (last_items.Count == 0) // unknown reason.
                            {
                                MyLogger.Info($"({m_account.mail_address}) : failed to get the last uid from {box.FullName} start uid = {start}");
                                continue;
                            }
                            last = last_items[0].UniqueId.Id;
                            MyLogger.Info($"({m_account.mail_address}) : fetch mail folder = {box.FullName} start uid = {start} last uid = {last}");

                            uint rng_start = start;
                            uint rng_last;
                            while (rng_start <= last)
                            {
                                rng_last = Math.Min(rng_start + 999, last);
                                var sub_range = new UniqueIdRange(new UniqueId(rng_start), new UniqueId(rng_last));
                                MyLogger.Info($"({m_account.mail_address}) : Start fetch mail folder = {box.FullName} from {rng_start} to {rng_last}");
                                var sub_items = box.Fetch(sub_range, MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope, cts.Token);
                                MyLogger.Info($"({m_account.mail_address}) : End fetch mail folder = {box.FullName} from {rng_start} to {rng_last}, ret num = {sub_items.Count} mails");

                                foreach (var item in sub_items)
                                {
                                    string subject = XMailHelper.get_subject_from_envelop(item.Envelope);
                                    string sender = XMailHelper.get_sender_from_envelop(item.Envelope);
                                    DateTime sent_time = XMailHelper.get_sentdate_from_envelop(item.Envelope);
                                    int mail_order = 0;

                                    if (sender != m_account.mail_address)
                                    {
                                        for (k = 0; k < m_mail_vendor_handlers.Count; k++)
                                        {
                                            if (m_mail_vendor_handlers[k].check_valid_mail(m_account.report_mode, subject, sender, out mail_order))
                                                break;
                                        }
                                        if (k == m_mail_vendor_handlers.Count)
                                        {
                                            MyLogger.Info($"({m_account.mail_address}) : It's not gift card mail. mail folder = {box.FullName}, mail uid = {item.UniqueId}, subject = {subject}, sender = {sender}");
                                            Program.g_db.add_skipped_mail_to_db(m_account.id, item.UniqueId.Id, clean_mail_folder, subject, sender, sent_time);
                                        }
                                        else
                                        {
                                            MyLogger.Info($"({m_account.mail_address}) : Found gift card mail. mail folder = {box.FullName}, validity = {box.UidValidity}, uid = {item.UniqueId}, subject = {subject}, sender = {sender}");
                                            MyLogger.Info($"({m_account.mail_address}) : mail type = {m_mail_vendor_handlers[k].GetType().ToString()}, order = {mail_order}");

                                            string mail_type = XMailHelper.get_mail_type_string(m_mail_vendor_handlers[k], mail_order);
                                            download_card_mail(box, item.UniqueId, mail_type, cts);
                                        }
                                    }
                                    else
                                    {
                                        MyLogger.Info($"({m_account.mail_address}) : Skip sent mail. mail folder = {box.FullName}, mail uid = {item.UniqueId}, subject = {subject}, sender = {sender}");
                                        Program.g_db.add_skipped_mail_to_db(m_account.id, item.UniqueId.Id, clean_mail_folder, subject, sender, sent_time);
                                    }
                                    Program.g_db.update_last_checked_mail_info(m_account.id, clean_mail_folder, box.UidValidity, item.UniqueId.Id, item.Envelope.Date?.DateTime ?? DateTime.MinValue);
                                }

                                rng_start = rng_last + 1;
                            }
                        }
                    }
                }

                MyLogger.Info($"({m_account.mail_address}) : End get folders");

                disconnect(client, cts);
            }
            catch (OperationCanceledException exception)
            {
                MyLogger.Info($"({System.Reflection.MethodBase.GetCurrentMethod().Name}): {nameof(OperationCanceledException)} thrown with message: {exception.Message}");

                if (client.IsConnected)
                    disconnect(client);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");

                if (client.IsConnected)
                    disconnect(client, cts);
            }
        }
        public void extract_mails_from_skipped_mails(CancellationTokenSource cts)
        {
            ImapClient client = null;
            DataTable skipped_mails_dt = null;
            int k;

            try
            {
                client = connect(cts);
                if (cts.IsCancellationRequested)
                    return;
                if (!client.IsConnected || !client.IsAuthenticated)
                    return;

                 MyLogger.Info($"({m_account.mail_address}) : Start extracting from skipped mails....");

                skipped_mails_dt = Program.g_db.get_skipped_mails(m_account.id);
                if (skipped_mails_dt == null || skipped_mails_dt.Rows == null || skipped_mails_dt.Rows.Count == 0)
                    return;

                MyLogger.Info($"({m_account.mail_address}) : skipped mail num = {skipped_mails_dt.Rows.Count}");
                MyLogger.Info($"({m_account.mail_address}) : venders num = {m_mail_vendor_handlers.Count}");

                string last_mail_folder = "";
                string mail_folder = "";
                IMailFolder box = null;
                foreach (DataRow row in skipped_mails_dt.Rows)
                {
                    if (cts.IsCancellationRequested)
                        return;

                    mail_folder = row["mail_folder"].ToString();
                    string subject = row["subject"].ToString().Trim();
                    string sender = row["sender"].ToString().Trim();
                    uint uid = uint.Parse(row["uid"].ToString().Trim());
                    int mail_order = 0;
                    int skipped_mail_id = int.Parse(row["id"].ToString().Trim());

                    if (sender != m_account.mail_address)
                    {
                        for (k = 0; k < m_mail_vendor_handlers.Count; k++)
                        {
                            if (m_mail_vendor_handlers[k].check_valid_mail(m_account.report_mode, subject, sender, out mail_order))
                                break;
                        }
                        if (k == m_mail_vendor_handlers.Count)
                        {
                            MyLogger.Info($"({m_account.mail_address}) : It's not gift card mail. id = {skipped_mail_id}, mail folder = {mail_folder}, mail uid = {uid}, subject = {subject}, sender = {sender}");
                        }
                        else
                        {
                            MyLogger.Info($"({m_account.mail_address}) : Found gift card mail. id = {skipped_mail_id}, mail folder = {mail_folder}, uid = {uid}, subject = {subject}, sender = {sender}");
                            MyLogger.Info($"({m_account.mail_address}) : mail type = {m_mail_vendor_handlers[k].GetType().ToString()}, order = {mail_order}");

                            if (last_mail_folder == "" || last_mail_folder != mail_folder)
                            {
                                box = client.GetFolder(mail_folder, cts.Token);
                                if (box == null)
                                {
                                    MyLogger.Info($"({m_account.mail_address}) : Can not open mail folder {mail_folder}");
                                    continue;
                                }
                                box.Open(FolderAccess.ReadOnly);
                                box.Subscribe();
                                last_mail_folder = mail_folder;
                            }

                            string mail_type = XMailHelper.get_mail_type_string(m_mail_vendor_handlers[k], mail_order);
                            download_card_mail(box, new UniqueId(uid), mail_type, cts);

                            Program.g_db.delete_skipped_mails(skipped_mail_id);
                        }
                    }
                    else
                    {
                        MyLogger.Info($"({m_account.mail_address}) : Skip sent mail. id = {skipped_mail_id}, mail folder = {mail_folder}, mail uid = {uid}, subject = {subject}, sender = {sender}");
                    }
                }

                disconnect(client, cts);
            }
            catch (OperationCanceledException exception)
            {
                 MyLogger.Info($"({System.Reflection.MethodBase.GetCurrentMethod().Name}): {nameof(OperationCanceledException)} thrown with message: {exception.Message}");

                if (client.IsConnected)
                    disconnect(client);
            }
            catch (Exception exception)
            {
                 MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");

                if (client.IsConnected)
                    disconnect(client, cts);
            }
        }
        private void download_card_mail(IMailFolder box, UniqueId uid, string mail_type, CancellationTokenSource cts)
        {
            MimeMessage message = box.GetMessage(uid, cts.Token);

            int message_hash = message.GetHashCode();
            string subject = XMailHelper.get_subject(message);
            string sender = XMailHelper.get_sender(message);
            string to = XMailHelper.get_mailto(message);
            DateTime date = DateTime.Parse(message.Date.ToString("yyyy-MM-dd HH:mm:ss"));

            MyLogger.Info($"({m_account.mail_address}), mail_folder = {box.FullName} fetch mail : from=({sender}), to=({to}), time=({date.ToString("yyyy-MM-dd HH:mm:ss")}), subject=({subject})");

            string mail_folder = Str_Utils.CleanPath(box.FullName);

            if (Program.g_db.is_already_fetched_mail(m_account.id, mail_folder, box.UidValidity, uid.Id))
            {
                MyLogger.Info($"({m_account.mail_address}), Already fetched. mail_folder = {box.FullName} fetch mail : box.UidValidity = {box.UidValidity}, uid.Id = {uid.Id}");
                return;
            }

            MyLogger.Info($"convert mail folder path : \"{box.FullName}\" -> \"{mail_folder}\"");

            string rel_path = get_mail_store_relative_path(mail_folder, date);
            Directory.CreateDirectory(rel_path);

            string eml_path = Path.Combine(rel_path, ConstEnv.LOCAL_MAIL_FILE_NAME);
            MyLogger.Info($"eml file path to save : {eml_path}");

            message.WriteTo(eml_path);

            string hash = XMailHelper.get_file_hash_md5(eml_path);
            if (Program.g_db.is_already_fetched_mail_by_hash(m_account.id, hash))
            {
                MyLogger.Info($"({m_account.mail_address}), Already fetched by hash. mail_folder = {box.FullName} fetch mail : box.UidValidity = {box.UidValidity}, hash = {hash}");
                return;
            }

            MyLogger.Info($"Saved mail to eml : {eml_path}");

            int new_id = Program.g_db.add_fetched_mail_to_db(m_account.id, mail_folder, box.UidValidity, uid.Id, hash, subject, sender, date, rel_path, mail_type);
            MyLogger.Info($"add mail to db : new id = {new_id}");

            if (new_id != -1)
                XMailParser.add_unchecked_mail_id(new_id);

            return;
        }
        public void download_mails_from_skipped_dt(DataTable dt, CancellationTokenSource cts)
        {
            try
            {
                List<uint> failed_uid_list = new List<uint>();
                List<string> failed_mail_folder_list = new List<string>();
                List<string> failed_sender_list = new List<string>();

                uint uid;
                string mail_folder;
                string sender;

                foreach (DataRow row in dt.Rows)
                {
                    uid = uint.Parse(row["uid"].ToString());
                    mail_folder = row["mail_folder"].ToString();
                    sender = row["sender"].ToString();

                    failed_uid_list.Add(uid);
                    failed_mail_folder_list.Add(mail_folder);
                    failed_sender_list.Add(sender);
                }

                ImapClient client = connect(cts);

                while (failed_uid_list.Count > 0)
                {
                    while (!client.IsConnected)
                    {
                        Thread.Sleep(120000);
                        client = connect(cts);
                        MyLogger.Info($"Try reconnect to mail server : connected = {client.IsConnected}");
                    }
                    uid = failed_uid_list[0];
                    mail_folder = failed_mail_folder_list[0];
                    sender = failed_sender_list[0];

                    string eml_file_path = Path.Combine("CancelMails", sender);
                    if (!Directory.Exists(eml_file_path))
                        Directory.CreateDirectory(eml_file_path);
                    eml_file_path = Path.Combine(eml_file_path, Str_Utils.CleanPath(mail_folder));
                    if (!Directory.Exists(eml_file_path))
                        Directory.CreateDirectory(eml_file_path);
                    eml_file_path = Path.Combine(eml_file_path, uid.ToString() + ".eml");

                    if (download_one_mail_directly(client, mail_folder, new MailKit.UniqueId(uid), eml_file_path))
                    {
                        MyLogger.Info($"Download a mail : mail_folder = {mail_folder}, sender = {sender}");
                        MyLogger.Info($"Remained number for this account : {failed_uid_list.Count - 1}");

                        failed_uid_list.RemoveAt(0);
                        failed_mail_folder_list.RemoveAt(0);
                        failed_sender_list.RemoveAt(0);
                    }
                    else
                    {
                        MyLogger.Info($"Failed download a mail : mail_folder = {mail_folder}, sender = {sender}");
                        MyLogger.Info($"Remained number for this account : {failed_uid_list.Count}");

                        failed_uid_list.RemoveAt(0);
                        failed_mail_folder_list.RemoveAt(0);
                        failed_sender_list.RemoveAt(0);

                        failed_uid_list.Add(uid);
                        failed_mail_folder_list.Add(mail_folder);
                        failed_sender_list.Add(sender);
                    }
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void redownload_fecthed_mails(List<string> mail_folders, List<uint> uids, CancellationTokenSource cts)
        {
            string old_mail_folder = "";
            IMailFolder box = null;

            try
            {
                ImapClient client = connect(cts);

                while (mail_folders.Count > 0)
                {
                    while (!client.IsConnected)
                    {
                        Thread.Sleep(120000);
                        client = connect(cts);
                        MyLogger.Info($"Try reconnect to mail server : connected = {client.IsConnected}");
                        old_mail_folder = "";
                    }
                    uint uid = uids[0];
                    string mail_folder = mail_folders[0];

                    if (mail_folder != old_mail_folder)
                    {
                        box = client.GetFolder(mail_folder);
                        if (box == null)
                        {
                            continue;
                        }
                        old_mail_folder = mail_folder;
                    }

                    box.Open(FolderAccess.ReadOnly);
                    box.Subscribe();

                    download_card_mail(box, new MailKit.UniqueId(uid), cts);

                    mail_folders.RemoveAt(0);
                    uids.RemoveAt(0);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        private void download_card_mail(IMailFolder box, UniqueId uid, CancellationTokenSource cts)
        {
            MimeMessage message = box.GetMessage(uid, cts.Token);

            string subject = XMailHelper.get_subject(message);
            string sender = XMailHelper.get_sender(message);
            string to = XMailHelper.get_mailto(message);
            DateTime date = DateTime.Parse(message.Date.ToString("yyyy-MM-dd HH:mm:ss"));

            int k;
            int mail_order = 0;
            string mail_type = "";
            for (k = 0; k < m_mail_vendor_handlers.Count; k++)
            {
                if (m_mail_vendor_handlers[k].check_valid_mail(m_account.report_mode, subject, sender, out mail_order))
                    break;
            }
            if (k == m_mail_vendor_handlers.Count)
            {
                MyLogger.Info($"({m_account.mail_address}) : It's not gift card mail. mail folder = {box.FullName}, mail uid = {uid.Id}, subject = {subject}, sender = {sender}");
                return;
            }
            else
            {
                mail_type = XMailHelper.get_mail_type_string(m_mail_vendor_handlers[k], mail_order);
            }

            MyLogger.Info($"({m_account.mail_address}), mail_folder = {box.FullName} fetch mail : from=({sender}), to=({to}), time=({date.ToString("yyyy-MM-dd HH:mm:ss")}), subject=({subject})");

            string mail_folder = Str_Utils.CleanPath(box.FullName);

            if (Program.g_db.is_already_fetched_mail(m_account.id, mail_folder, box.UidValidity, uid.Id))
            {
                MyLogger.Info($"({m_account.mail_address}), Already fetched. mail_folder = {box.FullName} fetch mail : box.UidValidity = {box.UidValidity}, uid.Id = {uid.Id}");
                return;
            }

            MyLogger.Info($"convert mail folder path : \"{box.FullName}\" -> \"{mail_folder}\"");

            string rel_path = get_mail_store_relative_path(mail_folder, date);
            Directory.CreateDirectory(rel_path);

            string eml_path = Path.Combine(rel_path, ConstEnv.LOCAL_MAIL_FILE_NAME);
            MyLogger.Info($"eml file path to save : {eml_path}");

            message.WriteTo(eml_path);

            string hash = XMailHelper.get_file_hash_md5(eml_path);
            if (Program.g_db.is_already_fetched_mail_by_hash(m_account.id, hash))
            {
                MyLogger.Info($"({m_account.mail_address}), Already fetched by hash. mail_folder = {box.FullName} fetch mail : box.UidValidity = {box.UidValidity}, hash = {hash}");
                return;
            }

            string rel_path1 = "Messages1" + rel_path.Substring("Messages".Length);
            Directory.CreateDirectory(rel_path1);

            string eml_path1 = Path.Combine(rel_path1, ConstEnv.LOCAL_MAIL_FILE_NAME);

            File.Copy(eml_path, eml_path1);

            MyLogger.Info($"Saved mail to eml : {eml_path}");

            int new_id = Program.g_db.add_fetched_mail_to_db(m_account.id, mail_folder, box.UidValidity, uid.Id, hash, subject, sender, date, rel_path, mail_type);
            MyLogger.Info($"add mail to db : new id = {new_id}");

            if (new_id != -1)
                XMailParser.add_unchecked_mail_id(new_id);

            return;
        }
        private bool download_one_mail_directly(ImapClient client, string mail_folder, UniqueId uid, string download_file)
        {
            try
            {
                IMailFolder box = client.GetFolder(mail_folder);
                if (box == null)
                    return false;

                box.Open(FolderAccess.ReadOnly);
                box.Subscribe();

                MimeMessage message = box.GetMessage(uid);
                message.WriteTo(download_file);

                return true;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            return false;
        }
        private string get_mail_store_relative_path(string mail_folder, DateTime date)
        {
            string rel_path = "";
            string root_path = ConstEnv.get_local_mail_root_path();
            string parent_path = Path.Combine(root_path, m_account.mail_address);

            parent_path = Path.Combine(parent_path, mail_folder);
            parent_path = Path.Combine(parent_path, date.ToString("yyyy-MM-dd"));

            int i = 1;
            while (true)
            {
                rel_path = Path.Combine(parent_path, i.ToString());
                if (!Directory.Exists(rel_path))
                    break;
                i++;
            }
            return rel_path;
        }
    }
    public class MailId
    {
        public readonly string folder_name;
        public readonly uint validity;
        public readonly uint unique_id;

        public MailId(string _folder_name, uint _validity, uint _unique_id)
        {
            folder_name = _folder_name;
            validity = _validity;
            unique_id = _unique_id;
        }
    }
}
