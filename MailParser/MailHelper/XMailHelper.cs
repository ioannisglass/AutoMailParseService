using BaseModule;
using DbHelper;
using MailParser;
using Logger;
using MailKit;
using MimeKit;
using MimeKit.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using UserHelper;

namespace MailHelper
{
    public class XMailHelper
    {
        public XMailHelper()
        {
        }
        public void start_fetch_mails(CancellationTokenSource cts)
        {
            MyLogger.Info($">>> start_fetch_mails");

            try
            {
                List<Task> task_list = new List<Task>();

                List<UserInfo> user_list = Program.g_user.user_info_list;
                MyLogger.Info($"user count = {user_list.Count}");

                Dictionary<UserInfo, DateTime> last_connect_times = new Dictionary<UserInfo, DateTime>();
                foreach (UserInfo user in user_list)
                    last_connect_times.Add(user, DateTime.MinValue);

                for (int i = 0; i < user_list.Count; i++)
                {
                    int index = i;
                    Task.Run(() =>
                    {
                        UserInfo user = user_list[index];
                        XMailChecker mail_checker = null;

                        while (!cts.IsCancellationRequested)
                        {
                            DateTime last_time = last_connect_times[user];
                            while (last_time != DateTime.MinValue && ((DateTime.Now - last_time).TotalSeconds) < Program.g_setting.mail_download_retry_second)
                                Thread.Sleep(1000);

                            mail_checker = new XMailChecker(user);
                            mail_checker.check_mails(cts);

                            last_connect_times[user] = DateTime.Now;
                        }
                    });
                }
            }
            catch (OperationCanceledException exception)
            {
                if (cts.IsCancellationRequested)
                {
                    MyLogger.Info("Cancelled.");
                }
                else
                {
                    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void extract_mails_from_skipped_mails(CancellationTokenSource cts)
        {
            MyLogger.Info($">>> extract_canclled_mails_from_skipped_mails");

            try
            {
                List<Task> task_list = new List<Task>();

                List<UserInfo> user_list = Program.g_db.get_user_list();
                if (user_list == null || user_list.Count == 0)
                    return;
                MyLogger.Info($"user count = {user_list.Count}");

                for (int i = 0; i < user_list.Count; i++)
                {
                    if (cts.IsCancellationRequested)
                        break;

                    UserInfo user = user_list[i];
                    XMailChecker mail_checker = new XMailChecker(user);
                    mail_checker.extract_mails_from_skipped_mails(cts);
                }
            }
            catch (OperationCanceledException exception)
            {
                if (cts.IsCancellationRequested)
                {
                    MyLogger.Info("Cancelled.");
                }
                else
                {
                    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message}");
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            MyLogger.Info($"<<< extract_canclled_mails_from_skipped_mails");
        }

        #region Get Mail Parts Functions
        static protected bool is_forward_mail(MimeMessage mail)
        {
            string subject = mail.Subject;
            subject = subject.ToLower();
            if (subject.StartsWith("fw:"))
                return true;
            return false;
        }

        static public string get_subject(MimeMessage mail)
        {
            string subject = mail.Subject;

            if (!is_forward_mail(mail))
            {
                return subject.Trim();
            }
            else
            {

            }

            string[] lines = get_bodytext(mail).Replace("\r", "").Split('\n');
            foreach (string line in lines)
            {
                if (line.StartsWith("Subject:"))
                {
                    subject = line.Substring("Subject:".Length);
                    subject = subject.Trim();
                    break;
                }
                if (line.StartsWith("*Subject:*"))
                {
                    subject = line.Substring("*Subject:*".Length);
                    subject = subject.Trim();
                    break;
                }
            }
            return subject.Trim();
        }
        static public string get_subject_from_envelop(Envelope envelop)
        {
            string subject = "";

            if (envelop == null)
                return subject;

            subject = envelop.Subject;
            if (subject == null)
                subject = "";
            return subject.Trim();
        }
        static public string get_sender(MimeMessage mail)
        {
            string sender = "";

            if (!is_forward_mail(mail))
            {
                if (mail.Sender != null)
                {
                    sender = mail.Sender.Address;
                }
                else if (mail.From != null && mail.From.Count > 0)
                {
                    foreach (MailboxAddress from in mail.From)
                    {
                        if (sender == "")
                            sender = from.Address;
                        else
                            sender += "|" + from.Address;
                    }
                }
                return sender.Trim();
            }
            else
            {
                string[] lines = get_bodytext(mail).Replace("\r", "").Split('\n');
                foreach (string line in lines)
                {
                    if (line.StartsWith("From:"))
                    {
                        sender = line.Substring("From:".Length);
                        sender = sender.Trim();
                        if (sender.IndexOf("<") != -1)
                        {
                            sender = sender.Substring(sender.LastIndexOf("<") + 1);
                            if (sender.IndexOf(">") != -1)
                                sender = sender.Substring(0, sender.LastIndexOf(">"));
                            sender = sender.Trim();
                        }
                        break;
                    }
                    if (line.StartsWith("*From:*"))
                    {
                        sender = line.Substring("*From:*".Length);
                        sender = sender.Trim();
                        if (sender.IndexOf("<") != -1)
                        {
                            sender = sender.Substring(sender.LastIndexOf("<") + 1);
                            if (sender.IndexOf(">") != -1)
                                sender = sender.Substring(0, sender.LastIndexOf(">"));
                            sender = sender.Trim();
                        }
                        if (sender.IndexOf("[") != -1)
                        {
                            sender = sender.Substring(sender.LastIndexOf("[") + 1);
                            if (sender.IndexOf("]") != -1)
                                sender = sender.Substring(0, sender.LastIndexOf("]"));
                            sender = sender.Trim();
                        }
                        break;
                    }
                }
                return sender.Trim();
            }
        }
        static public string get_sender_from_envelop(Envelope envelop)
        {
            string sender = "";

            if (envelop == null)
                return sender;

            if (envelop.Sender != null && envelop.Sender.Count > 0)
            {
                foreach (MailboxAddress from in envelop.Sender)
                {
                    if (sender == "")
                        sender = from.Address;
                    else
                        sender = sender + "|" + from.Address;
                }
            }
            else if (envelop.From != null && envelop.From.Count > 0)
            {
                foreach (MailboxAddress from in envelop.From)
                {
                    if (sender == "")
                        sender = from.Address;
                    else
                        sender = sender + "|" + from.Address;
                }
            }
            if (sender == null)
                sender = "";
            return sender.Trim();
        }
        static public string get_mailto(MimeMessage mail)
        {
            string mailto = "";

            if (mail.To != null && mail.To.Count > 0)
            {
                foreach (MailboxAddress to in mail.To)
                {
                    if (mailto == "")
                        mailto = to.Address;
                    else
                        mailto += "|" + to.Address;
                }
            }
            return mailto.Trim();
        }
        static public DateTime get_sentdate(MimeMessage mail)
        {
            if (!is_forward_mail(mail))
            {
                return DateTime.Parse(mail.Date.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            else
            {
                DateTime sent = DateTime.MinValue;
                string[] lines = get_bodytext(mail).Replace("\r", "").Split('\n');
                foreach (string line in lines)
                {
                    if (line.StartsWith("Sent:"))
                    {
                        string sent_str = line.Substring("Sent:".Length);
                        sent_str = sent_str.Trim();
                        sent = DateTime.Parse(sent_str);
                        break;
                    }
                }
                return sent;
            }

        }
        static public DateTime get_sentdate_from_envelop(Envelope envelop)
        {
            if (envelop.Date.HasValue)
                return DateTime.Parse(envelop.Date.Value.ToString("yyyy-MM-dd HH:mm:ss"));
            return DateTime.MinValue;
        }
        static protected string extract_text_from_html_by_regex(string htmltext)
        {
            if (htmltext == null || htmltext == "")
                return "";
            Regex reg1 = new Regex("<[^>]*>");
            Regex reg2 = new Regex("<style>[^<]*</style>");
            htmltext = reg2.Replace(htmltext, "");
            string body_text = reg1.Replace(htmltext, "");
            body_text = body_text.Replace("\u200B", "");
            return body_text;
        }
        static protected string extract_text_from_html_by_structure(string htmltext)
        {
            if (htmltext == null || htmltext == "")
                return "";
            using (var writer = new StringWriter())
            {
                using (var reader = new StringReader(htmltext))
                {
                    var tokenizer = new HtmlTokenizer(reader)
                    {
                        DecodeCharacterReferences = true
                    };
                    HtmlToken token;

                    while (tokenizer.ReadNextToken(out token))
                    {
                        switch (token.Kind)
                        {
                            case HtmlTokenKind.Data:
                                var data = (HtmlDataToken)token;
                                writer.Write(data.Data);
                                break;
                            case HtmlTokenKind.Tag:
                                var tag = (HtmlTagToken)token;
                                switch (tag.Id)
                                {
                                    case HtmlTagId.Br:
                                        writer.Write(Environment.NewLine);
                                        break;
                                    case HtmlTagId.P:
                                        if (tag.IsEndTag || tag.IsEmptyElement)
                                            writer.Write(Environment.NewLine);
                                        break;
                                }
                                break;
                        }
                    }
                }
                string body_text = writer.ToString();
                body_text = body_text.Replace("\u200B", "");
                return body_text;
            }
        }
        static public string html2text(string htmltext)
        {
            //string body_text = extract_text_from_html_by_structure(htmltext);
            string body_text = extract_text_from_html_by_regex(htmltext);
            if (body_text == "")
                return "";
            string body_text_ret = "";
            string[] lines = body_text.Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line == "" || line.ToLower() == "&nbsp" || line.ToLower() == "&nbsp;")
                    continue;
                body_text_ret += line + "\n";
            }

            Tuple<string, string>[] replace_keywords = new Tuple<string, string>[]
            {
                new Tuple<string, string>("&nbsp;", ""),
                new Tuple<string, string>("&nbsp", ""),
                new Tuple<string, string>("&NBSP", ""),
                new Tuple<string, string>("&#8203;", ""),
                new Tuple<string, string>("&ensp;", ""),
                new Tuple<string, string>("&emsp;", ""),
                new Tuple<string, string>("&thinsp;", ""),
                new Tuple<string, string>("&zwnj;", ""),
                new Tuple<string, string>("&zwj;", ""),
                new Tuple<string, string>("&lrm;", ""),
                new Tuple<string, string>("&rlm;", ""),
                new Tuple<string, string>("&ndash;", "–"),
                new Tuple<string, string>("&mdash;", "—"),
                new Tuple<string, string>("&lsquo;", "‘"),
                new Tuple<string, string>("&rsquo;", "’"),
                new Tuple<string, string>("&sbquo;", "‚"),
                new Tuple<string, string>("&ldquo;", "“"),
                new Tuple<string, string>("&rdquo;", "”"),
                new Tuple<string, string>("&rdquo;", "„"),
                new Tuple<string, string>("&dagger;", "†"),
                new Tuple<string, string>("&Dagger;", "‡"),
                new Tuple<string, string>("&bull;", "•"),
                new Tuple<string, string>("&hellip;", "…"),
                new Tuple<string, string>("&permil;", "‰"),
                new Tuple<string, string>("&prime;", "′"),
                new Tuple<string, string>("&Prime;", "″"),
                new Tuple<string, string>("&lsaquo;", "‹"),
                new Tuple<string, string>("&rsaquo;", "›"),
                new Tuple<string, string>("&oline;", "‾"),
                new Tuple<string, string>("&frasl;", "⁄"),
                new Tuple<string, string>("&euro;", "€"),
                new Tuple<string, string>("&loz;", "◊"),
                new Tuple<string, string>("&spades;", "♠"),
                new Tuple<string, string>("&clubs;", "♣"),
                new Tuple<string, string>("&hearts;", "♥"),
                new Tuple<string, string>("&diams;", "♦"),

                new Tuple<string, string>("&amp;", "&"),
                new Tuple<string, string>("&AMP;", "&"),
                new Tuple<string, string>("&lt;", "<"),
                new Tuple<string, string>("&LT;", "<"),
                new Tuple<string, string>("&gt;", ">"),
                new Tuple<string, string>("&GT;", ">"),
                new Tuple<string, string>("&quot;", "\""),
                new Tuple<string, string>("&QUOT;", "\""),
                new Tuple<string, string>("&copy;", "©"),
                new Tuple<string, string>("&COPY;", "©"),
                new Tuple<string, string>("&trade;", "™"),
                new Tuple<string, string>("&TRADE;", "™"),
                new Tuple<string, string>("&reg;", "®"),
                new Tuple<string, string>("&REG;", "®"),


                new Tuple<string, string>("\u200B", ""),
                new Tuple<string, string>("\u200C", "")
           };

            foreach (var r in replace_keywords)
                body_text_ret = body_text_ret.Replace(r.Item1, r.Item2);

            body_text_ret = body_text_ret.Trim();
            return body_text_ret;
        }
        static public string get_bodytext(MimeMessage mail)
        {
            if (mail != null)
            {
                if (mail.TextBody == null)
                {
                    MyLogger.Info($"CAUTION!!! TextBody IS NULL, WILL USE htmltext");
                }
                string body = mail.TextBody;
                if (body != null)
                    return body;
                if (mail.HtmlBody != null)
                    return html2text(mail.HtmlBody);
            }
            return "";
        }
        static public string get_bodytext2(MimeMessage mail)
        {
            if (mail != null && mail.HtmlBody != null)
                return html2text(mail.HtmlBody);
            return "";
        }
        static public bool is_bodytext_existed(MimeMessage mail)
        {
            if (mail.TextBody == null)
                MyLogger.Info($"CAUTION!!! TextBody IS NULL, WILL USE htmltext");
            else if (mail.TextBody == "")
                MyLogger.Info($"CAUTION!!! TextBody IS EMPTY, WILL USE htmltext");
            return (mail.TextBody != null && mail.TextBody != "");
        }
        static public string get_htmltext(MimeMessage mail)
        {
            return mail.HtmlBody;
        }
        static public string find_html_part(string src, string html_tag, out int next_pos)
        {
            next_pos = -1;
            string begin_tag = "<" + html_tag + ">";
            string begin_tag1 = "<" + html_tag + " ";
            string end_tag = "</" + html_tag + ">";
            string temp = src;
            int find_b;
            int find_b1 = temp.IndexOf(begin_tag);
            int find_b2 = temp.IndexOf(begin_tag1);
            if (find_b1 == -1 && find_b2 == -1)
                return "";
            find_b1 = (find_b1 == -1) ? int.MaxValue : find_b1;
            find_b2 = (find_b2 == -1) ? int.MaxValue : find_b2;
            find_b = Math.Min(find_b1, find_b2);
            find_b = temp.IndexOf(">", find_b) + 1;

            int find_e = -1;
            int first_find_b = find_b;
            int found_begin_tag = 1;
            while (found_begin_tag != 0)
            {
                find_e = temp.IndexOf(end_tag, find_b);
                if (find_e == -1)
                    return "";
                found_begin_tag--;

                string temp1 = temp.Substring(find_b, find_e - find_b);
                while (temp1.IndexOf(begin_tag) != -1 || temp1.IndexOf(begin_tag1) != -1)
                {
                    int b1 = temp1.IndexOf(begin_tag);
                    int b2 = temp1.IndexOf(begin_tag1);
                    b1 = (b1 == -1) ? int.MaxValue : b1;
                    b2 = (b2 == -1) ? int.MaxValue : b2;
                    int b = Math.Min(b1, b2);
                    found_begin_tag++;
                    temp1 = temp1.Substring(b);
                    temp1 = temp1.Substring(temp1.IndexOf(">"));
                }
                find_b = find_e + end_tag.Length;
            }
            if (find_e == -1)
                return "";
            temp = temp.Substring(first_find_b, find_e - first_find_b);
            next_pos = find_e + end_tag.Length;

            return temp.Trim();
        }

        static public string get_text(string src, string html_tag)
        {
            string ret = "";
            int pos = 0;

            ret = find_html_part(src, html_tag, out pos);

            while (ret.IndexOf("<") != -1 && ret.IndexOf(">") != -1)
            {
                ret = ret.Remove(ret.IndexOf("<"), ret.IndexOf(">") - ret.IndexOf("<") + 1);
            }

            return ret.Trim();
        }

        static public string exclude_tags(string src)
        {
            string ret = src;

            while (ret.IndexOf("<") != -1 && ret.IndexOf(">") != -1)
            {
                ret = ret.Remove(ret.IndexOf("<"), ret.IndexOf(">") - ret.IndexOf("<") + 1);
            }
            return ret;
        }

        static public string trim_link(string src_link)
        {
            string link = src_link.Substring(src_link.IndexOf('<') + 1);
            link = link.Trim();
            link = link.Substring(0, link.IndexOf('>'));
            link = link.Trim();
            return link;
        }
        static public string get_mail_type_string(KMailBaseParser parser_class, int order)
        {
            string vendor = "";

            if (parser_class.GetType() == typeof(KMailCR1))
                vendor = "CR_1";
            else if (parser_class.GetType() == typeof(KMailCR2))
                vendor = "CR_2";
            else if (parser_class.GetType() == typeof(KMailCR3))
                vendor = "CR_3";
            else if (parser_class.GetType() == typeof(KMailCR4))
                vendor = "CR_4";
            else if (parser_class.GetType() == typeof(KMailCR7))
                vendor = "CR_7";
            else if (parser_class.GetType() == typeof(KMailBaseOP))
                vendor = "OP";
            else if (parser_class.GetType() == typeof(KMailBaseSC))
                vendor = "SC";
            else if (parser_class.GetType() == typeof(KMailBaseCC))
                vendor = "CC";
            else
                return "";

            return vendor + "_" + order.ToString();
        }
        static public string get_mail_type_string(KReportBase card, int order)
        {
            return card.m_mail_type.ToString();
        }
        static public string revise_concated_bodytext(string bodytext, int max_blank_len = 6)
        {
            string revised_bodytext = "";
            int start_pos;
            string temp = bodytext;

            start_pos = temp.IndexOf(" ");
            while (start_pos != -1)
            {
                revised_bodytext += temp.Substring(0, start_pos);

                int i = start_pos;
                while (i < temp.Length && temp[i] == ' ')
                    i++;
                int len = i - start_pos;
                if (len > max_blank_len)
                    revised_bodytext += "\n";
                else
                    revised_bodytext += temp.Substring(start_pos, len);

                temp = temp.Substring(i);
                start_pos = temp.IndexOf(" ");
            }

            return revised_bodytext;
        }
        static public string get_address_state_name(string address)
        {
            /**
             *      Address Example:
             *      
             *          House number and street name + Apartment/Suite/Room number if any           Jeremy Martinson, Jr.
             *          Name of town + State abbreviation + ZIP code                                455 Larkspur Dr. Apt 23
             *          (typical handwritten format)                                                Baviera, CA 92908
             *          
             *          House number and street name +Apartment/Suite/Room number if any            JEREMY MARTINSON JR
             *          Name of town + State abbreviation + ZIP+4 code                              455 LARKSPUR DR APT 23
             *          (USPS-recommended format)                                                   BAVIERA CA 92908‑4601
             * 
             **/

            string state = "";

            string temp = address.Trim();
            if (temp.IndexOf(" ") == -1)
                return "";
            if (temp.EndsWith("USA", StringComparison.CurrentCultureIgnoreCase))
                temp = temp.Substring(0, temp.Length - "USA".Length).Trim();
            if (temp.EndsWith("UNITED STATES", StringComparison.CurrentCultureIgnoreCase))
                temp = temp.Substring(0, temp.Length - "UNITED STATES".Length).Trim();
            string zip = temp.Substring(temp.LastIndexOf(" ") + 1);
            if (!Regex.Match(zip, @"^\d{5}(?:[-\s]\d{4})?$").Success && !Regex.Match(zip, @"^\d{5}(?:\d{4})?$").Success) // some address does not contain hyphen. such as 087014528
                return "";
            temp = temp.Substring(0, temp.LastIndexOf(" ")).Trim();
            if (temp.IndexOf(",") == -1)
            {
                if (temp.IndexOf(" ") == -1)
                    return "";
                temp = temp.Substring(temp.LastIndexOf(" ") + 1).Trim();
            }
            else
            {
                temp = temp.Substring(temp.LastIndexOf(",") + 1).Trim();
            }
            if (temp.EndsWith("."))
                temp = temp.Substring(0, temp.Length - 1);
            var state_names = ConstEnv.STATE_NAMES.Where(s => s.abbr_name == temp.ToUpper() || s.full_name.ToUpper() == temp.ToUpper());
            if (state_names.Count() == 0)
                return "";
            state = state_names.ElementAt(0).abbr_name;
            return state;
        }
        static public string get_file_hash_md5(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filename))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToUpperInvariant();
                }
            }
        }

        #endregion Get Mail Parts Functions
    }
}
