using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace MailHelper
{
    public class KMailCR1 : KMailBaseCR
    {
        public KMailCR1() : base()
        {
        }
        #region override functions
        public override bool check_valid_mail(int work_mode, string subject, string sender, out int mail_seq_num)
        {
            mail_seq_num = 0;

            if (Program.g_user.is_report_mode_gs(work_mode))
                return false;

            if (sender == "support@cardcash.com" && subject == "CardCash Order Confirmation")
            {
                mail_seq_num = 1;
                return true;
            }
            else if (sender == "sales@cardcash.com" && subject.StartsWith("Your CardCash Order") && subject.EndsWith("eGift card has arrived!"))
            {
                mail_seq_num = 2;
                return true;
            }
            else if (sender == "sales@cardcash.com" && subject.StartsWith("Your CardCash Order") && subject.EndsWith("spreadsheet is attached"))
            {
                mail_seq_num = 3;
                return true;
            }

            return false;
        }
        protected override bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_report)
        {
            try
            {
                KReportCR1 report = base_report as KReportCR1;

                if (mail_order == 1)
                    parse_mail_cr_1_1(mail, report);
                else if (mail_order == 2)
                    parse_mail_cr_1_2(mail, report);
                else if (mail_order == 3)
                    parse_mail_cr_1_3(mail, report);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return false;
            }
            return true;
        }
        #endregion override functions

        #region class specific functions
        private void parse_mail_cr_1_1(MimeMessage mail, KReportCR1 report)
        {
            string str_chk_confirm = "Your confirmation number is #";

            report.m_purchase_date = XMailHelper.get_sentdate(mail);
            MyLogger.Info($"... 1st mail date = {report.m_purchase_date.ToString()}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.IndexOf(str_chk_confirm, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string w_str_order = line.Substring(line.IndexOf(str_chk_confirm, StringComparison.InvariantCultureIgnoreCase) + str_chk_confirm.Length);
                    w_str_order = w_str_order.Replace(" ", string.Empty);
                    report.set_order_id(w_str_order.Replace("\r", string.Empty));
                    MyLogger.Info($"... 1st mail order = {report.m_order_id}");
                }

                if (line.IndexOf("Payment Method:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string payment_type = "";
                    string last_4_digit = "";
                    string temp = line.Substring(line.IndexOf("Payment Method:", StringComparison.InvariantCultureIgnoreCase) + "Payment Method:".Length).Trim();
                    if (temp.IndexOf("(") > -1)
                    {
                        payment_type = temp.Substring(0, temp.IndexOf("(")).Trim();
                        int idx = line.IndexOf("Last 4 of CC:", StringComparison.InvariantCultureIgnoreCase);
                        if (idx != -1)
                        {
                            temp = line.Substring(idx + "Last 4 of CC:".Length);
                            temp = temp.Substring(0, temp.IndexOf(")"));
                            last_4_digit = temp.Trim();
                        }
                    }
                    else
                    {
                        payment_type = temp;
                    }
                    report.m_cr_payment_type = payment_type;
                    report.m_cr_payment_id = last_4_digit;
                    MyLogger.Info($"... 1st mail payment_type = {payment_type}, payment_id = {last_4_digit}");
                }

                if (line.IndexOf("Total:", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    string temp = line.Substring(line.IndexOf("Total:", StringComparison.InvariantCultureIgnoreCase) + 6);
                    temp = temp.Replace("\r", string.Empty);
                    temp = temp.Replace(" ", "");
                    report.set_total(Str_Utils.string_to_float(temp.Substring(1)));
                    MyLogger.Info($"...1st mail total = {report.m_total}");
                }
            }
        }
        private void parse_mail_cr_1_2(MimeMessage mail, KReportCR1 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (!subject.StartsWith("Your CardCash Order ") || !subject.EndsWith(" eGift card has arrived!"))
                throw new Exception($"Invalid Vendor1 2nd mail subject. {subject}");

            string order = subject.Substring("Your CardCash Order ".Length);
            order = order.Substring(0, order.Length - " eGift card has arrived!".Length);
            order = order.Trim();
            report.set_order_id(order);
            MyLogger.Info($"... 2nd mail order = {report.m_order_id}");
        }
        private void parse_mail_cr_1_3(MimeMessage mail, KReportCR1 report)
        {
            string subject = XMailHelper.get_subject(mail);

            if (!subject.StartsWith("Your CardCash Order ") || !subject.EndsWith(" spreadsheet is attached"))
                throw new Exception($"Invalid Vendor1 2nd mail subject. {subject}");

            string order = subject.Substring("Your CardCash Order ".Length);
            order = order.Substring(0, order.Length - " spreadsheet is attached".Length);
            order = order.Trim();
            report.set_order_id(order);
            MyLogger.Info($"... 3rd mail order = {report.m_order_id}");

            foreach (MimeEntity att in mail.Attachments)
            {
                string fileName = att.ContentDisposition?.FileName ?? att.ContentType.Name;
                string att_type = Path.GetExtension(fileName);

                if (!att_type.Equals(ConstEnv.MAIL_ATTACH_TYPE_CSV, StringComparison.InvariantCultureIgnoreCase))
                    continue;

                MyLogger.Info($"... Found csv in the 3rd mail. name = {fileName}");

                string csv_att_fpath = Path.Combine(m_temp_path, fileName);
                using (var stream = File.Create(csv_att_fpath))
                {
                    if (att is MessagePart)
                    {
                        var rfc822 = (MessagePart)att;
                        rfc822.Message.WriteTo(stream);
                    }
                    else
                    {
                        var part = (MimePart)att;
                        part.Content.DecodeTo(stream);
                    }
                }

                DataTable dt_att_csv = CSVUtil.csv2dt(csv_att_fpath);
                if (dt_att_csv == null)
                {
                    MyLogger.Error($"... Failed to parse csv in the 3rd mail. name = {fileName}");
                    continue;
                }

/*
                string[] valid_col_names = new string[] {
                    "Order_ID", "Merchant", "Card_ID", "Filing_ID", "Face_Value", "Purchase_price", "Number", "Pin", "Card_Type"
                };
                if (dt_att_csv.Columns.Count != valid_col_names.Length)
                    throw new Exception($"Invalid vendor 1 CSV column count. {dt_att_csv.Columns.Count} != {valid_col_names.Length}");
                for (int i = 0; i < dt_att_csv.Columns.Count; i++)
                {
                    if (dt_att_csv.Columns[i].ColumnName != valid_col_names[i])
                        throw new Exception($"Invalid vendor 1 CSV column name. {dt_att_csv.Columns[i].ColumnName} != {valid_col_names[i]}");
                }
*/

                foreach (DataRow row in dt_att_csv.Rows)
                {
                    int idx = dt_att_csv.Columns.IndexOf("Order_ID");
                    if (idx == -1)
                        break;
                    string order_row = row[idx].ToString();
                    if (order_row != report.m_order_id)
                    {
                        MyLogger.Info($"Invalid order in csv row. {order_row} != {report.m_order_id}");
                        continue;
                    }

                    string merchant = "";
                    string face_value = "";
                    string cost = "";
                    string gift = "";
                    string pin = "";

                    idx = dt_att_csv.Columns.IndexOf("Merchant");
                    if (idx != -1)
                        merchant = row[idx].ToString().Trim();
                    idx = dt_att_csv.Columns.IndexOf("Face_Value");
                    if (idx != -1)
                        face_value = row[idx].ToString().Trim();
                    idx = dt_att_csv.Columns.IndexOf("Purchase_price");
                    if (idx != -1)
                        cost = row[idx].ToString().Trim();
                    idx = dt_att_csv.Columns.IndexOf("Number");
                    if (idx != -1)
                    {
                        gift = row[idx].ToString().Trim();
                    }
                    else
                    {
                        idx = dt_att_csv.Columns.IndexOf("Card #");
                        if (idx != -1)
                            gift = row[idx].ToString().Trim();
                    }
                    idx = dt_att_csv.Columns.IndexOf("Pin");
                    if (idx != -1)
                        pin = row[idx].ToString().Trim();

                    MyLogger.Info($"...... csv info. retailer  = {merchant}");
                    MyLogger.Info($"...... csv info. value     = {face_value}");
                    MyLogger.Info($"...... csv info. cost      = {cost}");
                    MyLogger.Info($"...... csv info. gift card = {gift}");
                    MyLogger.Info($"...... csv info. pin       = {pin}");

                    report.add_giftcard_details(merchant, Str_Utils.string_to_currency(face_value), Str_Utils.string_to_currency(cost), gift, pin);
                }
                
                File.Delete(csv_att_fpath);
            }
        }
        #endregion class specific functions
    }
}
