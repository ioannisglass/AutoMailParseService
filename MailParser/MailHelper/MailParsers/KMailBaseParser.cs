using MailParser;
using Logger;
using MimeKit;
using MimeKit.Text;
using ReportStat;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using UserHelper;
using WebAuto;

namespace MailHelper
{
    public class KMailBaseParser
    {
        #region Parsing Information
        protected object m_lock_card_list;
        protected List<KReportBase> m_card_list;
        protected string m_temp_path; // for vendor 1
        #endregion Parsing Information

        public KMailBaseParser()
        {
            m_lock_card_list = new object();
            m_card_list = new List<KReportBase>();
        }
        #region virtual functions.
        public bool check_valid_mail(int work_mode, MimeMessage mail, out int mail_order)
        {
            string subject = XMailHelper.get_subject(mail);
            string sender = XMailHelper.get_sender(mail);
            return check_valid_mail(work_mode, subject, sender, out mail_order);
        }
        public virtual bool check_valid_mail(int work_mode, string subject, string sender, out int mail_order)
        {
            mail_order = 0;
            return false;
        }
        protected virtual bool parse_mail(MimeMessage mail, int mail_order, KReportBase base_card)
        {
            return false;
        }
        #endregion virtual functions.
        #region DB handler virtual functions.
        public bool parse(int mail_id, string eml_file)
        {
            try
            {
                int account_id = Program.g_db.get_account_id_from_mail_id(mail_id);
                UserInfo user = Program.g_user.get_user_by_id(account_id);
                if (user == null)
                    return false;

                MimeMessage mail = MimeMessage.Load(eml_file);

                // check if it is a valid mail.

                int mail_seq_num = 0;
                if (!check_valid_mail(user.report_mode, mail, out mail_seq_num))
                    return false;

                m_temp_path = Path.GetDirectoryName(eml_file);

                KReportBase tmp_report = null;
                if (this.GetType() == typeof(KMailCR1))
                    tmp_report = new KReportCR1(mail_seq_num);
                else if (this.GetType() == typeof(KMailCR2))
                    tmp_report = new KReportCR2(mail_seq_num);
                else if (this.GetType() == typeof(KMailCR3))
                    tmp_report = new KReportCR3(mail_seq_num);
                else if (this.GetType() == typeof(KMailCR4))
                    tmp_report = new KReportCR4(mail_seq_num);
                else if (this.GetType() == typeof(KMailCR7))
                    tmp_report = new KReportCR7(mail_seq_num);
                else if (this.GetType() == typeof(KMailCR8))
                    tmp_report = new KReportCR8(mail_seq_num);
                else if (this.GetType() == typeof(KMailBaseOP))
                    tmp_report = new KReportOP();
                else if (this.GetType() == typeof(KMailBaseSC))
                    tmp_report = new KReportSC();
                else if (this.GetType() == typeof(KMailBaseCC))
                    tmp_report = new KReportCC();
                else
                    return false;

                tmp_report.m_mail_sent_date = XMailHelper.get_sentdate(mail);

                MyLogger.Info($"*** CARD MAIL **** mail id = {mail_id}, mail_seq_num = {mail_seq_num}, type = {this.GetType()}, sentdate = {tmp_report.m_mail_sent_date}");

                // set mail id and account id.

                tmp_report.m_mail_id = mail_id;
                tmp_report.m_mail_account_id = Program.g_db.get_account_id_from_mail_id(mail_id);
                if (tmp_report.m_mail_account_id == -1)
                {
                    MyLogger.Error($"*** Failed to get accountid *** mail id = {mail_id}, mail_seq_num = {mail_seq_num}, type = {this.GetType()}");
                    return false;
                }

                // extract card information.

                if (!parse_mail(mail, mail_seq_num, tmp_report))
                {
                    MyLogger.Error($"*** Failed to extract card info *** mail id = {mail_id}, mail_seq_num = {mail_seq_num}, type = {this.GetType()}");
                    return false;
                }

                KReportBase next;

                // Revise order status.

                next = tmp_report;
                while (next != null)
                {
                    next.m_order_status = get_order_status(next);
                    next = next.next_report;
                }

                /*
                 * [DEL 2019-12-01] Some OP email (OP-17 Office Depot) may have status as cancellation.
                 * We will add multiple reports to report_mail DB as one order.
                 * Even the multiple orders have same order, same mail type, but their status may be different, so we must record&report them separately.
                 * 
                // Check if the report has been already handled.
                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_UPDATE_DB))
                {
                    bool status_changed;
                    if (Program.g_db.already_exist_report(tmp_report, out status_changed))
                    {
                        if (!status_changed)
                        {
                            Program.g_db.set_mail_checked_flag(mail_id, ConstEnv.MAIL_PARSING_SUCCEED);

                            MyLogger.Info($"*** Already Report Handled *** mail id = {mail_id}, order_id = {tmp_report.m_order_id}, mail_type = {tmp_report.m_mail_type}");
                            return true;
                        }
                    }
                }
                * -> DEL
                */

                // Convert mail to PDF if needs.

                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_CREATE_PDF_IN_PARSING) && Program.g_user.is_report_mode_crm(user.report_mode))
                {
                    if (this.GetType() == typeof(KMailBaseOP))
                    {
                        bool has_tax = false;
                        string order_id = "";
                        next = tmp_report;
                        while (next != null)
                        {
                            if (next.m_tax > 0)
                            {
                                has_tax = true;
                            }
                            order_id += (order_id == "") ? next.m_order_id : "," + next.m_order_id;
                            next = next.next_report;
                        }
                        if (has_tax)
                        {
                            tmp_report.m_pdf_file = XMail2Pdf.eml_to_pdf(eml_file, order_id, tmp_report.m_retailer);
                        }
                    }
                }

                int parse_state = ConstEnv.MAIL_PARSING_SUCCEED;

                // Scrap web sites if needs.

                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_WEB_SCRAP))
                {
                    next = tmp_report;
                    while (next != null)
                    {
                        KReportBase report = next;
                        next = report.next_report;

                        if (report.m_scrap_params != null)
                        {
                            KWebScrapper.scrap(report);

                            foreach (ZScrapParam param in report.m_scrap_params)
                            {
                                if (param.status != ConstEnv.SCRAP_SUCCESS /*&& param.status != ConstEnv.SCRAP_UNSUPPORTED && param.status != ConstEnv.SCRAP_OTHER*/)
                                {
                                    MyLogger.Error($"Web scrapping is failed. mail id = {mail_id}, mail_type = {report.m_mail_type}, link = {param.link}");
                                    parse_state = ConstEnv.MAIL_SCRAP_PARTIAL_FAILED;
                                }
                            }
                        }
                    }
                }

                // Set flag that mail has been parsed.
                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_UPDATE_DB))
                    Program.g_db.set_mail_checked_flag(mail_id, parse_state);

                // save to DB.

                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_UPDATE_DB))
                {
                    next = tmp_report;
                    while (next != null)
                    {
                        KReportBase report = next;

                        int report_id = report.insert_report_to_db(mail_id);
                        if (report_id == -1)
                        {
                            MyLogger.Error($"*** Failed insert to report info to DB *** mail id = {mail_id}, order = {mail_seq_num}, type = {this.GetType()}");
                            return false;
                        }
                        report.m_report_id = report_id;

                        if (report.is_order_report())
                            Program.g_db.insert_order_status_to_db(report);

                        next = report.next_report;
                    }
                }

                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_REPORT))
                {
                    if (Program.g_user.is_report_mode_gs(user.report_mode))
                    {
                        next = tmp_report;
                        while (next != null)
                        {
                            KReportBase report = next;
                            XStatHelper.update_report(report);
                            next = report.next_report;
                        }
                    }

                    // Send web post. Even it has multi orders, we will report as one because it comes from one mail.
                    if (Program.g_user.is_report_mode_crm(user.report_mode))
                        tmp_report.report();
                }


                // Delete local eml file here.
                if (ConstEnv.check_handle_flag(ConstEnv.APP_WORK_MODE_DELETE_LOCAL_MAIL))
                {
                    // To Do.
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return false;
            }

            return true;
        }
        protected virtual bool merge_card_to_db(int card_id)
        {
            try
            {
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return false;
            }
            return true;
        }

        #endregion DB handler virtual functions.

        #region protected functions.

        protected string get_order_status(KReportBase report)
        {
            string status = report.m_order_status;

            if (report.GetType() == typeof(KReportOP))
            {
                if (status == "" || status == ConstEnv.REPORT_ORDER_STATUS_PURCHAESD)
                {
                    int total_num = report.m_product_items.Count;
                    int item_canceled_num = report.m_product_items.Count(s => s.status == ConstEnv.REPORT_ORDER_STATUS_PARTIAL_CANCELED);

                    if (item_canceled_num == 0)
                    {
                        status = ConstEnv.REPORT_ORDER_STATUS_PURCHAESD;
                    }
                    else if (total_num == item_canceled_num)
                    {
                        status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
                    }
                    else
                    {
                        if (item_canceled_num > 0)
                            status = ConstEnv.REPORT_ORDER_STATUS_PARTIAL_CANCELED;
                    }
                }
            }
            if (report.GetType() == typeof(KReportSC))
            {
                int total_num = report.m_product_items.Count;
                int item_canceled_num = report.m_product_items.Count(s => s.status == ConstEnv.REPORT_ORDER_STATUS_PARTIAL_CANCELED);
                int item_not_shipped_num = report.m_product_items.Count(s => s.status == ConstEnv.REPORT_ORDER_STATUS_PARTIAL_NOT_SHIPPED);

                if (status == "" || status == ConstEnv.REPORT_ORDER_STATUS_SHIPPED)
                {
                    if (item_canceled_num == 0 && item_not_shipped_num == 0)
                    {
                        status = ConstEnv.REPORT_ORDER_STATUS_SHIPPED;
                    }
                    else if (total_num == item_canceled_num)
                    {
                        status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
                    }
                    else if (total_num == item_not_shipped_num)
                    {
                        status = ConstEnv.REPORT_ORDER_STATUS_NOT_SHIPPED;
                    }
                    else
                    {
                        if (item_canceled_num > 0)
                            status = ConstEnv.REPORT_ORDER_STATUS_PARTIAL_CANCELED;
                        if (item_not_shipped_num > 0)
                            status += (status == "") ? ConstEnv.REPORT_ORDER_STATUS_PARTIAL_NOT_SHIPPED : ", " + ConstEnv.REPORT_ORDER_STATUS_PARTIAL_NOT_SHIPPED;
                    }
                }
            }
            if (report.GetType() == typeof(KReportCC))
            {
                status = ConstEnv.REPORT_ORDER_STATUS_CANCELED;
            }

            return status;
        }

        #endregion protected functions.
    }
}
