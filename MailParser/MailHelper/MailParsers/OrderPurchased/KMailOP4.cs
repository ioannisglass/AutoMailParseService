using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Utils;

namespace MailHelper
{
    partial class KMailBaseOP : KMailBaseParser
    {
        private void parse_mail_op_4(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_4_for_htmltext(mail, report);
            }
            else
            {
                parse_mail_op_4_for_bodytext(mail, report);
            }
            if (report.m_order_id == "")
            {
                get_op4_order_id_from_html(mail, report);
            }
        }
        private void get_op4_order_id_from_html(MimeMessage mail, KReportOP report)
        {
            string htmltext = XMailHelper.get_htmltext(mail);
            string temp = htmltext.Replace("%3D", "=");
            while (temp.IndexOf("orderId=") != -1)
            {
                if (temp.IndexOf("%", temp.IndexOf("orderId=")) == -1)
                    break;
                string temp1 = temp.Substring(0, temp.IndexOf("%", temp.IndexOf("orderId=")));
                if (temp1.LastIndexOf("http://") == -1)
                    break;
                temp1 = temp1.Substring(temp1.LastIndexOf("http://"));
                temp1 = temp1.Substring(temp1.IndexOf("orderId=") + "orderId=".Length);
                report.set_order_id(temp1);

                temp = temp.Substring(temp.IndexOf("orderId=") + "orderId=".Length);
            }

            if (report.m_order_id != "")
            {
                MyLogger.Info($"... OP-4 order id = {report.m_order_id}");
            }
        }
        private void parse_mail_op_4_for_bodytext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_4;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_AMAZON;

            MyLogger.Info($"... OP-4 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-4 m_op_retailer = {report.m_retailer}");

            int report_count = 1;
            string[] html_lines = XMailHelper.get_htmltext(mail).Replace("\r", "").Split('\n');
            foreach (string html_line in html_lines)
            {
                string key = "Your purchase has been divided into";
                if (html_line.Contains(key))
                {
                    string tempp = html_line.Substring(html_line.IndexOf(key) + key.Length);
                    tempp = tempp.Substring(0, tempp.IndexOf("orders."));
                    tempp = XMailHelper.exclude_tags(tempp);
                    report_count = Str_Utils.string_to_int(tempp);
                    break;
                }
            }

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');

            //List<List<string>> html_parts = new List<List<string>>();
            List<List<string>> body_parts = new List<List<string>>();

            /*for(int i = 0; i < card_count; i ++)
            {
                html_parts.Add(new List<string>());
                body_parts.Add(new List<string>());
            }*/
            bool is_order_details = false;
            if (report_count == 1)
            {
                //html_parts.Add(html_lines.ToList());
                body_parts.Add(lines.ToList());
            }
            else
            {
                int i = 0;
                for (i = 0; i < report_count; i++)
                    body_parts.Add(new List<string>());

                i = 0;
                bool is_finished = false;
                for (int j = 0; j < lines.Length; j ++)
                {
                    string line = lines[j].Trim();

                    if (!is_order_details)
                        if (line.IndexOf("Order Details") != -1 ||
                                    line.IndexOf("Details") != -1 ||
                                    line.IndexOf("Order 1 of") != -1)
                            if (lines[j + 1].IndexOf("Order #") != -1)
                                is_order_details = true;

                    if (line.StartsWith("Order #") && is_order_details)
                    {
                        for (int k = j; k < lines.Length;)
                        {
                            body_parts[i].Add(lines[k]);
                            k++;
                            if (lines[k].StartsWith("Order #"))
                            {
                                i++;
                                j = k - 1;
                                break;
                            }
                            if (lines[k].Trim().StartsWith("To learn more about ordering,") ||
                                lines[k].Trim().StartsWith("We hope to see you again soon.") ||
                                lines[k].Trim().StartsWith("If you want more information or assistance,"))
                            {
                                is_finished = true;
                                break;
                            }
                        }
                    }
                    if (is_finished)
                        break;
                }
            }

            List<KReportBase> report_list = new List<KReportBase>();
            report_list.Add(report);
            for (int i = 1; i < report_count; i++)
                report_list.Add(new KReportOP());

            for (int idx_report = 0; idx_report < report_count; idx_report ++)
            {
                float total_before_tax = 0;
                for (int i = 0; i < body_parts[idx_report].Count; i++)
                {
                    string line = body_parts[idx_report][i].Trim();

                    if (line.StartsWith("Order #"))
                    {
                        string temp = line.Substring("Order #".Length).Trim();
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<"));

                        report_list[idx_report].set_order_id(temp);

                        MyLogger.Info($"... OP-4 order id = {temp}");
                        continue;
                    }
                    if (line.StartsWith("Placed on "))
                    {
                        string temp = line.Substring("Placed on ".Length).Trim();
                        DateTime date = DateTime.Parse(temp);
                        report_list[idx_report].m_op_purchase_date = date;
                        MyLogger.Info($"... OP-4 order date = {date}");
                        continue;
                    }
                    if (line == "Sold by Amazon.com Services, Inc")
                    {
                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;

                        string temp = body_parts[idx_report][i + 2].Trim();
                        price = Str_Utils.string_to_currency(temp);

                        temp = body_parts[idx_report][i - 2].Trim();
                        while (temp.IndexOf("[") != -1)
                            temp = temp.Substring(temp.IndexOf("]") + 1).Trim();
                        while (temp[0] == '<')
                            temp = temp.Substring(temp.IndexOf(">") + 1).Trim();
                        if (temp.IndexOf("<") != -1)
                            temp = temp.Substring(0, temp.IndexOf("<")).Trim();

                        if (temp.IndexOf(" x ") != -1)
                        {
                            string qty_part = temp.Substring(0, temp.IndexOf(" x "));
                            qty = Str_Utils.string_to_int(qty_part);

                            title = temp.Substring(temp.IndexOf(" x ") + " x ".Length);
                        }

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report_list[idx_report].m_product_items.Add(product);

                            MyLogger.Info($"... OP-4 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }

                        continue;
                    }
                    else
                    {
                        string key = "";
                        if (line.IndexOf("Sold by") != -1)
                            key = "Sold by";
                        else if (line.IndexOf("Provided by") != -1)
                            key = "Provided by";

                        if (key != "")
                        {
                            string title = "";
                            string sku = "";
                            int qty = 0;
                            float price = 0;

                            string temp = "";

                            if (body_parts[idx_report][i - 1].Trim().StartsWith("$"))
                                temp = body_parts[idx_report][i - 1].Trim();
                            else if (body_parts[idx_report][i + 2].Trim().StartsWith("$"))
                                temp = body_parts[idx_report][i + 2].Trim();
                            price = Str_Utils.string_to_currency(temp);

                            if (body_parts[idx_report][i - 3].Replace(" ", "") == string.Empty)
                            {
                                title = body_parts[idx_report][i - 2].Trim();
                            }
                            else if (body_parts[idx_report][i - 4].Replace(" ", "") == string.Empty)
                            {
                                title = body_parts[idx_report][i - 3].Trim() + body_parts[idx_report][i - 2].Trim();
                            }

                            bool excaped = false;
                            while (title != "")
                            {
                                foreach (string html_line in html_lines)
                                {
                                    string ttemp = "";

                                    if (!html_line.Contains("\"name\""))
                                        continue;

                                    if (!HttpUtility.HtmlDecode(html_line).Contains(title))
                                    {
                                        string similar_title = XMailHelper.get_text(HttpUtility.HtmlDecode(html_line), "a");
                                        if (similar_title != title)
                                            continue;
                                    }

                                    qty = 1;
                                    if (html_line.IndexOf("Qty : ") != -1)
                                    {
                                        ttemp = html_line.Substring(html_line.IndexOf("Qty : ") + "Qty : ".Length);
                                        ttemp = ttemp.Substring(0, ttemp.IndexOf("<"));
                                        qty = Str_Utils.string_to_int(ttemp);
                                    }
                                    break;
                                }
                                if (qty > 0 || excaped)
                                    break;
                                foreach (string html_line in html_lines)
                                {
                                    int pos;
                                    if (html_line.Contains("<td class=\"name\""))
                                    {
                                        title = XMailHelper.find_html_part(HttpUtility.HtmlDecode(html_line), "font", out pos);
                                        if (title == string.Empty)
                                            title = XMailHelper.get_text(HttpUtility.HtmlDecode(html_line), "a");
                                        break;
                                    }
                                }
                                excaped = true;
                            }

                            if (qty > 0)
                            {
                                ZProduct product = new ZProduct();
                                product.price = price;
                                product.sku = sku;
                                product.title = title;
                                product.qty = qty;
                                report_list[idx_report].m_product_items.Add(product);

                                MyLogger.Info($"... OP-4 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                            }

                            continue;
                        }
                    }

                    if (line.StartsWith("Estimated Tax:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string temp = line.Substring("Estimated Tax:".Length).Trim();
                        float tax = Str_Utils.string_to_currency(temp);
                        report_list[idx_report].m_tax = tax;
                        MyLogger.Info($"... OP-4 tax = {tax}");
                        continue;
                    }
                    if (line.StartsWith("Total Before Tax:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string temp = line.Substring("Total Before Tax:".Length).Trim();
                        total_before_tax = Str_Utils.string_to_currency(temp);
                        MyLogger.Info($"... OP-4 Total Before Tax: = {total_before_tax}");
                        continue;
                    }
                    if (line.IndexOf("Card:") != -1)
                    {
                        string payment_type = "";
                        string temp = "";
                        payment_type = line.Substring(0, line.IndexOf(":")).Trim();
                        temp = line.Substring(line.IndexOf(":") + 1).Trim();

                        if (temp != "" && temp[0] == '-' && temp[1] == '-')
                            temp = temp.Substring(1);
                        if (temp != "" && temp[0] == '-')
                            temp = temp.Substring(1);

                        float price = Str_Utils.string_to_currency(temp);

                        ZPaymentCard c = new ZPaymentCard(payment_type, "", price);
                        report_list[idx_report].add_payment_card_info(c);

                        MyLogger.Info($"... OP-4 payment_type = {payment_type}, last_digit = \"\", price = {price}");
                    }
                    if (line.StartsWith("Your order will be sent to:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string temp = body_parts[idx_report][++i].Trim();
                        string full_address = "";
                        string state_address = "";
                        int k = 0;
                        while (k++ < 4) // address lines will be not over 4 lines.
                        {
                            full_address += " " + temp;
                            temp = body_parts[idx_report][++i].Trim();

                            state_address = XMailHelper.get_address_state_name(full_address);
                            if (state_address != "")
                                break;
                        }
                        full_address = full_address.Trim();
                        state_address = state_address.Trim();
                        if (state_address != "")
                        {
                            report_list[idx_report].set_address(full_address, state_address);
                            MyLogger.Info($"... OP-4 full_address = {full_address}, state_address = {state_address}");
                        }
                        continue;
                    }

                }
                float total = total_before_tax + report_list[idx_report].m_tax;
                report_list[idx_report].set_total(total);
                MyLogger.Info($"... OP-4 total = {total}");
            }

            //card = (KCardOrderPurchased)card_list[0];
            KReportOP prev = report;
            if (report_count != 1)
            {
                for (int idx_count = 1; idx_count < report_count; idx_count++)
                {
                    prev.next_report = report_list[idx_count];
                    
                    prev.next_report.m_mail_type = KReportBase.MailType.OP_4;
                    prev.next_report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
                    prev.next_report.m_retailer = ConstEnv.RETAILER_AMAZON;
                    prev.next_report.m_mail_id = prev.m_mail_id;
                    prev.next_report.m_mail_account_id = prev.m_mail_account_id;
                    prev.next_report.m_mail_sent_date = prev.m_mail_sent_date;

                    prev = (KReportOP)prev.next_report;
                }
            }                
        }

        private void parse_mail_op_4_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_4;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_AMAZON;

            MyLogger.Info($"... OP-4 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-4 m_op_retailer = {report.m_retailer}");

            int report_count = 1;
            string[] html_lines = XMailHelper.get_htmltext(mail).Replace("\r", "").Split('\n');
            foreach(string html_line in html_lines)
            {
                string key = "Your purchase has been divided into";
                if (html_line.Contains(key))
                {
                    string tempp = html_line.Substring(html_line.IndexOf(key) + key.Length);
                    tempp = tempp.Substring(0, tempp.IndexOf("orders."));
                    tempp = XMailHelper.exclude_tags(tempp);
                    report_count = Str_Utils.string_to_int(tempp);
                    break;
                }
            }

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');

            List<List<string>> body_parts2 = new List<List<string>>();
            bool is_order_details = false;

            if (report_count == 1)
            {
                body_parts2.Add(new List<string>());
                body_parts2.Add(lines.ToList());
                for(int i = 0; i < lines.Length; i ++)
                {
                    if (lines[i].Trim().StartsWith("To learn more about ordering,") ||
                                lines[i].Trim().StartsWith("We hope to see you again soon.") ||
                                lines[i].Trim().StartsWith("If you want more information or assistance,"))
                        break;

                    if (lines[i].Trim() == "Order" &&
                           lines[i + 1].Trim().StartsWith("#"))
                    {
                        body_parts2[0].Add(lines[i].Trim() + " " + lines[i + 1].Trim());
                        i++;
                        continue;
                    }
                    body_parts2[0].Add(lines[i]);
                }
            }
            else
            {
                int i;
                for (i = 0; i < report_count; i++)
                    body_parts2.Add(new List<string>());

                i = 0;
                bool is_finished = false;
                for (int j = 0; j < lines.Length; j++)
                {
                    string line = lines[j].Trim();

                    if (!is_order_details)
                        if (line.IndexOf("Order Details") != -1 ||
                                    line.IndexOf("Details") != -1 ||
                                    line.IndexOf("Order 1 of") != -1)
                            if (lines[j + 1].IndexOf("Order #") != -1 ||
                                lines[j + 1].StartsWith("Order") &&
                                lines[j + 2].StartsWith("#"))
                                is_order_details = true;
                    int k;
                    if (!is_order_details)
                        continue;
                    if (line.StartsWith("Order #"))
                    {
                        body_parts2[i].Add(line);
                        k = j + 1;
                    }
                    else if (line == "Order" &&
                                lines[j + 1].Trim().StartsWith("#"))
                    {
                        body_parts2[i].Add(line + " " + lines[j + 1].Trim());
                        k = j + 2;
                    }
                    else
                        continue;

                    for (; k < lines.Length;)
                    {
                        body_parts2[i].Add(lines[k]);
                        k++;
                        if (lines[k].Trim().StartsWith("Order #") ||
                                lines[k].Trim().StartsWith("Order") &&
                                lines[k + 1].Trim().StartsWith("#"))
                        {
                            i++;
                            j = k - 1;
                            break;
                        }
                        if (lines[k].Trim().StartsWith("To learn more about ordering,") ||
                                lines[k].Trim().StartsWith("We hope to see you again soon.")||
                                lines[k].Trim().StartsWith("If you want more information or assistance,"))
                        {
                            is_finished = true;
                            break;
                        }
                    }
                    if (is_finished)
                        break;
                }
            }

            List<KReportBase> report_list = new List<KReportBase>();
            report_list.Add(report);
            for (int i = 1; i < report_count; i++)
                report_list.Add(new KReportOP());

            List<string> item_title_list = get_op4_item_title_list(html_lines);
            int item_total = 0;

            for (int idx_report = 0; idx_report < report_count; idx_report++)
            {
                int item_count = 0;
                foreach (string line in body_parts2[idx_report])
                {
                    if (line.IndexOf("Sold by") != -1)
                        item_count++;
                }

                if (item_title_list.Count != 0)
                {
                    string sum = "";
                    foreach (string line in body_parts2[idx_report])
                        sum = sum + line.Trim() + " ";
                    sum = sum.Trim();

                    for(int idx_item = item_total; idx_item < item_title_list.Count; idx_item ++)
                    {
                        string item = item_title_list[idx_item];

                        string title = item;
                        int qty = 1;
                        string sku = "";
                        float price = 0;

                        string temp;
                        sum = sum.Substring(sum.IndexOf(item) + item.Length);

                        string key = "";
                        if (sum.IndexOf("Sold by") != -1)
                            key = "Sold by";
                        else if (sum.IndexOf("Provided by") != -1)
                            key = "Provided by";

                        temp = sum.Substring(0, sum.IndexOf(key));
                        sum = sum.Substring(sum.IndexOf(key) + key.Length);

                        if (temp.IndexOf("Qty : ") != -1)
                        {
                            string sub_temp = temp.Substring(temp.IndexOf("Qty : ") + "Qty : ".Length);
                            qty = Str_Utils.string_to_int(sub_temp.Substring(0, sub_temp.IndexOf(" ")));
                        }
                        if (temp.IndexOf("$") != -1)
                        {
                            string sub_temp = temp.Substring(temp.IndexOf("$"));
                            price = Str_Utils.string_to_currency(sub_temp.Substring(0, sub_temp.IndexOf(" ")));
                        }

                        if (qty > 0)
                        {
                            ZProduct product = new ZProduct();
                            product.price = price;
                            product.sku = sku;
                            product.title = title;
                            product.qty = qty;
                            report_list[idx_report].m_product_items.Add(product);

                            MyLogger.Info($"... OP-4 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                        }
                        item_total++;
                    }
                    //item_total += item_count;
                }

                float total_before_tax = 0;
                float tax = 0;
                float shpping = 0;
                float sub_total = 0;

                for (int i = 0; i < body_parts2[idx_report].Count; i++)
                {
                    string line = body_parts2[idx_report][i].Trim();

                    if (line.IndexOf("Order #") != -1)
                    {
                        string temp = line.Substring(line.IndexOf("Order #") + "Order #".Length).Trim();

                        if (report_list[idx_report].m_order_id == string.Empty)
                            report_list[idx_report].set_order_id(temp);
                        else if (report_list[idx_report].m_order_id != temp)
                            report_list[idx_report].set_order_id(temp);

                        MyLogger.Info($"... OP-4 order id = {temp}");
                    }
                    if (line.ToUpper() == "SHIPPING & HANDLING" || line.ToUpper() == "SHIPPING & HANDLING:")
                    {
                        string temp = body_parts2[idx_report][i + 1].Trim();
                        shpping = Str_Utils.string_to_currency(temp);
                        MyLogger.Info($"... OP-4 Shipping & Handling = {shpping}");
                        continue;
                    }
                    if (((line.ToUpper() == "ITEMS" || line.ToUpper() == "ITEM") && lines[i + 2].Trim().ToUpper() == "SHIPPING & HANDLING") || line.ToUpper() == "ITEM SUBTOTAL:")
                    {
                        string temp = body_parts2[idx_report][i + 1].Trim();
                        sub_total = Str_Utils.string_to_currency(temp);
                        MyLogger.Info($"... OP-4 Sub Total = {sub_total}");
                        continue;
                    }
                    if (line.StartsWith("Estimated Tax"))
                    {
                        string temp = "";

                        if (line == "Estimated Tax:" || line == "Estimated Tax")
                            temp = body_parts2[idx_report][i + 1];
                        else
                        {
                            if (line.Contains("Estimated Tax:"))
                                temp = line.Substring("Estimated Tax:".Length);
                            else if(line.Contains("Estimated Tax"))
                                temp = line.Substring("Estimated Tax".Length);
                        }

                        tax = Str_Utils.string_to_currency(temp);
                        MyLogger.Info($"... OP-4 Estimated Tax = {tax}");
                        continue;
                    }
                    if (line.StartsWith("Total Before Tax"))
                    {
                        string temp = "";

                        if (line == "Total Before Tax:" || line == "Total Before Tax")
                            temp = body_parts2[idx_report][i + 1];
                        else
                        {
                            if (line.Contains("Total Before Tax:"))
                                temp = line.Substring("Total Before Tax:".Length);
                            else if (line.Contains("Total Before Tax"))
                                temp = line.Substring("Total Before Tax".Length);
                        }

                        total_before_tax = Str_Utils.string_to_currency(temp);
                        MyLogger.Info($"... OP-4 total before tax = {total_before_tax}");
                        continue;
                    }
                    if (line.EndsWith("/Card:"))
                    {
                        string payment_type;
                        string temp;
                        temp = body_parts2[idx_report][i + 1].Trim();
                        payment_type = line.Substring(0, line.IndexOf(":")).Trim();

                        if (!temp.StartsWith("$") && !temp.StartsWith("-$") && !temp.StartsWith("--$"))
                            continue;

                        if (temp != "" && temp[0] == '-' && temp[1] == '-')
                            temp = temp.Substring(1);
                        if (temp != "" && temp[0] == '-')
                            temp = temp.Substring(1);

                        float price = Str_Utils.string_to_currency(temp);

                        ZPaymentCard c = new ZPaymentCard(payment_type, "", price);
                        report_list[idx_report].add_payment_card_info(c);

                        MyLogger.Info($"... OP-4 payment_type = {payment_type}, last_digit = \"\", price = {price}");
                        continue;
                    }

                    if (line.StartsWith("Purchase Summary"))
                    {
                        string temp = body_parts2[idx_report][i + 1].Trim();
                        DateTime date = DateTime.Parse(temp);
                        report_list[idx_report].m_op_purchase_date = date;
                        MyLogger.Info($"... OP-4 order date = {date}");

                        i += 2;
                        temp = body_parts2[idx_report][i].Trim();
                        if (temp.StartsWith("Est. Delivery:", StringComparison.CurrentCultureIgnoreCase))
                            temp = body_parts2[idx_report][++i].Trim();

                        string full_address = "";
                        string state_address = "";
                        int k = 0;
                        while (k++ < 4) // address lines will be not over 4 lines.
                        {
                            full_address += " " + temp;
                            temp = body_parts2[idx_report][++i].Trim();

                            state_address = XMailHelper.get_address_state_name(full_address);
                            if (state_address != "")
                                break;
                        }
                        full_address = full_address.Trim();
                        state_address = state_address.Trim();
                        if (state_address != "")
                        {
                            report_list[idx_report].set_address(full_address, state_address);
                            MyLogger.Info($"... OP-4 full_address = {full_address}, state_address = {state_address}");
                        }
                        continue;
                    }
                    if (line.StartsWith("Your order will be sent to:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string temp = body_parts2[idx_report][++i].Trim();
                        string full_address = "";
                        string state_address = "";
                        int k = 0;
                        while (k++ < 4) // address lines will be not over 4 lines.
                        {
                            full_address += " " + temp;
                            temp = body_parts2[idx_report][++i].Trim();

                            state_address = XMailHelper.get_address_state_name(full_address);
                            if (state_address != "")
                                break;
                        }
                        full_address = full_address.Trim();
                        state_address = state_address.Trim();
                        if (state_address != "")
                        {
                            report_list[idx_report].set_address(full_address, state_address);
                            MyLogger.Info($"... OP-4 full_address = {full_address}, state_address = {state_address}");
                        }
                        continue;
                    }
                    if (line.StartsWith("Email delivery:"))
                    {
                        string temp = body_parts2[idx_report][i + 1].Trim();
                        DateTime date = DateTime.Parse(temp);
                        report_list[idx_report].m_op_purchase_date = date;
                        MyLogger.Info($"... OP-4 order date = {date}");
                        continue;
                    }
                    if (line.StartsWith("Placed on "))
                    {
                        string temp = line.Substring("Placed on ".Length).Trim();
                        DateTime date = DateTime.Parse(temp);
                        report_list[idx_report].m_op_purchase_date = date;
                        MyLogger.Info($"... OP-4 order date = {date}");
                        continue;
                    }
                    if (line.StartsWith("Guaranteed delivery date:"))
                    {
                        string temp;
                        if (line == "Guaranteed delivery date:")
                            temp = body_parts2[idx_report][i + 1];
                        else
                            temp = line.Substring("Guaranteed delivery date:".Length);
                        DateTime date = DateTime.Parse(temp);
                        report_list[idx_report].m_op_purchase_date = date;
                        MyLogger.Info($"... OP-4 order date = {date}");
                        continue;
                    }
                }
                {
                    float total = (total_before_tax != 0) ? total_before_tax + tax : sub_total + tax;
                    report_list[idx_report].set_total(total);
                    MyLogger.Info($"... OP-4 total = {total}");

                    report_list[idx_report].m_tax = tax;
                    MyLogger.Info($"... OP-4 tax = {tax}");
                }
            }

            KReportOP prev = report;
            if (report_count != 1)
            {
                for (int idx_count = 1; idx_count < report_count; idx_count++)
                {
                    prev.next_report = report_list[idx_count];

                    prev.next_report.m_mail_type = KReportBase.MailType.OP_4;
                    prev.next_report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
                    prev.next_report.m_retailer = ConstEnv.RETAILER_AMAZON;

                    prev = (KReportOP)prev.next_report;
                }
            }
        }

        public List<string> get_op4_item_title_list(string[] htmllines)
        {
            List<string> ret = new List<string>();
            string temp = "";
            bool is_one_finished = false;
            bool is_one_started = false;

            foreach(string line in htmllines)
            {
                if (line.Contains("<td") &&
                        line.Contains("class=\"name\""))
                {
                    is_one_started = true;
                }
                if (line.Contains("</td>") && is_one_started)
                {
                    temp += line;
                    is_one_finished = true;
                    is_one_started = false;
                }

                if (is_one_started)
                    temp += line;

                if (is_one_finished)
                {
                    int pos;
                    temp = XMailHelper.find_html_part(temp, "font", out pos);
                    temp = XMailHelper.exclude_tags(temp);
                    ret.Add(XMailHelper.html2text(temp));
                    temp = "";
                    is_one_finished = false;
                }
            }
            return ret;
        }
    }
}
