using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace MailHelper
{
    partial class KMailBaseOP : KMailBaseParser
    {
        private void parse_mail_op_5(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_5;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_HOMEDEPOT;

            MyLogger.Info($"... OP-5 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-5 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            if (lines.Length < 10)
            {
                string revised_bodytext = XMailHelper.revise_concated_bodytext(XMailHelper.get_bodytext(mail));
                lines = revised_bodytext.Replace("\r", "").Split('\n');
            }
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order Number", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER NUMBER")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Number".Length).Trim();
                    if (temp.IndexOf("Order Date") != -1)
                        continue;
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-5 order id = {temp}");
                }
                if (line.StartsWith("Order Date", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER DATE")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Date".Length).Trim();
                    DateTime date = DateTime.Parse(temp);
                    report.m_op_purchase_date = date;
                    MyLogger.Info($"... OP-5 order date = {date}");
                    continue;
                }
                if (line.StartsWith("Qty ", StringComparison.CurrentCultureIgnoreCase))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = line;
                    temp = temp.Substring("Qty ".Length).Trim();
                    qty = (int)Str_Utils.string_to_float(temp);

                    int k = i - 1;
                    int k1 = k;
                    while (!lines[k].Trim().StartsWith("Unit Price", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (lines[k].Trim().StartsWith("Item", StringComparison.CurrentCultureIgnoreCase) || lines[k].Trim().StartsWith("Qty ", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        k--;
                    }
                    if (lines[k].Trim().StartsWith("Item", StringComparison.CurrentCultureIgnoreCase) || lines[k].Trim().StartsWith("Qty ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        price = 0;
                        k = k1;
                    }
                    else
                    {
                        temp = lines[k].Trim();
                        if (temp.ToUpper() == "UNIT PRICE")
                            temp = lines[k + 1].Trim();
                        else
                            temp = temp.Substring("Unit Price".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);
                    }

                    k--;
                    k1 = k;
                    while (!lines[k].Trim().StartsWith("Store SKU #", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (lines[k].Trim().StartsWith("Item", StringComparison.CurrentCultureIgnoreCase) || lines[k].Trim().StartsWith("Qty ", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        k--;
                    }
                    if (lines[k].Trim().StartsWith("Item", StringComparison.CurrentCultureIgnoreCase) || lines[k].Trim().StartsWith("Qty ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        k = k1;
                        sku = "";
                    }
                    else
                    {
                        temp = lines[k].Trim();
                        if (temp.ToUpper() == "STORE SKU #")
                            temp = lines[k + 1].Trim();
                        else
                            temp = temp.Substring("Store SKU #".Length).Trim();
                        sku = temp;
                    }

                    k--;
                    k1 = k;
                    temp = "";
                    while (!lines[k].Trim().StartsWith("Item", StringComparison.CurrentCultureIgnoreCase) && !lines[k].Trim().StartsWith("Item Total", StringComparison.CurrentCultureIgnoreCase) && !lines[k].Trim().StartsWith("Qty", StringComparison.CurrentCultureIgnoreCase))
                    {
                        temp = lines[k].Trim() + " " + temp;
                        k--;
                    }
                    title = temp.Trim();

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-5 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    continue;
                }
                if (line.StartsWith("Sales Tax", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "SALES TAX")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Sales Tax".Length).Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-5 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Order Total", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER TOTAL")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Total".Length).Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-5 total = {total}");
                    continue;
                }
                if (line.ToUpper() == "SHIPPING ADDRESS")
                {
                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 10)
                    {
                        full_address += " " + temp;
                        temp = lines[++i].Trim();

                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-5 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (line.ToUpper() == "PICKUP STORE" && lines[i + 1].Trim().ToUpper() != "PICKUP PERSON")
                {
                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 10)
                    {
                        full_address += " " + temp;
                        temp = lines[++i].Trim();

                        state_address = XMailHelper.get_address_state_name(full_address);
                        if (state_address != "")
                            break;
                    }
                    full_address = full_address.Trim();
                    state_address = state_address.Trim();
                    if (state_address != "")
                    {
                        report.set_address(full_address, state_address);
                        MyLogger.Info($"... OP-5 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }
        }
    }
}
