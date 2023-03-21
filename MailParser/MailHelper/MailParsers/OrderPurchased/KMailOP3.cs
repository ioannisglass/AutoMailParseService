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
        private void parse_mail_op_3(MimeMessage mail, KReportOP report)
        {
            if (!XMailHelper.is_bodytext_existed(mail))
            {
                parse_mail_op_3_for_htmltext(mail, report);
                return;
            }
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_3;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = ConstEnv.RETAILER_BESTBUY;

            MyLogger.Info($"... OP-3 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-3 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line == "Order #")
                {
                    string temp = lines[++i].Trim();
                    if (temp.IndexOf("-->") != -1)
                        temp = temp.Substring(0, temp.IndexOf("-->")).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-3 order id = {temp}");
                }
                if (line.StartsWith("SKU:"))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = lines[i - 1].Trim();
                    if (temp.StartsWith("Model:"))
                    {
                        // To Do.
                    }
                    temp = lines[i - 2].Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    title = temp;

                    temp = line.Substring("SKU:".Length).Trim();
                    sku = temp;

                    temp = lines[++i].Trim(); // Qty     Price
                    temp = lines[++i].Trim();
                    string qty_part = temp.Substring(0, temp.IndexOf(" "));
                    qty = Str_Utils.string_to_int(qty_part);
                    temp = temp.Substring(temp.IndexOf(" ") + 1).Trim();
                    price = Str_Utils.string_to_currency(temp);

                    if (qty > 0)
                    {
                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... OP-3 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    }
                    continue;
                }
                if (line.StartsWith("Tax:*"))
                {
                    string temp = line.Substring("Tax:*".Length);
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-3 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Order Total:*"))
                {
                    string temp = line.Substring("Order Total:*".Length);
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-3 total = {total}");
                    continue;
                }
            }
        }
        private void parse_mail_op_3_for_htmltext(MimeMessage mail, KReportOP report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.OP_3;

            report.m_receiver = Program.g_user.get_mailaddress_by_account_id(report.m_mail_account_id);
            report.m_retailer = "Best Buy";

            MyLogger.Info($"... OP-3 m_op_receiver = {report.m_receiver}");
            MyLogger.Info($"... OP-3 m_op_retailer = {report.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = "";
                    if (line.ToUpper() == "ORDER #")
                    {
                        temp = lines[++i].Trim();
                        if (temp.IndexOf("ORDER #", StringComparison.CurrentCultureIgnoreCase) != -1)
                            temp = temp.Substring(0, temp.IndexOf("ORDER #", StringComparison.CurrentCultureIgnoreCase)).Trim();
                        if (temp.IndexOf("-->") != -1)
                            temp = temp.Substring(0, temp.IndexOf("-->")).Trim();
                    }
                    else
                    {
                        temp = line;
                        temp = temp.Substring("ORDER #".Length).Trim();
                        if (temp.IndexOf("-->") != -1)
                            temp = temp.Substring(0, temp.IndexOf("-->")).Trim();
                    }
                    report.set_order_id(temp);
                    MyLogger.Info($"... OP-3 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Model:", StringComparison.CurrentCultureIgnoreCase) && lines[i + 1].Trim().StartsWith("SKU:", StringComparison.CurrentCultureIgnoreCase) && lines[i + 2].Trim().ToUpper() == "QTY" && lines[i + 3].Trim().ToUpper() == "PRICE")
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = lines[i - 1].Trim();
                    title = temp;

                    temp = lines[i + 1].Trim();
                    temp = temp.Substring("SKU:".Length).Trim();
                    sku = temp;

                    temp = lines[i + 4].Trim();
                    qty = Str_Utils.string_to_int(temp);
                    temp = lines[i + 5].Trim();
                    price = Str_Utils.string_to_currency(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-3 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    i += 4;
                    continue;
                }
                if (line.ToUpper() == "QTY" && lines[i + 1].Trim().ToUpper() == "PRICE")
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp = lines[i - 1].Trim();
                    if (temp.IndexOf("SKU:") != -1)
                    {
                        title = temp.Substring(0, temp.IndexOf("SKU:")).Trim();
                        sku = temp.Substring(temp.IndexOf("SKU:") + "SKU:".Length).Trim();
                    }
                    else
                    {
                        title = temp;
                    }

                    temp = lines[i + 2].Trim();
                    qty = Str_Utils.string_to_int(temp);
                    temp = lines[i + 3].Trim();
                    price = Str_Utils.string_to_currency(temp);

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... OP-3 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                    i += 3;
                    continue;
                }
                if (line.StartsWith("Tax:*", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line == "TAX:*")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Tax:*".Length);
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-3 tax = {tax}");
                    continue;
                }
                if (line.ToUpper() == "TAX:")
                {
                    string temp = lines[++i].Trim();
                    float tax = Str_Utils.string_to_currency(temp);
                    report.m_tax = tax;
                    MyLogger.Info($"... OP-3 tax = {tax}");
                    continue;
                }
                if (line.StartsWith("Order Total:*", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER TOTAL:*")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Total:*".Length);
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-3 total = {total}");
                    continue;
                }
                if (line.ToUpper() == "ORDER TOTAL:")
                {
                    string temp = lines[++i].Trim();
                    float total = Str_Utils.string_to_currency(temp);
                    report.set_total(total);
                    MyLogger.Info($"... OP-3 total = {total}");
                    continue;
                }

                if (line.ToUpper() == "GET IT BY:")
                {
                    /**
                     *
                     *          GET IT BY:
                     *          WED 05/17
                     *          Pace Cohen
                     *          49 CHELSEA RD
                     *          JACKSON, NJ 08527
                     *          
                     *          
                     *          GET IT BY:
                     *          TUE 05/10
                     *          Delivered To:
                     *          Yitzchok Yoselovsky
                     *          422 3RD ST
                     *          LAKEWOOD, NJ 08701
                     *          
                     *          GET IT BY:
                     *          SAT 08/26
                     *          Samuel Forest106 FOREST DRLAKEWOOD, NJ 08701
                     * 
                     **/

                    string temp = lines[++i].Trim(); // date
                    temp = lines[++i].Trim();
                    if (temp.ToUpper() == "DELIVERED TO:")
                        temp = lines[++i].Trim();


                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 4) // address lines will be not over 4 lines.
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
                        MyLogger.Info($"... OP-3 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (line.ToUpper() == "EXPECTED TO ARRIVE BETWEEN:")
                {
                    /*
                     *          Expected to arrive between:
                     *          WED 12/23 -
                     *          WED 01/13
                     *          Double Deals
                     *          419 CEDAR BRIDGE AVE
                     *          LAKEWOOD, NJ 08701
                     **/

                    i += 2;
                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 4) // address lines will be not over 4 lines.
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
                        MyLogger.Info($"... OP-3 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
                if (line.ToUpper() == "WE'LL NOTIFY YOU WHEN IT'S READY")
                {
                    /*
                     *          WE'LL NOTIFY YOU WHEN IT'S READY
                     *          Pace Cohen49 CHELSEA RDJACKSON, NJ 08527
                     **/

                    string temp = lines[++i].Trim();
                    string full_address = "";
                    string state_address = "";
                    int k = 0;
                    while (k++ < 4) // address lines will be not over 4 lines.
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
                        MyLogger.Info($"... OP-3 full_address = {full_address}, state_address = {state_address}");
                    }
                    continue;
                }
            }
        }
    }
}
