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
    partial class KMailBaseCC : KMailBaseParser
    {
        private void parse_mail_cc_4(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);
            string sender = XMailHelper.get_sender(mail);

            card.m_mail_type = KReportBase.MailType.CC_4;

            card.m_retailer = ConstEnv.RETAILER_AMAZON;

            MyLogger.Info($"... CC-4 m_cc_retailer = {card.m_retailer}");

            if (sender == "fba-noreply@amazon.com")
            {
                analyze_cc4_for_type1(mail, card);
                return;
            }
            if (sender == "order-update@amazon.com" || sender == "payments-messages@amazon.com" || sender == "payments-update@amazon.com")
            {
                analyze_cc4_for_type2(mail, card);
                return;
            }
            if (sender == "qla@amazon.com")
            {
                analyze_cc4_for_type3(mail, card);
                return;
            }
            if (sender == "seller-notification@amazon.com")
            {
                analyze_cc4_for_type4(mail, card);
                return;
            }
        }
        private void analyze_cc4_for_type1(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            string temp = subject;
            if (temp.IndexOf("(") != -1)
            {
                temp = temp.Substring(temp.IndexOf("(") + 1).Trim();
                if (temp.IndexOf(")") != -1)
                {
                    temp = temp.Substring(0, temp.IndexOf(")")).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-7 order id = {temp}");
                }
            }

            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("Your recent order", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = line.Substring(line.IndexOf("Your recent order", StringComparison.CurrentCultureIgnoreCase) + "Your recent order".Length).Trim();
                    if (temp.IndexOf("(") != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("(") + 1).Trim();
                        if (temp.IndexOf(")") != -1)
                        {
                            temp = temp.Substring(0, temp.IndexOf(")")).Trim();
                            card.set_order_id(temp);
                            MyLogger.Info($"... CC-7 order id = {temp}");
                        }
                    }
                    continue;
                }

                if (line.Replace(" ", "").ToUpper() == "QTYITEM")
                {
                    /**
                     *          --------------------------------------------------------------------------
                     *          Qty    Item
                     *          --------------------------------------------------------------------------
                     *           1      Lenovo Chromebook C330 2-in-1 Convertible Laptop, 11.6-Inch HD...
                     *          
                     *          You can see more information about this order by visiting Seller Central at: 
                     **/

                    i++;

                    while (lines[i].Trim()[0] == '-')
                        i++;

                    while (i < lines.Length)
                    {
                        string title = "";
                        string sku = "";
                        int qty = 1;
                        float price = 0;

                        temp = lines[i].Trim();
                        if (temp == "" || temp.StartsWith("You can see more information", StringComparison.CurrentCultureIgnoreCase))
                            break;
                        if (temp.IndexOf(" ") == -1)
                            break;

                        string qty_part = temp.Substring(0, temp.IndexOf(" "));
                        qty = Str_Utils.string_to_int(qty_part);

                        title = temp.Substring(temp.IndexOf(" ") + 1).Trim();

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        card.m_product_items.Add(product);

                        MyLogger.Info($"... CC-4 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                        i++;
                    }
                    continue;
                }
            }
        }
        private void analyze_cc4_for_type2(MimeMessage mail, KReportCC card)
        {
            /**
             *          Order #114-0509967-1741051  Placed on Wednesday, June 26, 2019
             *          Moen 7185EC Brantford Motionsense Two-Sensor Touchless One-Handle High Arc Pulldown Kitchen Faucet Featuring Reflex, Chrome
             *          Sold by Amazon.com Services, Inc
             *
             *          Order #
             *          115-0257535-8577004
             *          Placed
             *          on Tuesday, November 29, 2016
             *        + Lenovo
             *        - Chromebook N22 11.6" Notebook, IPS Touchscreen, Intel N3060 Dual-Core, 16GB eMMC SSD, 4GB DDR3, 802.11ac, Bluetooth, ChromeOS
             *          Sold by
             *          Woot
             *          Cancel Reason: Customer Canceled
             *        + Lenovo
             *        - Chromebook N22 11.6" Notebook, IPS Touchscreen, Intel N3060 Dual-Core, 16GB eMMC SSD, 4GB DDR3, 802.11ac, Bluetooth, ChromeOS
             *          Sold by
             *          Woot
             *          Cancel Reason: Customer Canceled
             * Ignore > Items Shipping Soon
             *          Order #115-0257535-8577004
             *    qty > 2x
             *          Lenovo Chromebook N22 11.6" Notebook, IPS Touchscreen, Intel N3060 Dual-Core, 16GB eMMC SSD, 4GB DDR3, 802.11ac, Bluetooth, ChromeOS
             *          Sold by
             *          Woot
             *          We hope to see you again soon.
             *
             **/

            string temp;
            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (line.ToUpper() == "ORDER #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("Placed on", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf("Placed on", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    }
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-7 order id = {temp}");
                    continue;
                }

                if (line.IndexOf("Sold by", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = "";
                    if (!line.StartsWith("Sold by", StringComparison.CurrentCultureIgnoreCase))
                        temp = line.Substring(0, line.IndexOf("Sold by", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    int k = i - 1;
                    while (!lines[k].Trim().StartsWith("Cancel Reason", StringComparison.CurrentCultureIgnoreCase) && !lines[k].Trim().StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (lines[k].Trim().IndexOf("Sold by", StringComparison.CurrentCultureIgnoreCase) != -1)
                            break;
                        if (lines[k].Trim().IndexOf("Placed on", StringComparison.CurrentCultureIgnoreCase) != -1)
                            break;
                        if (lines[k].StartsWith("on", StringComparison.CurrentCultureIgnoreCase) && lines[k - 1].Trim().StartsWith("Placed", StringComparison.CurrentCultureIgnoreCase))
                            break;

                        temp = lines[k].Trim() + " " + temp;
                        k--;
                    }

                    string title = temp;
                    string sku = "";
                    int qty = 1;
                    float price = 0;

                    temp = temp.Trim();
                    if (temp.IndexOf("x ") != -1)
                    {
                        string qty_part = temp.Substring(0, temp.IndexOf("x ")).Trim();
                        int qty_tmp;
                        if (int.TryParse(qty_part, out qty_tmp))
                        {
                            qty = qty_tmp;
                            temp = temp.Substring(temp.IndexOf("x ") + 2).Trim();
                        }
                    }
                    title = temp;

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... CC-4 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                    continue;
                }

                if (line.ToUpper() != "/*OTHER ITEMS SHIPPING SOON*/" && line.IndexOf("Items Shipping Soon", StringComparison.CurrentCultureIgnoreCase) != -1)
                    break;
            }
        }
        private void analyze_cc4_for_type3(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            string temp;
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.IndexOf("we have canceled your order ", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = line.Substring(line.IndexOf("we have canceled your order ", StringComparison.CurrentCultureIgnoreCase) + "we have canceled your order ".Length).Trim();
                    if (temp.IndexOf(" ") != -1)
                    {
                        temp = temp.Substring(0, temp.IndexOf(" ")).Trim();
                        card.set_order_id(temp);
                        MyLogger.Info($"... CC-7 order id = {temp}");
                    }
                    continue;
                }
            }
        }
        private void analyze_cc4_for_type4(MimeMessage mail, KReportCC card)
        {
            string temp = XMailHelper.get_bodytext(mail).Replace("\r", "").Replace("\n", " ").Trim();

            if (temp.IndexOf("following items have been canceled:", StringComparison.CurrentCultureIgnoreCase) != -1)
            {
                temp = temp.Substring(temp.IndexOf("following items have been canceled:", StringComparison.CurrentCultureIgnoreCase) + "following items have been canceled:".Length).Trim();
                if (temp.IndexOf("Your items no longer appear", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    temp = temp.Substring(0, temp.IndexOf("Your items no longer appear", StringComparison.CurrentCultureIgnoreCase)).Trim();

                    string title = temp;
                    string sku = "";
                    int qty = 1;
                    float price = 0;

                    if (temp.IndexOf(" - ") != -1)
                    {
                        sku = temp.Substring(0, temp.IndexOf(" - ")).Trim();
                        title = temp.Substring(temp.IndexOf(" - ") + 3).Trim();
                    }
                    else
                    {
                        title = temp;
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... CC-4 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                }
            }
        }
    }
}
