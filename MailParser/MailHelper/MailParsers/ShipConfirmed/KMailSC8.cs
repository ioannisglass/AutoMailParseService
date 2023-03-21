using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace MailHelper
{
    partial class KMailBaseSC : KMailBaseParser
    {
        private void parse_mail_sc_8(MimeMessage mail, KReportSC report)
        {
            string subject = XMailHelper.get_subject(mail);

            report.m_mail_type = KReportBase.MailType.SC_8;

            report.m_retailer = ConstEnv.RETAILER_SAMSCLUB;

            MyLogger.Info($"... SC-8 m_sc_retailer = {report.m_retailer}");

            string htmlbody = XMailHelper.get_htmltext(mail);
            string[] lines = XMailHelper.get_bodytext(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order #".Length).Trim();
                    if (temp.IndexOf("Track your order", StringComparison.CurrentCultureIgnoreCase) != -1)
                        temp = temp.Substring(0, temp.IndexOf("Track your order", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-8 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("Order Number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp = line.Substring("Order Number:".Length).Trim();
                    report.set_order_id(temp);
                    MyLogger.Info($"... SC-8 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("tracking number", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(0, line.IndexOf("tracking number", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    string post_type = get_post_type(temp);
                    report.m_sc_post_type = post_type;

                    if (line.EndsWith("tracking number"))
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring(line.IndexOf("tracking number", StringComparison.CurrentCultureIgnoreCase) + "tracking number".Length).Trim();
                    if (temp.IndexOf("<") != -1)
                        temp = temp.Substring(0, temp.IndexOf("<")).Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);
                    MyLogger.Info($"... SC-8 post_type = {post_type}, tracking = {tracking}");
                    continue;
                }
                if (line.IndexOf("Tracking:", StringComparison.CurrentCultureIgnoreCase) != -1 && line.IndexOf("Shipping:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Tracking:", StringComparison.CurrentCultureIgnoreCase) + "Tracking:".Length).Trim();
                    string tracking = temp;
                    report.set_tracking(tracking);

                    temp = line.Substring(0, line.IndexOf("Tracking:", StringComparison.CurrentCultureIgnoreCase)).Trim();
                    if (temp.IndexOf("Shipping:", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        temp = temp.Substring(temp.IndexOf("Shipping:", StringComparison.CurrentCultureIgnoreCase) + "Shipping:".Length).Trim();
                        string post_type = get_post_type(temp);
                        report.m_sc_post_type = post_type;
                    }
                    MyLogger.Info($"... SC-8 post_type = {report.m_sc_post_type}, tracking = {tracking}");
                    continue;
                }
                if (line == "ITEMQTYPRICE")
                {
                    string temp = lines[++i].Trim();

                    // EpsonWorkForceDS40Scanner5$119.98

                    while (temp != "")
                    {
                        if (temp.IndexOf("Track this shipment", StringComparison.CurrentCultureIgnoreCase) != -1 || temp.IndexOf("$") == -1)
                            break;

                        string title = "";
                        string sku = "";
                        int qty = 0;
                        float price = 0;
                        int pos = temp.IndexOf("$");
                        string temp1 = temp.Substring(pos + 1).Trim();
                        price = Str_Utils.string_to_currency(temp1);

                        temp1 = temp.Substring(0, pos).Trim();
                        int n = temp1.Length;
                        if (n < 2)
                            break;
                        int k;
                        for (k = 0; k < n - 1; k++)
                        {
                            string temp2 = temp1.Substring(0, n - k);
                            if (htmlbody.IndexOf(temp2 + "</td>") != -1)
                                break;
                        }
                        if (k == n - 2)
                            break;
                        title = temp1.Substring(0, n - k);
                        string qty_part = temp1.Substring(n - k);
                        qty = Str_Utils.string_to_int(qty_part);

                        ZProduct product = new ZProduct();
                        product.price = price;
                        product.sku = sku;
                        product.title = title;
                        product.qty = qty;
                        report.m_product_items.Add(product);

                        MyLogger.Info($"... SC-8 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");

                        temp = lines[++i].Trim();
                    }
                }
                if (line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    /**
                     * There are 5 kinds.
                     *      
                     * 1) No shipping, No savings : one item has 1 1ine.
                     * 
                     *      ItemQtyOrig. price
                     *      EpsonExpressionET2750SpecialEditionEcoTankAllinOnePrinterItem # 9800946001$249.87
                     *      EpsonExpressionET2750SpecialEditionEcoTankAllinOnePrinterItem # 9800946001$249.87
                     *      EpsonExpressionET2750SpecialEditionEcoTankAllinOnePrinterItem # 980094600Qty: 1Orig price: $249.87
                     *      EpsonExpressionET2750SpecialEditionEcoTankAllinOnePrinterItem # 980094600Qty: 1Orig price: $249.87
                     *      Payment info
                     * 
                     * 2) has shipping : one item has 2 1ine.
                     * 
                     *      ItemQtyOrig. price
                     *      SharkRocketDeluxeProUprightVacuumItem # 1031541$119.98Free Shipping
                     *      SharkRocketDeluxeProUprightVacuumItem # 1031541$119.98Free Shipping
                     *      SharkRocketDeluxeProUprightVacuumItem # 103154 |
                     *      Free ShippingQty: 1Orig price: $119.98
                     *      SharkRocketDeluxeProUprightVacuumItem # 103154 |
                     *      Free ShippingQty: 1Orig price: $119.98
                     *      Payment info
                     *      
                     * 3) has savings : one item has 2 1ine.
                     * 
                     *      ItemQtyOrig. priceTotal
                     *      LexmarkC2425dwWirelessColorLaserPrinterItem # 9801385891$209.98$449.90$120.00 off Instant Savings$600.00 savings included
                     *      LexmarkC2425dwWirelessColorLaserPrinterItem # 9801385891$209.98$449.90$120.00 off Instant Savings$600.00 savings included
                     *      LexmarkC2425dwWirelessColorLaserPrinterItem # 980138589 |
                     *      $120.00 off Instant SavingsQty: 1Orig price: $209.98Subtotal: $449.90$600.00 savings included
                     *      LexmarkC2425dwWirelessColorLaserPrinterItem # 980138589 |
                     *      $120.00 off Instant SavingsQty: 1Orig price: $209.98Subtotal: $449.90$600.00 savings included
                     *      Payment info
                     *      
                     * 4) has shipping, has savings : one item has 3 1ine.
                     * 
                     *      ItemQtyOrig. priceTotal
                     *      SAMSUNG32MONITORCURVEDItem # 9800062221$229.96$379.92$40.00 off Instant Savings$80.00 savings includedFree Shipping
                     *      SAMSUNG32MONITORCURVEDItem # 9800062221$229.96$379.92$40.00 off Instant Savings$80.00 savings includedFree Shipping
                     *      SAMSUNG32MONITORCURVEDItem # 980006222 |
                     *      $40.00 off Instant Savings |
                     *      Free ShippingQty: 1Orig price: $229.96Subtotal: $379.92$80.00 savings included
                     *      SAMSUNG32MONITORCURVEDItem # 980006222 |
                     *      $40.00 off Instant Savings |
                     *      Free ShippingQty: 1Orig price: $229.96Subtotal: $379.92$80.00 savings included
                     *      Payment info
                     *      
                     * 5) old kind. : It is handled above.
                     * 
                     *      ITEMQTYPRICE
                     *      EpsonWorkForceDS40Scanner5$119.98
                     *      Track this shipmentTrack this shipment
                     **/

                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;
                    int shipping_pos = line.IndexOf("Free Shipping", StringComparison.CurrentCultureIgnoreCase);
                    int savings_pos = line.IndexOf("savings included", StringComparison.CurrentCultureIgnoreCase);

                    if (shipping_pos == -1 && savings_pos == -1)
                    {
                        // EpsonExpressionET2750SpecialEditionEcoTankAllinOnePrinterItem # 980094600Qty: 1Orig price: $249.87
                        //                                                          ^      ^        ^    ^^           ^

                        int pos = line.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            continue;
                        title = line.Substring(0, pos);

                        string temp = line.Substring(pos + "Item #".Length).Trim();
                        pos = temp.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            continue;
                        string temp1 = temp.Substring(0, pos);
                        sku = temp1;

                        temp = temp.Substring(pos + "Qty:".Length).Trim();
                        pos = temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            continue;
                        temp1 = temp.Substring(0, pos);
                        qty = Str_Utils.string_to_int(temp1);
                        temp = temp.Substring(pos + "Orig price:".Length).Trim();
                        price = Str_Utils.string_to_currency(temp);
                    }
                    else
                    {
                        string title_line = (shipping_pos != -1 && savings_pos != -1) ? lines[i - 2].Trim() : lines[i - 1].Trim();

                        // SAMSUNG32MONITORCURVEDItem # 980006222 |
                        //                       ^      ^        ^

                        int pos = title_line.IndexOf("Item #", StringComparison.CurrentCultureIgnoreCase);
                        if (pos == -1)
                            continue;

                        title = title_line.Substring(0, pos);

                        string temp = title_line.Substring(pos + "Item #".Length).Trim();
                        pos = temp.IndexOf("|");
                        if (pos != -1)
                            temp = temp.Substring(0, pos).Trim();
                        sku = temp;

                        float savings = 0;

                        if (shipping_pos != -1 && savings_pos == -1)
                        {
                            // Free ShippingQty: 1Orig price: $119.98
                            //              ^    ^^           ^

                            pos = line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase);
                            temp = line.Substring(pos + "Qty:".Length).Trim();
                            pos = temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            string temp1 = temp.Substring(0, pos);
                            qty = Str_Utils.string_to_int(temp1);
                            temp = temp.Substring(pos + "Orig price:".Length).Trim();
                            price = Str_Utils.string_to_currency(temp);
                        }
                        else if (shipping_pos == -1)
                        {
                            // $120.00 off Instant SavingsQty: 1Orig price: $209.98Subtotal: $449.90$600.00 savings included
                            // ^      ^                   ^    ^^                  ^

                            pos = line.IndexOf("off Instant Savings", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            temp = line.Substring(0, pos).Trim();
                            savings = Str_Utils.string_to_currency(temp);

                            pos = line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            temp = line.Substring(pos + "Qty:".Length).Trim();
                            pos = temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            string temp1 = temp.Substring(0, pos);
                            qty = Str_Utils.string_to_int(temp1);
                            temp = temp.Substring(pos + "Orig price:".Length).Trim();
                            pos = temp.IndexOf("Subtotal:", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            temp = temp.Substring(0, pos).Trim();
                            price = Str_Utils.string_to_currency(temp);
                            price -= savings;
                        }
                        else
                        {
                            // $40.00 off Instant Savings |
                            // Free ShippingQty: 1Orig price: $229.96Subtotal: $379.92$80.00 savings included

                            temp = lines[i - 1].Trim();
                            pos = temp.IndexOf("off Instant Savings", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            temp = temp.Substring(0, pos).Trim();
                            savings = Str_Utils.string_to_currency(temp);

                            pos = line.IndexOf("Qty:", StringComparison.CurrentCultureIgnoreCase);
                            temp = line.Substring(pos + "Qty:".Length).Trim();
                            pos = temp.IndexOf("Orig price:", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            string temp1 = temp.Substring(0, pos);
                            qty = Str_Utils.string_to_int(temp1);
                            temp = temp.Substring(pos + "Orig price:".Length).Trim();
                            pos = temp.IndexOf("Subtotal:", StringComparison.CurrentCultureIgnoreCase);
                            if (pos == -1)
                                continue;
                            temp = temp.Substring(0, pos).Trim();
                            price = Str_Utils.string_to_currency(temp);
                            price -= savings;
                        }
                    }

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    report.m_product_items.Add(product);

                    MyLogger.Info($"... SC-8 qty = {qty}, price = {price}, sku = {sku}, item title = {title}");
                    continue;
                }
            }
        }
    }
}
