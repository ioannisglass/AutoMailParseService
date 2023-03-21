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
        private void parse_mail_cc_2(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_2;

            card.m_retailer = ConstEnv.RETAILER_TARGET;

            MyLogger.Info($"... CC-2 m_cc_retailer = {card.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("order #", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER #")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("order #".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-2 order id = {temp}");
                    continue;
                }
                if (line.IndexOf("Order #", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    string temp = line.Substring(line.IndexOf("Order #", StringComparison.CurrentCultureIgnoreCase) + "Order #".Length).Trim();
                    temp += lines[++i].Trim(); // In some cases, order is continued with the next line.
                    if (temp.IndexOf(")") != -1)
                        temp = temp.Substring(0, temp.IndexOf(")")).Trim();
                    if (temp.IndexOf(".") != -1)
                        temp = temp.Substring(0, temp.IndexOf(".")).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-2 order id = {temp}");
                    continue;
                }
                if (line.StartsWith("qty:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string title = "";
                    string sku = "";
                    int qty = 0;
                    float price = 0;

                    string temp;
                    if (line.ToUpper() == "QTY:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("qty:".Length).Trim();
                    qty = Str_Utils.string_to_int(temp);

                    title = lines[i - 1].Trim();

                    ZProduct product = new ZProduct();
                    product.price = price;
                    product.sku = sku;
                    product.title = title;
                    product.qty = qty;
                    card.m_product_items.Add(product);

                    MyLogger.Info($"... CC-2 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
                    continue;
                }
                if (line.ToUpper() == "ITEMS" && lines[i + 1].Trim().ToUpper() == "QUANTITY" && lines[i + 2].Trim().ToUpper() == "SHIPPING METHOD" && lines[i + 3].Trim().ToUpper() == "EST.DELIVERY DATE" && lines[i + 4].Trim().ToUpper() == "PRICE")
                {
                    get_cc2_items_from_html(mail, card);

                    i += 5;
                    continue;
                }
                if (line.ToUpper() == "ESTIMATED TAXES:")
                {
                    string temp = lines[++i].ToString();
                    float tax = Str_Utils.string_to_currency(temp);

                    card.m_tax = tax;
                    MyLogger.Info($"... CC-2 tax = {tax}");
                    continue;
                }
            }
        }
        private void get_cc2_items_from_html(MimeMessage mail, KReportCC card)
        {
            string html_text = XMailHelper.get_htmltext(mail);
            int pos;
            string temp;
            int next_pos;

            pos = html_text.IndexOf("est.delivery date", StringComparison.CurrentCultureIgnoreCase);
            if (pos == -1)
                return;
            temp = html_text.Substring(0, pos);
            pos = temp.LastIndexOf("<tbody>");
            if (pos == -1)
                return;
            temp = html_text.Substring(pos);

            string items_body_text = XMailHelper.find_html_part(temp, "tbody", out next_pos);
            if (items_body_text == "")
                return;

            temp = items_body_text;
            string tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos); // canceled item LABEL
            temp = temp.Substring(next_pos);
            if (tr_part == "" || temp == "")
                return;

            tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos); // items TITLE (items, quantity, shipping method, est.delivery date, price)
            temp = temp.Substring(next_pos);
            if (tr_part == "" || temp == "")
                return;

            tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos); // split line
            temp = temp.Substring(next_pos);
            if (tr_part == "" || temp == "")
                return;

            while (temp != "")
            {
                string title = "";
                string sku = "";
                int qty = 0;
                float price = 0;

                tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
                temp = temp.Substring(next_pos).Trim();
                if (tr_part == "" || temp == "")
                    break;

                string tr_text = XMailHelper.html2text(tr_part);
                if (tr_text == "" || tr_text.IndexOf("$") == -1)
                    continue;

                int next_pos1;
                string temp1 = tr_part;
                string td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1); // item picture
                temp1 = temp1.Substring(next_pos1).Trim();
                if (td_part == "" || temp1 == "")
                    break;

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1); // qty
                temp1 = temp1.Substring(next_pos1).Trim();
                if (td_part == "" || temp1 == "")
                    break;
                string td_text = XMailHelper.html2text(td_part);
                qty = Str_Utils.string_to_int(td_text);

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1); // title
                temp1 = temp1.Substring(next_pos1).Trim();
                if (td_part == "" || temp1 == "")
                    break;
                td_text = XMailHelper.html2text(td_part);
                td_text = td_text.Replace("\r\n", " ");
                td_text = td_text.Replace("\n", " ");
                title = td_text;

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1); // shipping method
                temp1 = temp1.Substring(next_pos1).Trim();
                if (td_part == "" || temp1 == "")
                    break;

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1); // est.delivery date
                temp1 = temp1.Substring(next_pos1).Trim();
                if (td_part == "" || temp1 == "")
                    break;

                td_part = XMailHelper.find_html_part(temp1, "td", out next_pos1); // price
                temp1 = temp1.Substring(next_pos1).Trim();
                if (td_part == "")
                    break;
                td_text = XMailHelper.html2text(td_part);
                float f = Str_Utils.string_to_currency(td_text);
                price = (qty == 0) ? f : f / qty;

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                card.m_product_items.Add(product);

                MyLogger.Info($"... CC-2 qty = {qty}, price = {price}, sku = {sku}, title = {title}");
            }
        }
    }
}
