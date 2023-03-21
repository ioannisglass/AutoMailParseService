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
        private void parse_mail_cc_7(MimeMessage mail, KReportCC card)
        {
            string subject = XMailHelper.get_subject(mail);

            card.m_mail_type = KReportBase.MailType.CC_7;

            card.m_retailer = ConstEnv.RETAILER_SAMSCLUB;

            MyLogger.Info($"... CC-7 m_cc_retailer = {card.m_retailer}");

            string[] lines = XMailHelper.get_bodytext2(mail).Replace("\r", "").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Order Number:", StringComparison.CurrentCultureIgnoreCase))
                {
                    string temp;
                    if (line.ToUpper() == "ORDER NUMBER:")
                        temp = lines[++i].Trim();
                    else
                        temp = line.Substring("Order Number:".Length).Trim();
                    card.set_order_id(temp);
                    MyLogger.Info($"... CC-7 order id = {temp}");
                    continue;
                }
            }
            get_cc7_items(mail, card);
        }
        private void get_cc7_items(MimeMessage mail, KReportCC card)
        {
            string html_text = XMailHelper.get_htmltext(mail);
            int next_pos;
            string temp;
            int pos;

            pos = html_text.IndexOf("qty", StringComparison.CurrentCultureIgnoreCase);
            if (pos == -1)
                return;
            pos = html_text.IndexOf("</tr>", pos);
            if (pos == -1)
                return;
            pos += "</tr>".Length;
            temp = html_text.Substring(pos).Trim();

            string tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);

            // ignore the first tr.
            temp = temp.Substring(next_pos).Trim();

            tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
            string tr_text = XMailHelper.html2text(tr_part);
            while (tr_text != "")
            {
                string title = "";
                string sku = "";
                int qty = 1;
                float price = 0;

                int next_pos1;
                string td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos1);
                string td_text = XMailHelper.html2text(td_part);
                td_text = td_text.Replace("\r\n", " ");
                td_text = td_text.Replace("\n", " ");
                title = td_text.Trim();

                tr_part = tr_part.Substring(next_pos1);
                td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos1);
                td_text = XMailHelper.html2text(td_part);

                qty = Str_Utils.string_to_int(td_text);

                tr_part = tr_part.Substring(next_pos1);
                td_part = XMailHelper.find_html_part(tr_part, "td", out next_pos1);
                td_text = XMailHelper.html2text(td_part);

                price = Str_Utils.string_to_currency(td_text);

                ZProduct product = new ZProduct();
                product.price = price;
                product.sku = sku;
                product.title = title;
                product.qty = qty;
                card.m_product_items.Add(product);

                MyLogger.Info($"... CC-2 qty = {qty}, price = {price}, sku = {sku}, title = {title}");

                temp = temp.Substring(next_pos);
                tr_part = XMailHelper.find_html_part(temp, "tr", out next_pos);
                tr_text = XMailHelper.html2text(tr_part);
            }
        }
    }
}
