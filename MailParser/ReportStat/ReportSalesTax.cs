using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ReportStat
{
    public class ReportSalesTax
    {

        public ReportSalesTax()
        {
        }
        public void collect_tax_data(int account_id)
        {
            try
            {
                GReportTax greport = new GReportTax();

                MyLogger.Info("*** Google Sheet *** : START MySQL Query.");

                DataTable dt = Program.g_db.collect_sales_tax_order_data(account_id);
                if (dt == null || dt.Rows == null || dt.Rows.Count == 0)
                    return;

                MyLogger.Info($"*** Google Sheet *** : END MySQL Query. num = {dt.Rows.Count}");

                int num = 0;
                foreach (DataRow row in dt.Rows)
                {
                    DateTime time = DateTime.Parse(row["time"].ToString());
                    string order_id = row["order_id"].ToString();
                    string retailer = row["retailer"].ToString();
                    float total = float.Parse(row["total"].ToString());
                    float tax = float.Parse(row["tax"].ToString());
                    int op_report_id = int.Parse(row["op_id"].ToString());

                    DataTable dt_pay = Program.g_db.collect_sales_tax_payments(op_report_id);
                    if (dt_pay == null || dt_pay.Rows == null || dt_pay.Rows.Count == 0)
                        continue;

                    List<ZPaymentCard> payments = new List<ZPaymentCard>();
                    foreach (DataRow row1 in dt_pay.Rows)
                    {
                        string payment_type = row1["payment_type"].ToString();
                        string last_4_digit = row1["last_4_digit"].ToString();
                        float price = float.Parse(row1["price"].ToString());

                        payments.Add(new ZPaymentCard(payment_type, last_4_digit, price));
                    }

                    MyLogger.Info($"*** Google Sheet *** : Add tax : order = {order_id}, retailer = {retailer}, time = {time.ToString("yyyy-MM-dd")}, total = {total}, tax = {tax}");

                    greport.add_tax_gsheet(new ZSalesTaxPayData(time, order_id, retailer, total, tax, payments));

                    num++;
                    MyLogger.Info($"*** Google Sheet *** : Remained tax data : {dt.Rows.Count - num}");

                    Thread.Sleep(500);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
    }
}
