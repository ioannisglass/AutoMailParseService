using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportStat
{
    public class GReportTax
    {
        protected string gsheet_credential;
        protected string gsheet_id;
        protected string sheet_name;
        private int column_num = 8;
        private string col_name_retailer = "Retailer";
        private string col_name_order = "Order #";

        private GoogleSheetsHelper gsh;
        int raw_num;
        public GReportTax()
        {
            gsheet_credential = Program.g_setting.gsheets_credential_json_file;

            gsheet_id = Program.g_setting.gsheets_id_tax_step2;
            sheet_name = Program.g_setting.gsheets_sheet_name_tax_step2;

            gsh = new GoogleSheetsHelper(gsheet_credential, gsheet_id);
            //var gsp = new GoogleSheetParameters() { RangeColumnStart = 1, RangeRowStart = 1, RangeColumnEnd = column_num, RangeRowEnd = -1, FirstRowIsHeaders = true, SheetName = sheet_name };
            //var rowValues = gsh.GetDataFromSheet(gsp);
            //raw_num = rowValues.Count + 1; // 1 : header

            raw_num = 1;
        }
        public void add_tax_gsheet(ZSalesTaxPayData data)
        {
            var rows = new List<GoogleSheetRow>();

            for (int i = 0; i < data.payments.Count; i++)
            {
                ZPaymentCard payment = data.payments[i];

                var row = new GoogleSheetRow();

                var cell_time = new GoogleSheetCell() { CellValue = (i == 0) ? data.purchase_time.ToString("yyyy-MM-dd") : "" };
                var cell_order = new GoogleSheetCell() { CellValue = (i == 0) ? data.order_id : "" };
                var cell_retailer = new GoogleSheetCell() { CellValue = (i == 0) ? data.retailer : "" };
                var cell_total = new GoogleSheetCell() { CellValue = (i == 0) ? data.total.ToString() : "" };
                var cell_tax = new GoogleSheetCell() { CellValue = (i == 0) ? data.tax.ToString() : "" };

                var cell_payment_type = new GoogleSheetCell() { CellValue = payment.payment_type };
                var cell_last_4_digits = new GoogleSheetCell() { CellValue = payment.last_4_digit };
                var cell_price = new GoogleSheetCell() { CellValue = payment.price.ToString() };

                row.Cells.AddRange(new List<GoogleSheetCell>() { cell_time, cell_order, cell_retailer, cell_total, cell_tax, cell_payment_type, cell_last_4_digits, cell_price });

                rows.Add(row);

                MyLogger.Info($"*** Google Sheet *** : Add tax (payment Info) : raw_num = {raw_num}, type = {payment.payment_type}, last_4_digits = {payment.last_4_digit}, price = {payment.price}");
            }
            gsh.AddCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 1, RangeRowStart = raw_num + 1 }, rows);

            if (data.payments.Count > 1)
            {
                gsh.MergeCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 1, RangeColumnEnd = 2, RangeRowStart = raw_num + 1, RangeRowEnd = raw_num + 1 + data.payments.Count });
                gsh.MergeCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 2, RangeColumnEnd = 3, RangeRowStart = raw_num + 1, RangeRowEnd = raw_num + 1 + data.payments.Count });
                gsh.MergeCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 3, RangeColumnEnd = 4, RangeRowStart = raw_num + 1, RangeRowEnd = raw_num + 1 + data.payments.Count });
                gsh.MergeCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 4, RangeColumnEnd = 5, RangeRowStart = raw_num + 1, RangeRowEnd = raw_num + 1 + data.payments.Count });
                gsh.MergeCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 5, RangeColumnEnd = 6, RangeRowStart = raw_num + 1, RangeRowEnd = raw_num + 1 + data.payments.Count });
            }
            raw_num += rows.Count;
        }
    }
}
