using MailParser;
using Logger;
using MailHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportStat
{
    public class GReport4Retailers
    {
        protected string gsheet_credential;
        protected string gsheet_id;
        protected string sheet_name;

        private int column_num = 5;
        private int col_idx_date = 1;
        private int col_idx_time = 2;
        private int col_idx_retailer = 3;
        private int col_idx_order = 4;
        private int col_idx_status = 5;
        private string col_name_date = "Date";
        private string col_name_time = "Time";
        private string col_name_retailer = "Retailer";
        private string col_name_order = "Order #";
        private string col_name_status = "Status";
        public GReport4Retailers()
        {
            gsheet_credential = Program.g_setting.gsheets_credential_json_file;

            gsheet_id = Program.g_setting.gsheets_id_4_retailers;
            sheet_name = Program.g_setting.gsheets_sheet_name_4_retailers;
        }
        public void report(KReportBase report)
        {
            try
            {
                if (!report.is_4_retailers())
                    return;

                string status = report.m_order_status;
                string retailer = report.m_retailer;

                var gsh = new GoogleSheetsHelper(gsheet_credential, gsheet_id);

                // Find if the row with the same order is already existed.

                int row_count = 0;

                var gsp = new GoogleSheetParameters() { RangeColumnStart = 1, RangeRowStart = 2, RangeColumnEnd = column_num, RangeRowEnd = -1, FirstRowIsHeaders = true, SheetName = sheet_name };
                var rowValues = gsh.GetDataFromSheet(gsp);
                row_count = rowValues.Count;

                int find_pos = 2 + rowValues.Count;
                int last_pos = 2 + rowValues.Count;

                string[] order_ids = report.m_order_id.Split(',');
                foreach (string order in order_ids)
                {
                    string old_status = "";
                    bool found = false;
                    if (order != "")
                    {
                        int i = 1;
                        foreach (var row1 in rowValues)
                        {
                            i++;
                            string order_in_row = row1[col_name_order].ToString();
                            string[] orders_in_row = order_in_row.Split(',');
                            if (orders_in_row.Count(s => s == order) > 0)
                            {
                                found = true;
                                find_pos = i;
                                old_status = row1[col_name_status].ToString();
                                break;
                            }
                        }
                    }
                    if (found)
                    {
                        if (old_status == status)
                            return;
                        if (old_status == ConstEnv.REPORT_ORDER_STATUS_CANCELED)
                            return;
                        if (old_status == ConstEnv.REPORT_ORDER_STATUS_SHIPPED && status != ConstEnv.REPORT_ORDER_STATUS_CANCELED)
                            return;
                    }

                    if (found)
                    {
                        var cell_status = new GoogleSheetCell() { CellValue = status };
                        gsh.UpdateOneCell(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = col_idx_status, RangeRowStart = find_pos }, cell_status);

                        MyLogger.Info($"*** Google Sheet *** : Update status to {status} : retailer = {retailer}, order = {report.m_order_id}");
                    }
                    else
                    {
                        var row = new GoogleSheetRow();

                        var cell_date = new GoogleSheetCell() { CellValue = report.m_mail_sent_date.ToString("yyyy-MM-dd") };
                        var cell_time = new GoogleSheetCell() { CellValue = report.m_mail_sent_date.ToString("HH:mm") };
                        var cell_retailer = new GoogleSheetCell() { CellValue = retailer };
                        var cell_order = new GoogleSheetCell() { CellValue = order };
                        var cell_status = new GoogleSheetCell() { CellValue = status };

                        row.Cells.AddRange(new List<GoogleSheetCell>() { cell_date, cell_time, cell_retailer, cell_order, cell_status });
                        var rows = new List<GoogleSheetRow>() { row };
                        gsh.AddCells(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = 1, RangeRowStart = last_pos }, rows);

                        MyLogger.Info($"*** Google Sheet *** : Add status : retailer = {retailer}, order = {report.m_order_id}, status = {status}");

                        last_pos++;
                    }
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
        public void monitor()
        {
            try
            {
                var gsh = new GoogleSheetsHelper(gsheet_credential, gsheet_id);

                int row_count = 0;

                var gsp = new GoogleSheetParameters() { RangeColumnStart = 1, RangeRowStart = 2, RangeColumnEnd = column_num, RangeRowEnd = -1, FirstRowIsHeaders = true, SheetName = sheet_name };
                var rowValues = gsh.GetDataFromSheet(gsp);
                row_count = rowValues.Count;

                int i = 1;
                foreach (var row1 in rowValues)
                {
                    i++;

                    DateTime cell_date = DateTime.Parse(row1[col_name_date].ToString());
                    DateTime cell_time = DateTime.Parse(row1[col_name_time].ToString());
                    string status = row1[col_name_status].ToString();
                    string order = row1[col_name_order].ToString();
                    string retailer = row1[col_name_retailer].ToString();

                    DateTime last_time = new DateTime(cell_date.Year, cell_date.Month, cell_date.Day, cell_time.Hour, cell_time.Minute, 0);
                    TimeSpan time_diff = DateTime.Now - last_time;

                    if (status == ConstEnv.REPORT_ORDER_STATUS_PURCHAESD && time_diff.TotalHours >= Program.g_setting.status_manual_check_hours)
                    {
                        var cell_status = new GoogleSheetCell() { CellValue = ConstEnv.REPORT_ORDER_STATUS_MANUAL_CHECK };
                        gsh.UpdateOneCell(new GoogleSheetParameters() { SheetName = sheet_name, RangeColumnStart = col_idx_status, RangeRowStart = i }, cell_status);
                        MyLogger.Info($"*** Google Sheet *** : Need Manual Check : retailer = {retailer}, order = {order}");
                    }
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }
    }
}
