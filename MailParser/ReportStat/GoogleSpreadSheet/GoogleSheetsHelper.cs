using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.IO;
using System.Dynamic;

namespace ReportStat
{
    public class GoogleSheetsHelper
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheetsHelper";

        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;

        public GoogleSheetsHelper(string credentialFileName, string spreadsheetId)
        {
            var credential = GoogleCredential.FromStream(new FileStream(credentialFileName, FileMode.Open)).CreateScoped(Scopes);

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            _spreadsheetId = spreadsheetId;
        }
        public List<string> GetColumnHeaders(string sheet_name)
        {
            var columnNames = new List<string>();
            var range = GetA1Notation(sheet_name, -1, 1, -1, 1);
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count == 1)
            {
                var row = values[0];
                {
                    foreach (var r in row)
                        columnNames.Add(r.ToString());
                }
            }

            return columnNames;
        }

        public List<Dictionary<String, object>> GetDataFromSheet(GoogleSheetParameters googleSheetParameters)
        {
            googleSheetParameters = MakeGoogleSheetDataRangeColumnsZeroBased(googleSheetParameters);
            var range = GetA1Notation(googleSheetParameters.SheetName, googleSheetParameters.RangeColumnStart, googleSheetParameters.RangeRowStart, googleSheetParameters.RangeColumnEnd, googleSheetParameters.RangeRowEnd);

            SpreadsheetsResource.ValuesResource.GetRequest request =
                _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);

            var numberOfColumns = googleSheetParameters.RangeColumnEnd - googleSheetParameters.RangeColumnStart + 1;
            var columnNames = new List<string>();
            var returnValues = new List<Dictionary<String, object>>();

            if (!googleSheetParameters.FirstRowIsHeaders)
            {
                for (var i = 0; i < numberOfColumns; i++)
                {
                    columnNames.Add($"Column{i}");
                }
            }
            else
            {
                columnNames = GetColumnHeaders(googleSheetParameters.SheetName);
            }

            var response = request.Execute();

            int rowCounter = 0;
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    if (googleSheetParameters.FirstRowIsHeaders && rowCounter == 0 && googleSheetParameters.RangeRowStart == 1)
                    {
                        rowCounter++;
                        continue;
                    }

                    var expandoDict = new Dictionary<String, object>();
                    var columnCounter = 0;
                    foreach (var columnName in columnNames)
                    {
                        if (columnCounter >= row.Count)
                        {
                            expandoDict.Add(columnName, "");
                        }
                        else
                        {
                            expandoDict.Add(columnName, row[columnCounter].ToString());
                        }
                        columnCounter++;
                    }
                    returnValues.Add(expandoDict);
                    rowCounter++;
                }
            }

            return returnValues;
        }

        public void AddCells(GoogleSheetParameters googleSheetParameters, List<GoogleSheetRow> rows)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };

            var sheetId = GetSheetId(_sheetsService, _spreadsheetId, googleSheetParameters.SheetName);

            GridCoordinate gc = new GridCoordinate
            {
                ColumnIndex = googleSheetParameters.RangeColumnStart - 1,
                RowIndex = googleSheetParameters.RangeRowStart - 1,
                SheetId = sheetId
            };

            var request = new Request { UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" } };

            var listRowData = new List<RowData>();

            foreach (var row in rows)
            {
                var rowData = new RowData();
                var listCellData = new List<CellData>();
                foreach (var cell in row.Cells)
                {
                    var cellData = new CellData();
                    var extendedValue = new ExtendedValue { StringValue = cell.CellValue };

                    cellData.UserEnteredValue = extendedValue;
                    var cellFormat = new CellFormat { TextFormat = new TextFormat() };

                    if (cell.IsBold)
                    {
                        cellFormat.TextFormat.Bold = true;
                    }

                    //cellFormat.BackgroundColor = new Color { Blue = (float)cell.BackgroundColor.B / 255, Red = (float)cell.BackgroundColor.R / 255, Green = (float)cell.BackgroundColor.G / 255 };

                    cellData.UserEnteredFormat = cellFormat;
                    listCellData.Add(cellData);
                }
                rowData.Values = listCellData;
                listRowData.Add(rowData);
            }
            request.UpdateCells.Rows = listRowData;

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            requests.Requests.Add(request);

            _sheetsService.Spreadsheets.BatchUpdate(requests, _spreadsheetId).Execute();
        }
        public void UpdateOneCell(GoogleSheetParameters googleSheetParameters, GoogleSheetCell cell)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };

            var sheetId = GetSheetId(_sheetsService, _spreadsheetId, googleSheetParameters.SheetName);

            GridCoordinate gc = new GridCoordinate
            {
                ColumnIndex = googleSheetParameters.RangeColumnStart - 1,
                RowIndex = googleSheetParameters.RangeRowStart - 1,
                SheetId = sheetId
            };

            var request = new Request { UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" } };

            var listRowData = new List<RowData>();

            var rowData = new RowData();
            var listCellData = new List<CellData>();

            var cellData = new CellData();
            var extendedValue = new ExtendedValue { StringValue = cell.CellValue };

            cellData.UserEnteredValue = extendedValue;
            var cellFormat = new CellFormat { TextFormat = new TextFormat() };

            if (cell.IsBold)
            {
                cellFormat.TextFormat.Bold = true;
            }

            //cellFormat.BackgroundColor = new Color { Blue = (float)cell.BackgroundColor.B / 255, Red = (float)cell.BackgroundColor.R / 255, Green = (float)cell.BackgroundColor.G / 255 };

            cellData.UserEnteredFormat = cellFormat;
            listCellData.Add(cellData);

            rowData.Values = listCellData;
            listRowData.Add(rowData);

            request.UpdateCells.Rows = listRowData;

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            requests.Requests.Add(request);

            _sheetsService.Spreadsheets.BatchUpdate(requests, _spreadsheetId).Execute();
        }
        public void MergeCells(GoogleSheetParameters googleSheetParameters)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };

            var sheetId = GetSheetId(_sheetsService, _spreadsheetId, googleSheetParameters.SheetName);

            GridRange grange = new GridRange
            {
                StartRowIndex = googleSheetParameters.RangeRowStart - 1,
                StartColumnIndex = googleSheetParameters.RangeColumnStart - 1,
                EndRowIndex = googleSheetParameters.RangeRowEnd - 1,
                EndColumnIndex = googleSheetParameters.RangeColumnEnd - 1,
                SheetId = sheetId
            };

            var request = new Request { MergeCells = new MergeCellsRequest { Range  = grange, MergeType  = "MERGE_ALL" } };

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            requests.Requests.Add(request);

            _sheetsService.Spreadsheets.BatchUpdate(requests, _spreadsheetId).Execute();
        }

        private string GetColumnName(int index)
        {
            if (index == -1)
                return "";
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];
            return value;
        }
        private string GetRowRangeString(int index)
        {
            if (index == -1)
                return "";
            return index.ToString();
        }
        private string GetA1Notation(string sheet_name, int col_start, int row_start, int col_end, int row_end)
        {
            /**
             *      Sheet1!A1:B2 refers to the first two cells in the top two rows of Sheet1.
             *      Sheet1!A:A refers to all the cells in the first column of Sheet1.
             *      Sheet1!1:2 refers to the all the cells in the first two rows of Sheet1.
             *      Sheet1!A5:A refers to all the cells of the first column of Sheet 1, from row 5 onward.
             *      A1:B2 refers to the first two cells in the top two rows of the first visible sheet.
             *      Sheet1 refers to all the cells in Sheet1.
             * 
             *
             **/
            string a1_notation = "";
            if (sheet_name == "" || (col_start == -1 && row_start == -1 && col_end == -1 && row_end == -1))
                a1_notation = sheet_name;
            else
                a1_notation += "!";

            if (col_start == -1 && row_start == -1 && col_end == -1 && row_end == -1)
                return a1_notation;

            string r = $"{GetColumnName(col_start)}{GetRowRangeString(row_start)}:{GetColumnName(col_end)}{GetRowRangeString(row_end)}";
            return a1_notation + r;
        }

        private GoogleSheetParameters MakeGoogleSheetDataRangeColumnsZeroBased(GoogleSheetParameters googleSheetParameters)
        {
            googleSheetParameters.RangeColumnStart = googleSheetParameters.RangeColumnStart - 1;
            googleSheetParameters.RangeColumnEnd = googleSheetParameters.RangeColumnEnd - 1;
            return googleSheetParameters;
        }

        private int GetSheetId(SheetsService service, string spreadSheetId, string spreadSheetName)
        {
            var spreadsheet = service.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == spreadSheetName);
            int sheetId = (int)sheet.Properties.SheetId;
            return sheetId;
        }
    }
    public class GoogleSheetCell
    {
        public string CellValue { get; set; }
        public bool IsBold { get; set; }
        //public System.Drawing.Color BackgroundColor { get; set; } = System.Drawing.Color.White;
    }

    public class GoogleSheetParameters
    {
        public int RangeColumnStart { get; set; }
        public int RangeRowStart { get; set; }
        public int RangeColumnEnd { get; set; }
        public int RangeRowEnd { get; set; }
        public string SheetName { get; set; }
        public bool FirstRowIsHeaders { get; set; }
    }

    public class GoogleSheetRow
    {
        public GoogleSheetRow() => Cells = new List<GoogleSheetCell>();

        public List<GoogleSheetCell> Cells { get; set; }
    }
}
