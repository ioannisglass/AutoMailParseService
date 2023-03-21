using CsvHelper;
using Logger;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    class CSVUtil
    {
        public static void dt2csv(DataTable dt, string strFilePath)
        {
            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine("\"" + string.Join("\",\"", columnNames) + "\"");

            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                List<string> escaped = new List<string>();
                foreach (var f in fields)
                {
                    escaped.Add($"\"{f}\"");
                }
                sb.AppendLine(string.Join(",", escaped.ToArray()));
            }

            File.WriteAllText(strFilePath, sb.ToString());
        }

        public static string[] preprocess_csv_row(string line)
        {
            line = line.Trim();
            if (line.EndsWith(","))
                line = line.Substring(0, line.Length - 1);
            List<string> parsed = new List<string>();
            string[] fields = line.Split(',');
            for (int i = 0; i < fields.Length; i++)
            {
                if (fields[i].StartsWith("\""))
                {
                    string val = "";

                    int cnt = 0;
                    while (i < fields.Length)
                    {
                        cnt += fields[i].Count(x => x == '"');
                        val += fields[i];
                        if (fields[i].EndsWith("\"") && cnt % 2 == 0)
                            break;
                        i++;
                    }

                    try
                    {
                        parsed.Add(val.Substring(1, val.Length - 2));
                    }
                    catch
                    {
                        System.Diagnostics.Debug.WriteLine($"Error occured parsing CSV line:\n{line}\n{i}");
                    }
                }
                else
                    parsed.Add(fields[i]);
            }
            return parsed.ToArray();
        }

        public static string[] preprocess_csv_row_strong(string line)
        {
            try
            {
                // we assume the fields are delimetered by , and all fields are surrounded by "
                line = line.Trim();
                if (line.EndsWith(","))
                    line = line.Substring(0, line.Length - 1);
                line = line.Substring(1, line.Length - 2);
                return line.Split(new string[] { "\",\"" }, StringSplitOptions.None);
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error csv line = {line}, ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return null;
            }
        }


        public static DataTable csv2dt_strong(string strFilePath)
        {
            try
            {
                string tbl_name = Path.GetFileNameWithoutExtension(strFilePath);
                DataTable dt = new DataTable(tbl_name);
                string[] lines = File.ReadAllLines(strFilePath);

                for (int ln = 0; ln < lines.Length; ln++)
                {
                    string line = lines[ln];
                    string[] fields = preprocess_csv_row_strong(line);
                    if (ln == 0)
                    {
                        foreach (string field in fields)
                        {
                            if (dt.Columns.IndexOf(field) >= 0)
                            {
                                int cnt = 1;
                                while (true)
                                {
                                    string new_field = $"{field}_{cnt}";
                                    if (dt.Columns.IndexOf(field) < 0)
                                    {
                                        dt.Columns.Add(new_field);
                                        break;
                                    }
                                    cnt++;
                                }
                            }
                            else
                                dt.Columns.Add(field);
                        }
                    }
                    else
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < fields.Length; i++)
                        {
                            dr[i] = fields[i];
                        }
                        dt.Rows.Add(dr);
                    }
                }

                foreach (DataColumn cl in dt.Columns)
                    cl.DataType = typeof(string);
                return dt;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return null;
            }
        }

        /*public static DataTable csv2dt(string strFilePath)
        {
            try
            {
                string tbl_name = Path.GetFileNameWithoutExtension(strFilePath);
                DataTable dt = new DataTable(tbl_name);
                using (TextFieldParser parser = new TextFieldParser(strFilePath))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    bool first = true;
                    int ln = 0;
                    while (!parser.EndOfData)
                    {
                        ln++;
                        string[] fields = parser.ReadFields();
                        if (first)
                        {
                            first = false;
                            foreach (string field in fields)
                            {
                                if (dt.Columns.IndexOf(field) >= 0)
                                {
                                    int cnt = 1;
                                    while (true)
                                    {
                                        string new_field = $"{field}_{cnt}";
                                        if (dt.Columns.IndexOf(field) < 0)
                                        {
                                            dt.Columns.Add(new_field);
                                            break;
                                        }
                                        cnt++;
                                    }
                                }
                                else
                                    dt.Columns.Add(field);
                            }
                        }
                        else
                        {
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < fields.Length; i++)
                            {
                                dr[i] = fields[i];
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                    parser.Close();
                }

                foreach (DataColumn cl in dt.Columns)
                    cl.DataType = typeof(string);
                return dt;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message + "\n" + ex.StackTrace);
                return null;
            }
        }*/

        public static DataTable csv2dt(string strFilePath)
        {
            try
            {
                string tbl_name = Path.GetFileNameWithoutExtension(strFilePath);
                DataTable dt = new DataTable(tbl_name);

                using (var reader = new StreamReader(strFilePath))
                using (var csv = new CsvReader(reader))
                {
                    // Do any configuration to `CsvReader` before creating CsvDataReader.
                    using (var dr = new CsvDataReader(csv))
                    {
                        dt.Load(dr);
                    }
                }

                return dt;
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return null;
            }
        }
    }
}
