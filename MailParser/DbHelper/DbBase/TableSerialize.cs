using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Data;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace DbHelper
{
    public static class TableSerialize<T> where T : new()
    {
        public static List<string> get_column_names()
        {
            T t = new T();
            List<string> column_names = new List<string>();

            string json_str = (new JavaScriptSerializer()).Serialize(t);
            JObject jsonObj = JObject.Parse(json_str);
            column_names = jsonObj.Properties().Select(p => p.Name).ToList();

            return column_names;
        }
        public static DataTable create_empty_dt()
        {
            DataTable dt = new DataTable();

            List<string> column_names = get_column_names();

            foreach (string x in column_names)
                dt.Columns.Add(x);

            return dt;
        }
        public static DataTable to_dt(T t)
        {
            DataTable dt = new DataTable();

            List<string> column_names = new List<string>();

            string json_str = (new JavaScriptSerializer()).Serialize(t);
            JObject jsonObj = JObject.Parse(json_str);
            Dictionary<string, string> dictObj = jsonObj.ToObject<Dictionary<string, string>>();

            foreach (var pair in dictObj)
                dt.Columns.Add(pair.Key);

            DataRow row = dt.NewRow();
            foreach (var pair in dictObj)
            {
                row[pair.Key] = pair.Value;
            }
            dt.Rows.Add(row);

            return dt;
        }
        public static T from_dt(DataTable dt)
        {
            List<T> t_list = dt_2_list(dt);
            if (t_list == null || t_list.Count != 1)
                return new T();
            return t_list[0];
        }
        public static List<T> dt_2_list(DataTable dt)
        {
            if (dt == null)
                return null;
            List<T> t_list = new List<T>();
            string json_str = JsonConvert.SerializeObject(dt, Formatting.Indented);
            t_list = JsonConvert.DeserializeObject<List<T>>(json_str);

            return t_list;
        }
    }
}
