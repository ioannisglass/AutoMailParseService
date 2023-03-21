using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class ZProduct
    {
        public string title;
        public string sku;
        public int qty;
        public float price;
        public string status; // Only valid in OP and SC if it is not shipped or canceled.

        public ZProduct()
        {
            title = "";
            sku = "";
            qty = 0;
            price = 0;
            status = "";
        }
        public string make_json_text()
        {
            JObject jsonObject;

            if (status == "")
            {
                jsonObject = new JObject(
                            new JProperty("title", title),
                            new JProperty("sku", sku),
                            new JProperty("qty", qty),
                            new JProperty("price", price)
                );
            }
            else
            {

                jsonObject = new JObject(
                            new JProperty("title", title),
                            new JProperty("sku", sku),
                            new JProperty("qty", qty),
                            new JProperty("price", price),
                            new JProperty("status", status)
                );
            }

            return JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
        }
    }
}
