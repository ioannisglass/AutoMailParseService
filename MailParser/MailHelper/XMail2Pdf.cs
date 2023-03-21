using MailParser;
using Logger;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailHelper
{
    public class XMail2Pdf
    {
        public XMail2Pdf()
        {

        }
        public static string eml_to_pdf(string eml_file_path, string order_id, string retailer)
        {
            string ret = ConstEnv.PDF_CONVERT_FAILED;
            string temp_html_file = "";

            try
            {
                if (!File.Exists(eml_file_path))
                    return ConstEnv.PDF_CONVERT_FAILED;
                if (Program.g_must_end)
                    return ConstEnv.PDF_CONVERT_CANCELED;

                string eml_folder_path = Path.GetDirectoryName(eml_file_path);
                string pdf_file_name = "";

                if (order_id == "")
                {
                    pdf_file_name = $"{retailer}_no_order_id.pdf";
                }
                else
                {
                    pdf_file_name = $"{retailer}_{order_id}.pdf";
                }

                string outpdf_file = Path.Combine(eml_folder_path, pdf_file_name);
                if (File.Exists(outpdf_file))
                    return outpdf_file;

                temp_html_file = Path.ChangeExtension(eml_file_path, "html");
                if (File.Exists(temp_html_file))
                    File.Delete(temp_html_file);
                //if (File.Exists(outpdf_file))
                //    File.Delete(outpdf_file);

                // save as html.

                MimeMessage mail = MimeMessage.Load(eml_file_path);
                if (mail == null)
                    throw new Exception("Invalid eml file");

                string html_text = mail.HtmlBody;

                using (StreamWriter sw = new StreamWriter(File.Open(temp_html_file, FileMode.Create), Encoding.UTF8))
                {
                    sw.Write(html_text);
                }

                // html -> eml

                string arguments = $"--viewport-size 2480x3508 --quiet --image-quality 100 --encoding utf-8 \"{temp_html_file}\" \"{outpdf_file}\"";

                string converter_path = (ConstEnv.OS_TYPE == ConstEnv.OS_WINDOWS) ? $"wkhtmltopdf.exe" : $"{Path.DirectorySeparatorChar}usr{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}wkhtmltopdf";

                ProcessStartInfo start = new ProcessStartInfo();
                start.Arguments = arguments;
                start.FileName = converter_path;
                start.WindowStyle = ProcessWindowStyle.Hidden;
                start.CreateNoWindow = true;

                MyLogger.Info($"start eml->pdf : {eml_file_path} -> {outpdf_file}");

                using (Process proc = Process.Start(start))
                {
                    proc.WaitForExit();
                    int exitCode = proc.ExitCode;

                    MyLogger.Info($"end eml->pdf exit code = {exitCode} : {eml_file_path} -> {outpdf_file}");
                }
                if (!File.Exists(outpdf_file))
                {
                    MyLogger.Info($"Failed eml->pdf : {eml_file_path} -> {outpdf_file}");
                }
                else
                {
                    ret = outpdf_file;
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
            finally
            {
                if (temp_html_file != "" && File.Exists(temp_html_file))
                    File.Delete(temp_html_file);
            }
            return ret;
        }
    }
}
