using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using MailParser;
using Logger;
using PdfHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatementHelper
{
    public class BSParser
    {
        protected string name;
        public Dictionary<string, List<BankTransactions>> transactions;
        public BSParser(string _name)
        {
            name = _name;
            transactions = new Dictionary<string, List<BankTransactions>>();
        }
        public string get_bank_name()
        {
            return name;
        }
        public bool parse(string pdf_file)
        {
            //try
            {
                int parsing_state = Program.g_db.get_statment_file_parsing_state(pdf_file);
                if (parsing_state == ConstEnv.STATEMENT_FILE_PARSE_IMAGE || parsing_state == ConstEnv.STATEMENT_FILE_PARSE_SUCCEED)
                {
                    MyLogger.Info($"Already parsed : state = {parsing_state}, {pdf_file}");
                    return true;
                }

                string pdf_text = "";

                if (!File.Exists(pdf_file + ".txt"))
                {
                    File.Delete(pdf_file + ".txt");
                    pdf_text = pdf_to_text(pdf_file);
                    File.AppendAllText(pdf_file + ".txt", pdf_text);
                }
                else
                {
                    pdf_text = File.ReadAllText(pdf_file + ".txt");
                }

                string simple_text = pdf_text.Replace(" ", "");
                if (pdf_text == "" || simple_text.Length < 512)
                {
                    Program.g_db.insert_statement_files_to_db(pdf_file, ConstEnv.STATEMENT_FILE_PARSE_IMAGE);
                    MyLogger.Info($"Can not get text from the Bank Statement File : {pdf_file}");
                    return true;
                }

                if (!is_valid_pdf(pdf_text))
                {
                    return false;
                }

                if (this.GetType() == typeof(BSParserCOB))
                {
                    string temp = System.IO.Path.GetFileNameWithoutExtension(pdf_file);
                    if (temp.IndexOf(" ") == -1)
                    {
                        if (temp.IndexOf("_") == -1)
                            throw new Exception("Invalid COB statement file name format. No blank to find the account.");
                        temp = temp.Substring(temp.LastIndexOf("_") + 1).Trim();
                    }
                    else
                    {
                        temp = temp.Substring(temp.LastIndexOf(" ") + 1).Trim();
                    }
                    if (temp.Length != 4)
                        throw new Exception("Invalid COB statement file name format. Too long account.");
                    ((BSParserCOB)this).set_account(temp);
                }

                parse_pdf(pdf_text);

                int pdf_file_id = Program.g_db.insert_statement_files_to_db(pdf_file, ConstEnv.STATEMENT_FILE_PARSE_SUCCEED);

                MyLogger.Info($"Parseing Bank statement OK : id = {pdf_file_id}, File = {pdf_file}");

                foreach (var d in transactions)
                {
                    List<BankTransactions> t_list = d.Value;

                    MyLogger.Info($"Account : {d.Key}, Transaction Count : {t_list.Count}");

                    foreach (var t in t_list)
                    {
                        Program.g_db.insert_bank_transaction(pdf_file_id, get_bank_name(), d.Key, t.t_type.ToString(), t.date, t.description, t.amount, t.number);

                        MyLogger.Info($"... Date        : {t.date}");
                        if (t.t_type == BankTransactions.BankTransactionType.Deposit_and_Additions)
                            MyLogger.Info("... Type        : DEPOSITS AND ADDITIONS");
                        else if (t.t_type == BankTransactions.BankTransactionType.Check_Paid)
                            MyLogger.Info("... Type        : CHECKS PAID");
                        else if (t.t_type == BankTransactions.BankTransactionType.Withdrawals_and_Debits)
                            MyLogger.Info("... Type        : ELECTRONIC WITHDRAWALS");
                        else if (t.t_type == BankTransactions.BankTransactionType.Fees)
                            MyLogger.Info($"... Type       : FEES");
                        else
                            MyLogger.Info($"... Type       : Invalid");

                        string[] desc_lines = t.description.Split('\n');
                        for (int i = 0; i < desc_lines.Length; i++)
                        {
                            if (i == 0)
                                MyLogger.Info($"... Description : {desc_lines[i]}");
                            else
                                MyLogger.Info($"...             : {desc_lines[i]}");
                        }

                        MyLogger.Info($"... Amount      : {t.amount}");
                        MyLogger.Info($"... Card No     : {t.number}");
                        MyLogger.Info("");
                    }
                }

            }
            //catch (Exception exception)
            //{
            //    Program.g_db.insert_statement_files_to_db(pdf_file, ConstEnv.STATEMENT_FILE_PARSE_FAILED);
            //
            //    MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            //    return false;
            //}
            return true;
        }
        protected virtual bool is_valid_pdf(string pdf_text)
        {
            return false;
        }
        protected string pdf_to_text(string fileName)
        {
            var text = new StringBuilder();
            if (File.Exists(fileName))
            {
                var pdfReader = new PdfReader(fileName);
                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    var strategy = new myLocationTextExtractionStrategy();
                    //strategy.DUMP_STATE = true;
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            string all_text = text.ToString();
            all_text = all_text.Replace("\r\n", "\n");
            string[] lines = all_text.Split('\n');
            string ret_text = "";
            foreach (string s in lines)
            {
                if (s.Trim() != "")
                {
                    ret_text += (ret_text == "") ? s : "\n" + s;
                }
            }
            return ret_text;
        }
        protected virtual void parse_pdf(string pdf_text)
        {
        }
        protected bool is_same_row(int pos_1, int pos_2)
        {
            int diff = pos_1 - pos_2;
            if (Math.Abs(diff) < 2)
                return true;
            return false;
        }
        protected bool is_same_row(int pos, BSTableHdrLasyout layout)
        {
            if (layout.start != -1 && layout.end != -1)
            {
                return (layout.start <= pos && pos <= layout.end);
            }
            if (layout.start != -1)
            {
                return (pos >= layout.start - 1);
            }
            return false;
        }
        protected int next_non_space(string src, int start_pos)
        {
            if (start_pos == src.Length)
                return -1;
            int i = start_pos;
            while (src[i] == ' ')
                i++;
            return i;
        }
        protected string get_next_data(string src)
        {
            if (src.Length <= 2)
                return src;

            string dst = "";
            for (int i = 0; i < src.Length - 2; i++)
            {
                if (src[i] == ' ' && src[i + 1] == ' '/* && src[i + 2] == ' ' && src[i + 3] == ' ' && src[i + 4] == ' '*/)
                    return dst;
                dst += src[i];
            }
            return src.Trim();
        }
        private void print_transactions()
        {
            foreach (var d in transactions)
            {
                MyLogger.Info($"Account : {d.Key}");

                List<BankTransactions> t_list = d.Value;

                foreach (var t in t_list)
                {
                    if (t.t_type == BankTransactions.BankTransactionType.Deposit_and_Additions)
                    {
                        MyLogger.Info($"... DEPOSITS AND ADDITIONS : Date = {t.date}, description = {t.description}, amount = {t.amount:N2}");
                    }
                    else if (t.t_type == BankTransactions.BankTransactionType.Check_Paid)
                    {
                        MyLogger.Info($"... CHECKS PAID : CheckNum = {t.number}, description = {t.description}, Date = {t.date}, amount = {t.amount:N2}");
                    }
                    else if (t.t_type == BankTransactions.BankTransactionType.Withdrawals_and_Debits)
                    {
                        MyLogger.Info($"... ELECTRONIC WITHDRAWALS : Date = {t.date}, description = {t.description}, amount = {t.amount:N2}");
                    }
                    else if (t.t_type == BankTransactions.BankTransactionType.Fees)
                    {
                        MyLogger.Info($"... FEES : Date = {t.date}, description = {t.description}, amount = {t.amount:N2}");
                    }
                    else if (t.t_type == BankTransactions.BankTransactionType.Invalid_Type)
                    {
                        MyLogger.Error($"... Invalid Type : Date = {t.date}, description = {t.description}, amount = {t.amount:N2}");
                    }
                }
            }
        }
    }
    public class BSTableHdrLasyout
    {
        public int start;
        public int end;
        public BSTableHdrLasyout()
        {
            start = -1;
            end = -1;
        }
        public void init()
        {
            start = -1;
            end = -1;
        }
    }
}
