using MailParser;
using Logger;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace StatementHelper
{
    public class BSParserWF : BSParser
    {
        public BSParserWF() :base(ConstEnv.BANK_NAME_WELLSFARGO)
        {

        }
        protected override bool is_valid_pdf(string pdf_text)
        {
            string[] lines = pdf_text.Split('\n');
            if (lines.Length < 3)
                return false;
            for (int i = 0; i < 3; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("Wells Fargo"))
                    return true;
            }
            if (pdf_text.IndexOf("Online:  wellsfargo.com/biz") != -1)
                return true;
            return false;
        }
        protected override void parse_pdf(string pdf_text)
        {
            bool get_date_period = false;
            bool transaction_start = false;
            bool is_header_line = false;
            string account = "";
            BankTransactions old_transactions = null;

            BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_check = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_deposits_credits = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_withdrawals_debits = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_ending_daily_balabce = new BSTableHdrLasyout();

            int year = -1;

            string[] lines = pdf_text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                if (!get_date_period && (line.Trim().StartsWith("Primary account number:") || line.Trim().StartsWith("Account number:")) && line.IndexOf("¦") != -1)
                {
                    string temp = line.Substring(line.IndexOf("¦") + 1).Trim();
                    if (temp.IndexOf("¦") == -1)
                        continue;
                    temp = temp.Substring(0, temp.IndexOf("¦")).Trim();
                    if (temp.IndexOf("-") == -1)
                        continue;

                    string from = temp.Substring(0, temp.IndexOf("-")).Trim();
                    string to = temp.Substring(temp.IndexOf("-") + 1).Trim();
                    DateTime from_date = DateTime.Parse(from);
                    DateTime to_date = DateTime.Parse(to);

                    year = from_date.Year;

                    //MyLogger.Log($"from {from_date} to {to_date}");
                    get_date_period = true;
                    continue;
                }
                if (!get_date_period && line.IndexOf("¦   Page ") != -1)
                {
                    string temp = line.Substring(0, line.IndexOf("¦   Page ")).Trim();
                    if (temp.IndexOf("¦") != -1)
                    {
                        temp = temp.Substring(temp.LastIndexOf("¦") + 1).Trim();
                    }
                    if (temp.IndexOf("-") == -1)
                    {
                        DateTime date;
                        if (DateTime.TryParse(temp, out date))
                        {
                            year = date.Year;

                            //MyLogger.Log($"Date {date}");
                            get_date_period = true;
                        }
                    }
                    else
                    {
                        string from = temp.Substring(0, temp.IndexOf("-")).Trim();
                        string to = temp.Substring(temp.IndexOf("-") + 1).Trim();
                        DateTime from_date = DateTime.Parse(from);
                        DateTime to_date = DateTime.Parse(to);

                        year = from_date.Year;

                        //MyLogger.Log($"from {from_date} to {to_date}");
                        get_date_period = true;
                    }
                    continue;
                }
                if (line.Trim().StartsWith("Account number:") && lines[i + 1].Trim() == "Activity summary")
                {
                    string temp = line.Trim().Substring("Account number:".Length).Trim();
                    account = temp;
                    transactions.Add(account, new List<BankTransactions>());
                    //MyLogger.Log($"Account {account}");
                }

                if (account != "" && line.Trim().StartsWith("Transaction history"))
                {
                    pos_date.init();
                    pos_check.init();
                    pos_description.init();
                    pos_deposits_credits.init();
                    pos_withdrawals_debits.init();
                    pos_ending_daily_balabce.init();

                    bool get_layout = false;
                    bool page_gap = false;
                    int start_pos;
                    int pos;
                    string data;
                    DateTime date;
                    int k = i;
                    for (k = i; ; k++)
                    {
                        line = lines[k];

                        if (line.Trim().StartsWith("Ending balance"))
                            break;
                        if (line.Trim().StartsWith("Page ") || (line.Trim().StartsWith("Account number:") && line.IndexOf("Page ") != -1))
                        {
                            page_gap = true;
                            continue;
                        }
                        if (line.Trim().StartsWith("Transaction history"))
                        {
                            if (page_gap)
                                page_gap = false;
                            get_layout = true;
                            continue;
                        }
                        if (page_gap)
                        {
                            continue;
                        }
                        if (get_layout)
                        {
                            if (line.IndexOf("Check") == -1 && line.IndexOf("Deposits/") == -1 && line.IndexOf("Withdrawals/") == -1 && line.IndexOf("Ending daily") == -1)
                                continue;

                            pos_date.init();
                            pos_check.init();
                            pos_description.init();
                            pos_deposits_credits.init();
                            pos_withdrawals_debits.init();
                            pos_ending_daily_balabce.init();

                            if (line.IndexOf("Check") != -1)
                                pos_check.start = line.IndexOf("Check");
                            if (line.IndexOf("Deposits/") != -1)
                                pos_deposits_credits.start = line.IndexOf("Deposits/");
                            if (line.IndexOf("Withdrawals/") != -1)
                                pos_withdrawals_debits.start = line.IndexOf("Withdrawals/");
                            if (line.IndexOf("Ending daily") != -1)
                                pos_ending_daily_balabce.start = line.IndexOf("Ending daily");

                            line = lines[++k];

                            if (line.IndexOf("Date") != -1)
                                pos_date.start = line.IndexOf("Date");
                            if (line.IndexOf("Number") != -1)
                                pos_check.start = Math.Min(pos_check.start, line.IndexOf("Number"));
                            if (line.IndexOf("Description") != -1)
                                pos_description.start = line.IndexOf("Description");
                            if (line.IndexOf("Credits") != -1)
                                pos_deposits_credits.start = Math.Min(pos_deposits_credits.start, line.IndexOf("Credits"));
                            if (line.IndexOf("Debits") != -1)
                                pos_withdrawals_debits.start = Math.Min(pos_withdrawals_debits.start, line.IndexOf("Debits"));
                            if (line.IndexOf("balance") != -1)
                                pos_ending_daily_balabce.start = Math.Min(pos_ending_daily_balabce.start, line.IndexOf("balance"));

                            if (pos_date.start != -1)
                            {
                                if (pos_check.start != -1)
                                    pos_date.end = pos_check.start - 1;
                                else if (pos_description.start != -1)
                                    pos_date.end = pos_description.start - 1;
                            }
                            if (pos_check.start != -1)
                            {
                                if (pos_description.start != -1)
                                    pos_check.end = pos_description.start - 1;
                            }
                            if (pos_description.start != -1)
                            {
                                if (pos_deposits_credits.start != -1)
                                    pos_description.end = pos_deposits_credits.start - 1;
                            }
                            if (pos_deposits_credits.start != -1)
                            {
                                if (pos_withdrawals_debits.start != -1)
                                    pos_deposits_credits.end = pos_withdrawals_debits.start - 1;
                            }
                            if (pos_withdrawals_debits.start != -1)
                            {
                                if (pos_ending_daily_balabce.start != -1)
                                    pos_withdrawals_debits.end = pos_ending_daily_balabce.start - 1;
                            }

                            get_layout = false;
                            continue;
                        }

                        data = line.Substring(pos_date.start, pos_date.end - pos_date.start).Trim();
                        if (!DateTime.TryParse($"{data}/{year}", out date))
                        {
                            if (old_transactions != null)
                            {
                                data = line.Trim();
                                old_transactions.description += "\n" + data;
                            }
                            continue;
                        }
                        if (pos_check.start == -1)
                            data = "";
                        else
                            data = line.Substring(pos_check.start, pos_check.end - pos_check.start).Trim();
                        string check_num = data;

                        data = line.Substring(pos_description.start, pos_description.end - pos_description.start).Trim();
                        string description = data;

                        if (line.Length <= pos_deposits_credits.end)
                        {
                            data = line.Substring(pos_deposits_credits.start).Trim();
                        }
                        else
                        {
                            data = line.Substring(pos_deposits_credits.start, pos_deposits_credits.end - pos_deposits_credits.start).Trim();
                        }
                        string deposit = data;

                        if (line.Length <= pos_withdrawals_debits.end)
                        {
                            if (line.Length > pos_withdrawals_debits.start)
                                data = line.Substring(pos_withdrawals_debits.start).Trim();
                            else
                                data = "";
                        }
                        else
                        {
                            data = line.Substring(pos_withdrawals_debits.start, pos_withdrawals_debits.end - pos_withdrawals_debits.start).Trim();
                        }
                        string withdrawals = data;

                        old_transactions = new BankTransactions();
                        old_transactions.date = date;
                        old_transactions.description = description;

                        string amount = "";
                        if (check_num != "")
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Check_Paid;
                            amount = withdrawals;
                        }
                        else if (withdrawals == "")
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Deposit_and_Additions;
                            amount = deposit;
                        }
                        else
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Withdrawals_and_Debits;
                            amount = withdrawals;
                        }
                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        transactions[account].Add(old_transactions);
                    }
                    i = k - 1;
                    continue;
                }
                continue;
            }
        }
    }
}
