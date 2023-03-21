using MailParser;
using Logger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace StatementHelper
{
    public class BSParserChase : BSParser
    {
        public BSParserChase() : base(ConstEnv.BANK_NAME_CHASE)
        {

        }
        protected override bool is_valid_pdf(string pdf_text)
        {
            string[] lines = pdf_text.Split('\n');
            if (lines.Length < 5)
                return false;
            for (int i = 0; i < 5; i++)
            {
                string line = lines[i].Trim();

                if (line.StartsWith("JPMorgan Chase Bank, N.A."))
                    return true;
            }
            if (pdf_text.IndexOf("JPMorgan Chase Bank, N.A.") != -1)
                return true;
            return false;
        }
        protected override void parse_pdf(string pdf_text)
        {
            bool get_date_period = false;
            string account = "";
            string old_account = "";
            BankTransactions old_transactions = null;

            int year = -1;

            string[] lines = pdf_text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                if (!get_date_period && line.IndexOf("through") != -1)
                {
                    string temp = line.Trim();
                    string from = temp.Substring(0, temp.IndexOf("through")).Trim();
                    string to = temp.Substring(temp.IndexOf("through") + "through".Length).Trim();
                    DateTime from_date;
                    DateTime to_date;
                    if (DateTime.TryParse(from, out from_date) && DateTime.TryParse(to, out to_date))
                    {
                        year = from_date.Year;

                        MyLogger.Info($"from {from_date} to {to_date}");
                        get_date_period = true;
                        continue;
                    }
                }
                if (line.Trim().IndexOf("Account Number:") != -1)
                {
                    line = line.Trim();
                    string temp = line.Substring(line.IndexOf("Account Number:") + "Account number:".Length).Trim();
                    account = temp;
                    if (old_account != "" && old_account == account)
                        continue;
                    transactions.Add(account, new List<BankTransactions>());
                    old_account = account;
                    MyLogger.Info($"Account {account}");
                    continue;
                }

                if (account != "" && line.Trim() == "DEPOSITS AND ADDITIONS")
                {
                    BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_amount = new BSTableHdrLasyout();

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

                        if (line.Trim().StartsWith("Total Deposits and Additions"))
                            break;
                        if (line.Trim() == "*end*deposits and additions")
                        {
                            page_gap = true;
                            continue;
                        }
                        if (line.Trim() == "DEPOSITS AND ADDITIONS")
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
                            start_pos = 0;
                            pos = next_non_space(line, start_pos);
                            data = line.Substring(pos);
                            data = get_next_data(data);
                            if (!DateTime.TryParse($"{data}/{year}", out date))
                            {
                                continue;
                            }

                            pos_date.init();
                            pos_description.init();
                            pos_amount.init();

                            pos_date.start = pos;

                            start_pos = pos + data.Length;
                            pos = next_non_space(line, start_pos);

                            pos_description.start = pos;
                            pos_date.end = pos - 1;

                            pos = line.Length - 12; // max len of("-$###,###.##")
                            pos_amount.start = pos;
                            pos_amount.end = line.Length + 2;
                            pos_description.end = pos - 1;

                            get_layout = false;
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
                        data = line.Substring(pos_description.start, pos_description.end - pos_description.start).Trim();
                        string description = data;

                        data = line.Substring(pos_amount.start).Trim();
                        string amount = data;

                        old_transactions = new BankTransactions();
                        old_transactions.t_type = BankTransactions.BankTransactionType.Deposit_and_Additions;
                        old_transactions.date = date;
                        old_transactions.description = description;

                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        transactions[account].Add(old_transactions);

                        //MyLogger.Info($"... DEPOSITS AND ADDITIONS : Date = {date}, description = {description}, amount = {amount}");
                    }
                    i = k - 1;
                    continue;
                }
                if (account != "" && line.Trim() == "CHECKS PAID")
                {
                    BSTableHdrLasyout pos_num = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_amount = new BSTableHdrLasyout();

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

                        if (line.Trim().StartsWith("Total Checks Paid"))
                            break;
                        if (line.Trim().StartsWith("*end*checks paid"))
                        {
                            page_gap = true;
                            continue;
                        }
                        if (line.Trim() == "CHECKS PAID")
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
                            pos_num.init();
                            pos_date.init();
                            pos_description.init();
                            pos_amount.init();

                            start_pos = 0;
                            pos = next_non_space(line, start_pos);

                            pos_num.start = pos;
                            pos_num.end = pos_num.start + 20;

                            pos = line.Length - 12; // max len of("-$###,###.##")
                            pos_amount.start = pos;
                            pos_amount.end = line.Length + 2;

                            int j = pos - 1;
                            while (line[j] == ' ')
                                j--;
                            while (line[j] != ' ')
                                j--;
                            pos_date.start = j;
                            pos_date.end = pos - 1;

                            pos_description.start = pos_num.end + 1;
                            pos_description.end = j - 1;

                            get_layout = false;
                        }

                        data = line.Substring(pos_num.start, pos_num.end - pos_num.start).Trim();
                        data = data.Replace("^", "");
                        data = data.Replace("*", "");
                        data = data.Trim();
                        string check_num = data;

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

                        data = line.Substring(pos_description.start, pos_description.end - pos_description.start).Trim();
                        string description = data;

                        data = line.Substring(pos_amount.start).Trim();
                        string amount = data;

                        old_transactions = new BankTransactions();
                        old_transactions.t_type = BankTransactions.BankTransactionType.Check_Paid;
                        old_transactions.date = date;
                        old_transactions.description = description;

                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        old_transactions.number = check_num;

                        transactions[account].Add(old_transactions);

                        //MyLogger.Info($"... CHECKS PAID : CheckNum = {check_num}, description = {description}, Date = {date}, amount = {amount}");
                    }
                    i = k - 1;
                    continue;
                }
                if (account != "" && line.Trim() == "ELECTRONIC WITHDRAWALS")
                {
                    BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_amount = new BSTableHdrLasyout();

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

                        if (line.Trim().StartsWith("Total Electronic Withdrawals"))
                            break;
                        if (line.Trim() == "*end*electronic withdrawal")
                        {
                            page_gap = true;
                            continue;
                        }
                        if (line.Trim() == "ELECTRONIC WITHDRAWALS")
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
                            start_pos = 0;
                            pos = next_non_space(line, start_pos);
                            data = line.Substring(pos);
                            data = get_next_data(data);
                            if (!DateTime.TryParse($"{data}/{year}", out date))
                            {
                                continue;
                            }

                            pos_date.init();
                            pos_description.init();
                            pos_amount.init();

                            pos_date.start = pos;

                            start_pos = pos + data.Length;
                            pos = next_non_space(line, start_pos);

                            pos_description.start = pos;
                            pos_date.end = pos - 1;

                            pos = line.Length - 12; // max len of("-$###,###.##")
                            pos_amount.start = pos;
                            pos_amount.end = line.Length + 2;
                            pos_description.end = pos - 1;

                            get_layout = false;
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
                        data = line.Substring(pos_description.start, pos_description.end - pos_description.start).Trim();
                        string description = data;

                        data = line.Substring(pos_amount.start).Trim();
                        string amount = data;

                        old_transactions = new BankTransactions();
                        old_transactions.t_type = BankTransactions.BankTransactionType.Withdrawals_and_Debits;
                        old_transactions.date = date;
                        old_transactions.description = description;

                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        transactions[account].Add(old_transactions);

                        //MyLogger.Info($"... ELECTRONIC WITHDRAWALS : Date = {date}, description = {description}, amount = {amount}");
                    }
                    i = k - 1;
                    continue;
                }
                if (account != "" && line.Trim() == "FEES")
                {
                    BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_amount = new BSTableHdrLasyout();

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

                        if (line.Trim().StartsWith("Total Fees"))
                            break;
                        if (line.Trim() == "*end*fees" || line.Trim() == "*end*fee")
                        {
                            page_gap = true;
                            continue;
                        }
                        if (line.Trim() == "FEES")
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
                            start_pos = 0;
                            pos = next_non_space(line, start_pos);
                            data = line.Substring(pos);
                            data = get_next_data(data);
                            if (!DateTime.TryParse($"{data}/{year}", out date))
                            {
                                continue;
                            }

                            pos_date.init();
                            pos_description.init();
                            pos_amount.init();

                            pos_date.start = pos;

                            start_pos = pos + data.Length;
                            pos = next_non_space(line, start_pos);

                            pos_description.start = pos;
                            pos_date.end = pos - 1;

                            pos = line.Length - 12; // max len of("-$###,###.##")
                            pos_amount.start = pos;
                            pos_amount.end = line.Length + 2;
                            pos_description.end = pos - 1;

                            get_layout = false;
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
                        data = line.Substring(pos_description.start, pos_description.end - pos_description.start).Trim();
                        string description = data;

                        data = line.Substring(pos_amount.start).Trim();
                        string amount = data;

                        old_transactions = new BankTransactions();
                        old_transactions.t_type = BankTransactions.BankTransactionType.Fees;
                        old_transactions.date = date;
                        old_transactions.description = description;

                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        transactions[account].Add(old_transactions);

                        //MyLogger.Info($"... FEES : Date = {date}, description = {description}, amount = {amount}");
                    }
                    i = k - 1;
                    continue;
                }
                if (account != "" && line.Trim() == "TRANSACTION DETAIL")
                {
                    BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_amount = new BSTableHdrLasyout();
                    BSTableHdrLasyout pos_balance = new BSTableHdrLasyout();

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

                        if (line.Trim().StartsWith("Ending Balance"))
                            break;
                        if (line.Trim().StartsWith("Page "))
                        {
                            page_gap = true;
                            continue;
                        }
                        if (line.Trim() == "TRANSACTION DETAIL")
                        {
                            if (page_gap)
                                page_gap = false;
                            get_layout = true;

                            if (lines[k + 1].Trim().StartsWith("Beginning Balance"))
                                k++;
                            continue;
                        }
                        if (page_gap)
                        {
                            continue;
                        }
                        if (get_layout)
                        {
                            start_pos = 0;
                            pos = next_non_space(line, start_pos);
                            data = line.Substring(pos);
                            data = get_next_data(data);
                            if (!DateTime.TryParse($"{data}/{year}", out date))
                            {
                                continue;
                            }

                            pos_date.init();
                            pos_description.init();
                            pos_amount.init();

                            pos_date.start = pos;

                            start_pos = pos + data.Length;
                            pos = next_non_space(line, start_pos);

                            pos_description.start = pos;
                            pos_date.end = pos - 1;

                            pos = line.Length - 12; // max len of("-$###,###.##")
                            pos_balance.start = pos;
                            pos_balance.end = line.Length + 2;

                            int x = pos_balance.start - 1;
                            while (line[x] == ' ')
                                x--;

                            pos_amount.end = pos_balance.start - 1;
                            pos_amount.start = x - 12; // max len of("-$###,###.##")
                            pos_description.end = pos_amount.start - 1;

                            get_layout = false;
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
                        data = line.Substring(pos_description.start, pos_description.end - pos_description.start).Trim();
                        string description = data;

                        data = line.Substring(pos_amount.start).Trim();
                        string amount = data;

                        old_transactions = new BankTransactions();
                        old_transactions.date = date;
                        old_transactions.description = description;

                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        if (description.IndexOf("Online Transfer From") != -1 || description.IndexOf("Remote Online Deposit") != -1 || description.IndexOf("Deposit") != -1)
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Deposit_and_Additions;
                        }
                        else if (description.IndexOf("Online Transfer To") != -1)
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Withdrawals_and_Debits;
                        }
                        else if (description.StartsWith("Check") && description.Trim().Substring("Check".Length).Trim()[0] == '#')
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Check_Paid;
                            old_transactions.number = description.Trim().Substring("Check".Length).Trim().Substring(1);
                        }
                        else
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Invalid_Type;
                        }

                        transactions[account].Add(old_transactions);

                        //MyLogger.Info($"... FEES : Date = {date}, description = {description}, amount = {amount}");
                    }
                    i = k - 1;
                    continue;
                }
            }
        }
    }
}
