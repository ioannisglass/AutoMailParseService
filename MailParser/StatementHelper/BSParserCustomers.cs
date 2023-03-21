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
    public class BSParserCustomers : BSParser
    {
        public BSParserCustomers() : base(ConstEnv.BANK_NAME_CUSTOMERS)
        {
        }
        protected override bool is_valid_pdf(string pdf_text)
        {
            if (pdf_text.IndexOf("www.customersbank.com") != -1)
                return true;
            return false;
        }
        protected override void parse_pdf(string pdf_text)
        {
            string account = "";

            BankTransactions old_transactions = null;

            BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_debits = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_credits = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_balance = new BSTableHdrLasyout();

            string[] lines = pdf_text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                if (line.Trim().StartsWith("COMMERCIAL INTEREST CHECKING") && lines[i + 1].Trim().StartsWith("Account Summary"))
                {
                    string temp = line.Trim().Substring("COMMERCIAL INTEREST CHECKING".Length).Trim();

                    int k = 0;
                    while (temp[k] == 'X')
                        k++;
                    temp = temp.Substring(k);
                    if (temp.Length == 4 && !transactions.ContainsKey(temp))
                    {
                        account = temp;
                        transactions.Add(account, new List<BankTransactions>());
                    }
                }

                if (account != "" && line.Replace(" ", "") == "DateDescriptionDebitsCreditsBalance")
                {
                    string data;
                    DateTime date;
                    int k = i;
                    for (k = i; ; k++)
                    {
                        line = lines[k];

                        if (line.IndexOf("Ending Balance") != -1)
                            break;
                        if (line.Trim().StartsWith("In Case of Errors or Questions about"))
                            break;
                        if (line.Replace(" ", "") == "DateDescriptionDebitsCreditsBalance")
                        {
                            pos_date.init();
                            pos_description.init();
                            pos_debits.init();
                            pos_credits.init();
                            pos_balance.init();

                            pos_date.start = line.IndexOf("Date");
                            pos_description.start = line.IndexOf("Description");
                            pos_debits.start = line.IndexOf("Debits");
                            pos_credits.start = line.IndexOf("Credits");
                            pos_balance.start = line.IndexOf("Balance");

                            pos_date.end = pos_date.start + 12; // "## /## /####"
                            pos_date.start -= 2;
                            pos_description.start = pos_date.end + 1;
                            pos_debits.end = pos_debits.start + "Debits".Length + 1;
                            pos_debits.start = pos_debits.end - 14; // max len of("-$###,###.##") + padding
                            pos_description.end = pos_debits.start - 1;
                            pos_credits.end = pos_credits.start + "Credits".Length + 1;
                            pos_credits.start = pos_credits.end - 14; // max len of("-$###,###.##") + padding
                            pos_balance.end = pos_balance.start + "Balance".Length + 1;
                            pos_balance.start = pos_balance.end - 14; // max len of("-$###,###.##") + padding

                            continue;
                        }

                        int start_pos = next_non_space(line, 0);
                        if (Math.Abs(start_pos - pos_date.start) > 10 && Math.Abs(start_pos - pos_description.start) > 10)
                        {
                            break;
                        }

                        if (Math.Abs(start_pos - pos_description.start) < 10)
                        {
                            if (old_transactions != null)
                            {
                                data = line.Trim();
                                old_transactions.description += "\n" + data;
                            }
                            continue;
                        }

                        data = line.Substring(pos_date.start, pos_date.end - pos_date.start).Trim();
                        if (!DateTime.TryParse(data, out date))
                        {
                            break;
                        }

                        data = line.Substring(pos_description.start, Math.Min(line.Length, pos_description.end) - pos_description.start).Trim();
                        string description = data;
                        if (description == "Beginning Balance")
                            continue;

                        data = line.Substring(pos_debits.start, pos_debits.end - pos_debits.start).Trim();
                        string debits = data;

                        data = line.Substring(pos_credits.start, pos_credits.end - pos_credits.start).Trim();
                        string credits = data;

                        // balance
                        data = line.Substring(pos_balance.start, Math.Min(line.Length, pos_balance.end) - pos_balance.start).Trim();


                        old_transactions = new BankTransactions();
                        old_transactions.date = date;
                        old_transactions.description = description;

                        string amount = "";
                        if (debits != "")
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Withdrawals_and_Debits;
                            amount = debits;
                        }
                        else if (credits != "")
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Deposit_and_Additions;
                            amount = credits;
                        }
                        else
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Invalid_Type;
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
