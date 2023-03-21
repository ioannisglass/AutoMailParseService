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
    public class BSParserCOB : BSParser
    {
        /**
         * The account number is  displayed as an image in the COB statement PDF,
         * so we must indicate the account MANUALLY.
         **/
        private string account;
        public BSParserCOB() : base(ConstEnv.BANK_NAME_COB)
        {
            account = "";
        }
        public void set_account(string _account)
        {
            if (transactions.ContainsKey(_account))
                return;
            account = _account;
            transactions.Add(account, new List<BankTransactions>());
        }
        protected override bool is_valid_pdf(string pdf_text)
        {
            if (pdf_text.IndexOf("www.capitalone.com/nohasslerewards") != -1)
                return true;
            if (pdf_text.IndexOf("Capital One, All rights reserved.") != -1)
                return true;
            if (pdf_text.IndexOf("Capital One. All rights reserved.") != -1)
                return true;
            return false;
        }
        protected override void parse_pdf(string pdf_text)
        {
            if (account == "")
                throw new Exception("No account. COB SHOULD get the account before parsing.");

            bool get_date_period = false;
            BankTransactions old_transactions = null;

            BSTableHdrLasyout pos_date = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_amount = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_balance = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_tr_type = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_description = new BSTableHdrLasyout();
            BSTableHdrLasyout pos_card_no = new BSTableHdrLasyout();

            int year = -1;

            string[] lines = pdf_text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                if (!get_date_period && line.IndexOf("FOR PERIOD") != -1 && line.IndexOf("-") != -1)
                {
                    string temp = line.Substring(line.IndexOf("FOR PERIOD") + "FOR PERIOD".Length).Trim();
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

                if (line.Replace(" ", "") == "DateAmountResultingBalanceTransactionTypeDescriptionDebitCardNo."
                    || (line.Replace(" ", "") == "DebitCard" && lines[i + 1].Replace(" ", "") == "DateAmountResultingBalanceTransactionTypeDescription"))
                {
                    string data;
                    DateTime date;
                    int k = i;
                    for (k = i; ; k++)
                    {
                        line = lines[k];

                        if (line.Trim().StartsWith("Thank you for banking with us"))
                            break;
                        if (line.Trim().StartsWith("PAGE "))
                            break;
                        if (line.IndexOf("* designates gap in check sequence") != -1)
                            break;
                        if (line.Trim() == ".....")
                            break;
                        if (line.Trim() == "TWO WAYS TO AVOID A MONTHLY SERVICE CHARGE")
                            break;
                        if (line.Replace(" ", "") == "DateAmountResultingBalanceTransactionTypeDescriptionDebitCardNo."
                            || (line.Replace(" ", "") == "DebitCard" && lines[k + 1].Replace(" ", "") == "DateAmountResultingBalanceTransactionTypeDescription"))
                        {
                            pos_date.init();
                            pos_amount.init();
                            pos_balance.init();
                            pos_tr_type.init();
                            pos_description.init();
                            pos_card_no.init();

                            if (line.Replace(" ", "") == "DateAmountResultingBalanceTransactionTypeDescriptionDebitCardNo.")
                            {
                                pos_date.start = line.IndexOf("Date");
                                pos_amount.start = line.IndexOf("Amount");
                                pos_balance.start = line.IndexOf("Resulting Balance");
                                pos_tr_type.start = line.IndexOf("Transaction Type");
                                pos_description.start = line.IndexOf("Description");
                                pos_card_no.start = line.IndexOf("Debit Card");

                                pos_date.end = pos_date.start + 6; // "##/##"
                                pos_amount.start = pos_date.end + 1;
                                pos_amount.end = pos_balance.start - 1;
                                pos_balance.end = pos_tr_type.start - 1;
                                pos_tr_type.end = pos_description.start - 1;
                                pos_description.end = pos_card_no.start - 1;
                                pos_card_no.end = pos_card_no.start + 4;
                            }
                            else
                            {
                                pos_date.start = lines[k + 1].IndexOf("Date");
                                pos_amount.start = lines[k + 1].IndexOf("Amount");
                                pos_balance.start = lines[k + 1].IndexOf("Resulting Balance");
                                pos_tr_type.start = lines[k + 1].IndexOf("Transaction Type");
                                pos_description.start = lines[k + 1].IndexOf("Description");
                                pos_card_no.start = line.IndexOf("Debit Card");

                                pos_date.end = pos_date.start + 6; // "##/##"
                                pos_amount.start = pos_date.end + 1;
                                pos_amount.end = pos_balance.start - 1;
                                pos_balance.end = pos_tr_type.start - 1;
                                pos_tr_type.end = pos_description.start - 1;
                                pos_description.end = pos_card_no.start - 1;
                                pos_card_no.end = pos_card_no.start + 4;
                                k++;
                            }
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

                        data = line.Substring(pos_amount.start, pos_amount.end - pos_amount.start).Trim();
                        string amount = data;

                        // balance
                        data = line.Substring(pos_balance.start, pos_balance.end - pos_balance.start).Trim();

                        data = line.Substring(pos_tr_type.start, pos_tr_type.end - pos_tr_type.start).Trim();
                        string tr_type = data;

                        data = line.Substring(pos_description.start, Math.Min(line.Length, pos_description.end) - pos_description.start).Trim();
                        string description = data;

                        string card_no = "";
                        if (line.Length > pos_card_no.start)
                        {
                            int j = line.Length - 1;
                            while (line[j] != ' ')
                                j--;
                            if (j < line.Length - 1)
                            {
                                string temp = line.Substring(j + 1);

                                if (temp.Length == 4)
                                {
                                    if (temp.Count(s => Char.IsDigit(s)) == temp.Length)
                                    {
                                        card_no = temp;
                                    }
                                }
                            }
                        }

                        old_transactions = new BankTransactions();
                        old_transactions.date = date;
                        old_transactions.description = description;
                        old_transactions.number = card_no;

                        amount = amount.Replace(" ", "");
                        amount = amount.Replace(",", "");
                        old_transactions.amount = Str_Utils.string_to_currency(amount);

                        if (tr_type == "Debit" || tr_type == "Check" || tr_type == "Service Charge")
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Withdrawals_and_Debits;
                        }
                        else if (tr_type == "Deposit" || tr_type == "Credit")
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Deposit_and_Additions;
                        }
                        else
                        {
                            old_transactions.t_type = BankTransactions.BankTransactionType.Invalid_Type;
                        }

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
