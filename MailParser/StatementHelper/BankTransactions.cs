using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatementHelper
{
    public class BankTransactions
    {
        public enum BankTransactionType
        {
            Invalid_Type,
            Deposit_and_Additions, // or deposits and credits
            Check_Paid,
            Withdrawals_and_Debits,
            Fees,
        }

        public BankTransactionType t_type;
        public DateTime date;
        public string description;
        public float amount;
        public string number;

        public BankTransactions()
        {
            t_type = BankTransactionType.Invalid_Type;
            date = DateTime.MinValue;
            description = "";
            amount = 0;
            number = "";
        }
    }
}
