using MailParser;
using Logger;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class KReportCR7 : KReportCR
    {
        public KReportCR7(int mail_seq_num) : base()
        {
            if (mail_seq_num == 0)
                m_mail_type = MailType.CR_7;
            else if (mail_seq_num == 1)
                m_mail_type = MailType.CR_7_1;
            else if (mail_seq_num == 2)
                m_mail_type = MailType.CR_7_2;
            else if (mail_seq_num == 3)
                m_mail_type = MailType.CR_7_3;
            else if (mail_seq_num == 4)
                m_mail_type = MailType.CR_7_4;
            else if (mail_seq_num == 5)
                m_mail_type = MailType.CR_7_5;
            else
                throw new Exception($"Invalid CR_7 mail seq num. {mail_seq_num}");
        }
        public KReportCR7(MailType mail_type) : base()
        {
            m_mail_type = mail_type;
        }
    }
}
