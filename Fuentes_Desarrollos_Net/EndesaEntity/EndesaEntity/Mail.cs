using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Mail
    {
        public int mail_id { get; set; }
        public string name { get; set; }
        public string mailbox { get; set; }
        public DateTime begin_date { get; set; }
        public DateTime end_date { get; set; }
        public string description { get; set; }
        public string subject { get; set; }
        public string body { get; set; }
        public string to { get; set; }
        public string cc { get; set; }
        public string bcc { get; set; }
    }
}
