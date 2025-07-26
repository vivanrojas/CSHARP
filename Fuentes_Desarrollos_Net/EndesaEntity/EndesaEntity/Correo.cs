using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Correo
    {
        public string sender { get; set; }
        public int mail_id { get; set; }
        public string name { get; set; }
        public string mailbox { get; set; }
        public DateTime begin_date { get; set; }
        public DateTime end_date { get; set; }
        public string description { get; set; }
        public string subject { get; set; }
        public string body { get; set; }
        public DateTime receivedTime { get; set; }
    }
}
