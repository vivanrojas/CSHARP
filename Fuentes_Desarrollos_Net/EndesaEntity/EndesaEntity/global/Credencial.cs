using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.global
{
    public class Credencial
    {
        public string user_id { get; set; }
        public string user_name { get; set; }
        public DateTime from_date { get; set; }
        public DateTime to_date { get; set; }
        public string mail { get; set; }
        public string server_user { get; set; }
        public string server_password { get; set; }
    }
}
