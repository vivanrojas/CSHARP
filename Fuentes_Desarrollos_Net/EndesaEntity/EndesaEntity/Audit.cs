using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Audit
    {
        public string created_by { get; set; }
        public DateTime created_date { get; set; }
        public string last_update_by { get; set; }
        public DateTime last_update_date { get; set; }
        
        
    }
}
