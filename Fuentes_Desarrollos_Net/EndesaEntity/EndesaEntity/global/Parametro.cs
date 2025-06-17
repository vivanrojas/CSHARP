using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.global
{
    public class Parametro
    {
        public EndesaEntity.global.MySQLDB.Esquemas esquema { get; set; }        
        public string tabla { get; set; }
        public string code { get; set; }
        public DateTime from_date { get; set; }
        public DateTime to_date { get; set; }
        public string value { get; set; }
        public string description { get; set; }
        public string created_by { get; set; }
        public DateTime created_date { get; set; }
        public string last_update_by { get; set; }
        public DateTime last_update_date { get; set; }
        public bool existe { get; set; }
        public string user { get; set; }
    }
}
