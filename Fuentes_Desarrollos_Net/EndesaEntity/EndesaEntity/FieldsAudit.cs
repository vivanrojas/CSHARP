using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class FieldsAudit
    {
        public string created_by { get; set; }
        public DateTime creation_date { get; set; }
        public string last_update_by { get; set; }
        public DateTime last_update_date { get; set; }
    }
}
