using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Table_ag_comppd : FieldsAudit    
    {
        public int id_componente { get; set; }
        public string liquidacion { get; set; }
        public DateTime fecha { get; set; }
        public string unidad { get; set; }
        public double[] value { get; set; }

        public Table_ag_comppd()
        {
            value = new double[25];
        }
    }
}
