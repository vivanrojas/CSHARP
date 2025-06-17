using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class Table_ag_compp : FieldsAudit
    {
        public int id_componente { get; set; }
        public string descripcion_breve { get; set; }
        public string descripcion { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public string fichero { get; set; }
        public int columna { get; set; }
        public string unidad { get; set; }
        public int userid { get; set; }
        public bool tiene_liquidacion { get; set; }
        public bool generar_con_ceros { get; set; }
        public int num_columnas { get; set; }
        public int id_tipo_componente { get; set; }
        public int fila_inicio { get; set; }
    }
}
