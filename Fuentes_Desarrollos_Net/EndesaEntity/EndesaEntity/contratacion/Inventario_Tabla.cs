using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class Inventario_Tabla
    {
        // public int id_inventario { get; set; }
        public string cups22 { get; set; }        
        public string estado { get; set; }  
        public DateTime fecha_alta { get; set; }
        public DateTime fecha_baja { get; set; }
        public string  razon_social { get; set; }
        public string nif { get; set; }
        public string cod_provincia { get; set; }
        public string descripcion_poblacion { get; set; }
        public double[] potencias { get; set; }
        public bool vigente { get; set; }
        public bool tiene_incidencia { get; set; }
        public bool temporal { get; set; }
        public bool carterizado { get; set; }
        public DateTime informado_alta { get; set; }
        public Inventario_Tabla()
        {
            potencias = new double[6];            
            fecha_baja = DateTime.MinValue;
            vigente = false;
            temporal = false;
            tiene_incidencia = false;
            carterizado = false;
        }
              
        public string cups20()
        {
            return cups22.Substring(0, 20);
        }
    }
}
