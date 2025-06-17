using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class MedidaFunciones : EndesaEntity.Medida
    {
        public Dictionary<string, EndesaEntity.Medida> dicMedida { get; set; }

        EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes pm;
        public medida.CurvaResumenFunciones cr { get; set; }

        public MedidaFunciones(List<string> lista_puntos, DateTime fd, DateTime fh)
        {

            dicMedida = new Dictionary<string, EndesaEntity.Medida>();
            // Buscamos los puntos de medida
            pm = new EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes(lista_puntos);
            // Buscamos los resumenes de curva
            cr = new CurvaResumenFunciones(pm.lista_cups15, fd, fh);
            // Volcamos datos en Diccionario
            TrataCurvaResumen(lista_puntos, fd, fh);
        }

        private void CreaDiccionarioMedida(List<string> lista_puntos, DateTime fd, DateTime fh)
        {
            string key;
            EndesaEntity.Medida c;
            for (int i = 0; i < lista_puntos.Count; i++)
            {
                c = new EndesaEntity.Medida();
                key = lista_puntos[i] + fd.ToString("yyyyMMdd") + fh.ToString("yyyyMMdd");
                c.cups13 = lista_puntos[i];
                c.estado = "SIN CURVA";
                dicMedida.Add(key, c);
            }
        }

        private void TrataCurvaResumen(List<string> lista_puntos, DateTime fd, DateTime fh)
        {
            string key = "";

            foreach (KeyValuePair<string, EndesaEntity.medida.CurvaResumenTabla> p in cr.dic_curva_resumen)
            {
                EndesaEntity.Medida c;
                key = p.Value.cpuntmed.Substring(0, 13) + p.Value.fecha_desde.ToString("yyyyMMdd") + p.Value.fecha_hasta.ToString("yyyyMMdd");
                if (!dicMedida.TryGetValue(key, out c))
                {
                    EndesaEntity.Medida cc = new EndesaEntity.Medida();
                    cc.cups13 = p.Key.Substring(0, 13);
                    cc.fromdate = p.Value.fecha_desde;
                    cc.todate = p.Value.fecha_hasta;
                    cc.activa = p.Value.activa;
                    cc.reactiva = p.Value.reactiva;
                    cc.estado = p.Value.estado;
                    cc.fuente = p.Value.fuente;
                    cc.dias = p.Value.dias;
                    //cc.lista_curva_resumen.Add(p.Value);
                    dicMedida.Add(key, cc);
                }
                else
                {
                    c.activa = c.activa + p.Value.activa;
                    c.reactiva = c.reactiva + p.Value.reactiva;
                    // c.lista_curva_resumen.Add(p.Value);
                }
            }
        }
    }
}
