using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class Extractores_Jesus
    {
        utilidades.Param p;
        utilidades.Fechas f;
        logs.Log ficheroLog;
        utilidades.Seguimiento_Procesos ss_pp;
        public Extractores_Jesus()
        {
            p = new utilidades.Param("medida_param", servidores.MySQLDB.Esquemas.MED);
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_Extractores_Jesus");
            f = new utilidades.Fechas();
            ss_pp = new utilidades.Seguimiento_Procesos();
        }


        public void Extractor_bat_Diario()
        {
            ss_pp.Update_Fecha_Inicio("Medida", "BAT DIARIO", "1_1_Ejecución Extractor");
            ficheroLog.Add("Ejecutando extractor_bat_diario");
            utilidades.Fichero.EjecutaComando(p.GetValue("extractor_bat_diario"),f.UltimoDiaHabil().ToString("yyMMdd"));
            ficheroLog.Add("Fin extractor_bat_diario");
            ss_pp.Update_Fecha_Fin("Medida", "BAT DIARIO", "1_1_Ejecución Extractor");
        }

        public void Extractor_Comentarios_SCE()
        {
            ficheroLog.Add("Ejecutando extractor_comentarios");
            utilidades.Fichero.EjecutaComando(p.GetValue("extractor_comentarios"));
            ficheroLog.Add("Fin extractor_comentarios");
        }

        public void Extractor_Peajes()
        {
            ficheroLog.Add("Ejecutando extractor_peajes_1");
            utilidades.Fichero.EjecutaComando(p.GetValue("extractor_peajes_1"));
            ficheroLog.Add("Fin extractor_peajes_1");
            ficheroLog.Add("Ejecutando extractor_peajes_2");
            utilidades.Fichero.EjecutaComando(p.GetValue("extractor_peajes_2"));
            ficheroLog.Add("Fin extractor_peajes_2");
        }

    }
}
