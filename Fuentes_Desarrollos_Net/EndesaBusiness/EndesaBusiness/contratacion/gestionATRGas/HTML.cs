using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    class HTML
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ATRGas_HTML");
        utilidades.ParamUser p;
        public HTML()
        {
            p = new utilidades.ParamUser("atrgas_param_user", System.Environment.UserName, servidores.MySQLDB.Esquemas.CON);
        }

        public string GeneraCuerpoHTML(EndesaEntity.contratacion.gas.Solicitud sol)
        {
            string body = "";
            string linea = "";
            try
            {
                body = p.GetValue("html_body_head", DateTime.Now, DateTime.Now);
                for (int i = 0; i < sol.detalle.Count; i++)
                {
                    linea = p.GetValue("html_body_line", DateTime.Now, DateTime.Now);
                    linea = linea.Replace("CUPS", sol.cups);
                    linea = linea.Replace("TARIFA_SOLICITADA", sol.detalle[i].tarifa);
                    linea = linea.Replace("CAUDAL_SOLICITADA", sol.detalle[i].qd.ToString());
                    linea = linea.Replace("FECHA_INICIO_SOLICITADA", sol.detalle[i].fecha_inicio.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("FECHA_FIN_SOLICITADA", sol.detalle[i].fecha_fin.ToString("dd/MM/yyyy"));
                    linea = linea.Replace("PRODUCTO_SOLICITADO", sol.detalle[i].producto);
                    body += linea;
                }
                body += p.GetValue("html_body_foot", DateTime.Now, DateTime.Now);
                return body;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("GeneraCuerpoHTML: " + e.Message);
                return "";
            }
        }
    }
}
