using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.xml
{
    class XML_SolicitudesCodigos : EndesaEntity.contratacion.Solicitudes_Codigos_Tabla
    {
        public Dictionary<string, EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> dic { get; set; }
        public XML_SolicitudesCodigos()
        {
            dic = new Dictionary<string, EndesaEntity.contratacion.Solicitudes_Codigos_Tabla>();
            CargaDatos();
        }

        private void CargaDatos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select CodigoDelProceso, CodigoDePaso, Descripcion from cont_xml_param_solicitudes_codigos";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Solicitudes_Codigos_Tabla c = new EndesaEntity.contratacion.Solicitudes_Codigos_Tabla();
                    c.descripcion = r["Descripcion"].ToString();
                    c.codigoproceso = r["CodigoDelProceso"].ToString();
                    c.codigopaso = r["CodigoDePaso"].ToString();
                    dic.Add(c.codigoproceso + c.codigopaso, c);
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Estados - CargaDatos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
        public string TipoProceso(string codigoproceso, string codigopaso)
        {
            EndesaEntity.contratacion.Solicitudes_Codigos_Tabla o;
            if (dic.TryGetValue(codigoproceso + codigopaso, out o))
                return o.descripcion;
            else
                return "N/A";


        }
    
    
    }
}
