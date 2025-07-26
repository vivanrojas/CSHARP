using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.factoring
{
    public class PuntosActivosGas : EndesaEntity.contratacion.PS_AT_Tabla
    {
        public Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla> dic { get; set; }

        public PuntosActivosGas(List<string> lista_cups_20)
        {

            dic = new Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla>();
            if (lista_cups_20.Count > 0)
                Carga(GetQuery(lista_cups_20));
        }

        public PuntosActivosGas()
        {
            dic = new Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla>();
            Carga(GetQuery());
        }

        private string GetQuery(List<string> lista_cups_20)
        {
            string strSql = "";
            strSql = "SELECT g.ID_PS, g.CNIFDNIC, g.DAPERSOC,"
                + " IF(LENGTH(g.CUPSREE) <> 20 OR g.CUPSREE IS NULL, CONCAT('CISTERNA_', g.ID_PS), g.CUPSREE) AS CUPSREE"
                + " FROM cm_inventario_gas g WHERE"
                + " g.FINICIO <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' AND"
                + " (g.FFIN >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR g.FFIN IS NULL) and"
                + " g.ID_ESTADO_CTO in (3, 6, 7, 8, 9, 10, 15, 16)"
                + " and g.CUPSREE in ('" + lista_cups_20[0] + "'";

            for (int i = 1; i < lista_cups_20.Count; i++)
                strSql += " ,'" + lista_cups_20[i] + "'";

            strSql += ")";

            return strSql;
        }

        private string GetQuery()
        {
            string strSql = "";

            strSql = "SELECT g.ID_PS, g.CNIFDNIC, g.DAPERSOC,"
                + " IF(LENGTH(g.CUPSREE) <> 20 OR g.CUPSREE IS NULL, CONCAT('CISTERNA_', g.ID_PS), g.CUPSREE) AS CUPSREE"
                + " FROM cm_inventario_gas g WHERE"
                + " g.FINICIO <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' AND"
                + " (g.FFIN >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR g.FFIN IS NULL) and"
                + " g.ID_ESTADO_CTO in (3, 6, 7, 8, 9, 10, 15, 16)";

            return strSql;
        }


        private void Carga(string query)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(query, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.PS_AT_Tabla c = new EndesaEntity.contratacion.PS_AT_Tabla();
                if (r["CUPSREE"] != System.DBNull.Value)
                {
                    c.cups20 = r["CUPSREE"].ToString();
                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.nombre_cliente = r["DAPERSOC"].ToString();
                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.cif = r["CNIFDNIC"].ToString();
                    //c.empresa = r["EMPRESA"].ToString();
                    //c.estado_contrato_descripcion = r["Descripcion"].ToString();

                    EndesaEntity.contratacion.PS_AT_Tabla o;
                    if (!dic.TryGetValue(c.cups20, out o))
                        dic.Add(c.cups20, c);
                }

            }
            db.CloseConnection();
        }

        public bool ExisteAlta(string cups20)
        {
            bool existe = false;
            EndesaEntity.contratacion.PS_AT_Tabla o;
            if (dic.TryGetValue(cups20, out o))
            {
                existe = true;
                this.cups20 = o.cups20;
                //this.empresa = o.empresa;
                //this.estado_contrato_descripcion = o.estado_contrato_descripcion;
                this.nombre_cliente = o.nombre_cliente;
                this.cif = o.cif;
            }

            return existe;
        }

        public bool ExisteCIF(string cif)
        {
            bool existe = false;
            EndesaEntity.contratacion.PS_AT_Tabla o;
            foreach (KeyValuePair<string, EndesaEntity.contratacion.PS_AT_Tabla> p in dic)
            {
                if (p.Value.cif == cif)
                {
                    existe = true;
                    break;
                }
            }

            return existe;
        }
    }
}
