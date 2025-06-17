using EndesaBusiness.servidores;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.factoring
{
    public class PuntosActivosBTE_COMPOR : EndesaEntity.contratacion.PS_AT_Tabla
    { 
        public Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla> dic { get; set; }

        public PuntosActivosBTE_COMPOR(List<string> lista_cups_20)
        {

            dic = new Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla>();
            if (lista_cups_20.Count > 0)
                Carga(GetQuery(lista_cups_20));
        }

        public PuntosActivosBTE_COMPOR()
        {
            dic = new Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla>();
            Carga(GetQuery());
        }

        private string GetQuery(List<string> lista_cups_20)
        {
            string strSql = "";
            strSql = "SELECT CPE from COMPOR_OWNER.CONT_PUNTOS_ACTIVOS_BTE"
                + " where CPE in ('" + lista_cups_20[0] + "'";

            for (int i = 1; i < lista_cups_20.Count; i++)
                strSql += " ,'" + lista_cups_20[i] + "'";

            strSql += ");";

            return strSql;
        }

        private string GetQuery()
        {
            string strSql = "";

            strSql = "SELECT  CPE, TITULAR, CIF FROM COMPOR_OWNER.CONT_PUNTOS_ACTIVOS_BTE";

            return strSql;
        }


        private void Carga(string query)
        {
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;

            ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
            ora_command = new OracleCommand(query, ora_db.con);
            Console.WriteLine("Consultando BBDD ...");
            r = ora_command.ExecuteReader();
            Console.WriteLine("Copiando datos ...");
            while (r.Read())
            {
                if (r["CPE"] != System.DBNull.Value)
                {
                    EndesaEntity.contratacion.PS_AT_Tabla c = new EndesaEntity.contratacion.PS_AT_Tabla();
                    c.cups20 = r["CPE"].ToString();
                    if (r["TITULAR"] != System.DBNull.Value)
                        c.nombre_cliente = r["TITULAR"].ToString();

                    if (r["CIF"] != System.DBNull.Value)
                        c.cif = r["CIF"].ToString();

                    EndesaEntity.contratacion.PS_AT_Tabla o;
                    if (!dic.TryGetValue(c.cups20, out o))
                        dic.Add(c.cups20, c);
                }

            }
            ora_db.CloseConnection();
        }

        public bool ExisteAlta(string cups20)
        {
            bool existe = false;
            EndesaEntity.contratacion.PS_AT_Tabla o;
            if (dic.TryGetValue(cups20, out o))
            {
                existe = true;
                this.cups20 = o.cups20;
                // this.empresa = o.empresa;                
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
