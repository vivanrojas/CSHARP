using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.global
{
    public class Provincias : EndesaEntity.global.ProvinciaTabla
    {
        Dictionary<string, EndesaEntity.global.ProvinciaTabla> dic = new Dictionary<string, EndesaEntity.global.ProvinciaTabla>();

        public Provincias(string nombreTabla, servidores.MySQLDB.Esquemas esquema)
        {
            CargaDatos(nombreTabla, esquema);
        }

        private void CargaDatos(string nombreTabla, servidores.MySQLDB.Esquemas esquema)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select Territorio, DesProvincia, CodigoPostal from " + nombreTabla;
                db = new MySQLDB(esquema);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.global.ProvinciaTabla c = new EndesaEntity.global.ProvinciaTabla();

                    c.territorio = r["Territorio"].ToString();
                    c.des_provincia = r["DesProvincia"].ToString();
                    c.codPostal = r["CodigoPostal"].ToString();

                    dic.Add(c.codPostal, c);
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Estados_Documento - CargaDatos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public string DesProvincia(string codProvincia)
        {
            string res = "";
            EndesaEntity.global.ProvinciaTabla o;
            if(codProvincia.Length >= 2)
                if (dic.TryGetValue(codProvincia.Substring(0,2), out o))
                    res = o.des_provincia;

            return res;

        }
    }
}
