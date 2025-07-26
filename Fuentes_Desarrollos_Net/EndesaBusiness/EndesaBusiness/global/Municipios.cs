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
    public class Municipios: EndesaEntity.global.MunicipioTabla
    {
        Dictionary<string, string> dic = new Dictionary<string, string>();

        public Municipios()
        {
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
                strSql = "select CPRO, CMUN, DC, NOMBRE from cont_codmun";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.global.MunicipioTabla c = new EndesaEntity.global.MunicipioTabla();

                    c.codMunicipio = r["CMUN"].ToString() + r["CPRO"].ToString() + r["DC"].ToString();
                    c.descripcion = r["NOMBRE"].ToString();                    

                    dic.Add(c.codMunicipio, c.descripcion);
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

        public string DesMunicipio(string codMunicipio)
        {
            string res = "";
            string o;
            if (dic.TryGetValue(codMunicipio, out o))
                res = o;

            return res;

        }
    }
}
