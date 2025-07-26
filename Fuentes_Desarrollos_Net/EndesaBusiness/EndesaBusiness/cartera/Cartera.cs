using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.cartera
{
    class Cartera : EndesaEntity.cobros.CarteraSIOC_Tabla
    {
        Dictionary<string, EndesaEntity.cobros.CarteraSIOC_Tabla> dic;

        public Cartera(List<string> lista_nifs)
        {
            dic = new Dictionary<string, EndesaEntity.cobros.CarteraSIOC_Tabla>();
            Carga(lista_nifs);
        }

        private void Carga(List<string> lista_nifs)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;

            try
            {

                strSql = "select cif, nombreGestor, Apellido1Gestor, Apellido2Gestor,  direccion,"
                    + " descResponsableTerritorial, responsableTerritorial, responsableZona"
                    + " from carteraSIOC where cif in (";

                for (int i = 0; i < lista_nifs.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista_nifs[i] + "'";
                        firstOnly = false;
                    }
                    else
                    {
                        strSql += " ,'" + lista_nifs[i] + "'";
                    }

                }

                strSql += ");";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.COB);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cobros.CarteraSIOC_Tabla c = new EndesaEntity.cobros.CarteraSIOC_Tabla();
                    c.cif = r["cif"].ToString();

                    if (r["nombreGestor"] != System.DBNull.Value)
                        c.nombreGestor = r["nombreGestor"].ToString();

                    if (r["Apellido1Gestor"] != System.DBNull.Value)
                        c.apellido1Gestor = r["Apellido1Gestor"].ToString();

                    if (r["Apellido2Gestor"] != System.DBNull.Value)
                        c.apellido2Gestor = r["Apellido2Gestor"].ToString();

                    if (r["direccion"] != System.DBNull.Value)
                        c.direccion = r["direccion"].ToString();

                    if (r["descResponsableTerritorial"] != System.DBNull.Value)
                        c.descResponsableTerritorial = r["descResponsableTerritorial"].ToString();

                    if (r["responsableTerritorial"] != System.DBNull.Value)
                        c.responsableTerritorial = r["responsableTerritorial"].ToString();

                    if (r["responsableZona"] != System.DBNull.Value)
                        c.responsableZona = r["responsableZona"].ToString();

                    dic.Add(c.cif, c);

                }
                db.CloseConnection();



            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Cartera - Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        public void GetCartera(string nif)
        {
            EndesaEntity.cobros.CarteraSIOC_Tabla o;
            if (dic.TryGetValue(nif, out o))
            {
                this.apellido1Gestor = o.apellido1Gestor;
                this.apellido2Gestor = o.apellido2Gestor;
                this.nombreGestor = o.nombreGestor;
                this.responsableTerritorial = o.responsableTerritorial;
                this.responsableZona = o.responsableZona;
                this.direccion = o.direccion;
                this.descResponsableTerritorial = o.descResponsableTerritorial;
                
            }
            else
            {
                this.apellido1Gestor = "";
                this.apellido2Gestor = "";
                this.nombreGestor = "";
                this.responsableTerritorial = "";
                this.responsableZona = "";
                this.direccion = "";
                this.descResponsableTerritorial = "";
            }

        }

        public bool ExisteCartera(string nif)
        {
            bool existe = false;
            EndesaEntity.cobros.CarteraSIOC_Tabla o;
            if (dic.TryGetValue(nif, out o))
            {
                GetCartera(nif);
                existe = true;
            }


            return existe;
        }

        public string Direccion(string nif)
        {
            EndesaEntity.cobros.CarteraSIOC_Tabla o;
            if (dic.TryGetValue(nif, out o))
                return o.direccion;
            else
                return "";

        
        }
    }
}
