using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class FuentesMedida : EndesaEntity.medida.Fuentes
    {
        Dictionary<int, EndesaEntity.medida.Fuentes> dic_fuentes;

        public FuentesMedida()
        {
            dic_fuentes = new Dictionary<int, EndesaEntity.medida.Fuentes>();
            this.GetAll();
        }

        private void GetAll()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select NUM, FUENTE, FUENTEMADRE, FUENTEFINAL," 
                    + " FUENTECUARTOHORARIA, FUENTEHORARIA, FUENTEFINAL2"
                    + " from med_fuentesmedida m";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Fuentes f = new EndesaEntity.medida.Fuentes();
                    f.num = Convert.ToInt32(r["NUM"]);
                    f.fuente = r["FUENTE"].ToString();
                    f.fuente_madre = r["FUENTEMADRE"].ToString();
                    f.fuente_final = r["FUENTEFINAL"].ToString();

                    if (r["FUENTECUARTOHORARIA"] != System.DBNull.Value)
                        f.fuente_cuarto_horaria = r["FUENTECUARTOHORARIA"].ToString();

                    f.fuente_horaria = r["FUENTEHORARIA"].ToString();
                    if (r["FUENTEFINAL2"].ToString() == "IGUAL QUE LA CUARTOHORARIA")
                        f.fuente_final_2 = r["FUENTECUARTOHORARIA"].ToString();
                    else
                        f.fuente_final_2 = r["FUENTEHORARIA"].ToString();


                    EndesaEntity.medida.Fuentes c;
                    
                    if (!dic_fuentes.TryGetValue(f.num, out c))
                    {
                        dic_fuentes.Add(f.num, f);
                    }
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "FuentesMedidaFunciones.GetAll", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string FuenteHoraria(int codigo)
        {
            EndesaEntity.medida.Fuentes c;
            if (dic_fuentes.TryGetValue(codigo, out c))
                return c.fuente_horaria;

            return "";
        }

        public string FuenteCuartoHoraria(int codigo)
        {

            EndesaEntity.medida.Fuentes c;
            if (dic_fuentes.TryGetValue(codigo, out c))
                return c.fuente_cuarto_horaria;

            return "";
        }

        public string FuenteFinal(int codigoCuartoHoraria, int codigoHoraria)
        {
            EndesaEntity.medida.Fuentes c;
            if (dic_fuentes.TryGetValue(codigoCuartoHoraria, out c))
            {
                if (c.fuente_cuarto_horaria == "CALCULADA")
                {
                    return FuenteHoraria(codigoHoraria);
                }
                else
                {
                    return c.fuente_cuarto_horaria;
                }
            }


            return "";
        }

    }
}
