    using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida.Redshift
{
    public class FuentesMedidaFunciones: EndesaEntity.medida.FuentesMedida
    {

        Dictionary<int, FuentesMedida> dic_fuentes;
        public FuentesMedidaFunciones()
        {
            dic_fuentes = new Dictionary<int, FuentesMedida>();
            this.GetAll();
        }

        private void GetAll()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                strSql = "select * from med_fuentesmedida m";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    FuentesMedida f = new FuentesMedida();
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

                    FuentesMedida c;
                    if (!dic_fuentes.TryGetValue(f.num, out c))
                    {
                        dic_fuentes.Add(f.num, f);
                    }
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {

            }
        }

        public string FuenteHoraria(int codigo)
        {
            FuentesMedida c;
            if (dic_fuentes.TryGetValue(codigo, out c))
                return c.fuente_horaria;

            return "";
        }

        public string FuenteCuartoHoraria(int codigo)
        {

            FuentesMedida c;
            if (dic_fuentes.TryGetValue(codigo, out c))
                return c.fuente_cuarto_horaria;

            return "";
        }

        public string FuenteFinal(int codigoCuartoHoraria, int codigoHoraria)
        {
            FuentesMedida c;
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


        //public void GetNum(int fuente_horaria, int fuente_cuartohoraria)
        //{
        //    FuentesMedida c;
        //    if (dic_fuentes.TryGetValue(fuente_horaria, out c))
        //    {
        //        this.num = c.num;
        //        this.fuente = c.fuente;
        //        this.fuente_madre = c.fuente_madre;
        //        this.fuente_final = c.fuente_final;

        //        if (c.fuente_cuarto_horaria == "INFORMAR_HORARIA")
        //            this.fuente_cuarto_horaria = c.fuente_cuarto_horaria;

        //        this.fuente_horaria = c.fuente_horaria;
        //        this.fuente_final_2 = c.fuente_final_2;

        //        if (this.fuente_final_2 == "IGUAL QUE LA HORARIA")
        //        {
        //            if (dic_fuentes.TryGetValue(fuente_cuartohoraria, out c))
        //            {
        //                this.num = c.num;
        //                this.fuente = c.fuente;
        //                this.fuente_madre = c.fuente_madre;
        //                this.fuente_final = c.fuente_final;

        //                if (c.fuente_cuarto_horaria != null)
        //                    this.fuente_cuarto_horaria = c.fuente_cuarto_horaria;

        //                this.fuente_horaria = c.fuente_horaria;
        //                this.fuente_final_2 = c.fuente_final_2;
        //            }

        //        }
        //    }
        //}

    }
}
