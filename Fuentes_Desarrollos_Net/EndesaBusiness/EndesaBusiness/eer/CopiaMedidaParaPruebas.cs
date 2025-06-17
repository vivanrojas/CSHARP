using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.eer
{
    public class CopiaMedidaParaPruebas
    {

        public List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
        public CopiaMedidaParaPruebas()
        {
            DateTime fd = new DateTime(2020, 03, 01);
            DateTime fh = new DateTime(2020, 03, 31);
            List<string> lista_cups = new List<string>();
            lista_cups.Add("ES0031405083724001RS");
            lista_cups.Add("ES0031405427879001ZN");
            lista_cups.Add("ES0031405475635001ZP");
            lista_cups.Add("ES0031405475635002ZD");
            lista_cups.Add("ES0031406125921001QQ");
            lista_cups.Add("ES0031406143055001AM");
            lista_cups.Add("ES0031406349751001HL");
            lista_cups.Add("ES0031408599774006XT");
            lista_cups.Add("ES0031408599774020XZ");
            lista_cups.Add("ES0031408599774021XS");
            lista_cups.Add("ES0031408599774022XQ");
            lista_cups.Add("ES0031408599774023XV");



            GetCurvaRedShift(lista_cups, fd, fh, "F");
            

        }



        private void GetCurvaRedShift(List<string> lista_cups20, DateTime fd, DateTime fh, string estado)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            MySQLDB dbm;
            MySqlCommand commandm;


            string strSql;
            StringBuilder sb = new StringBuilder();

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int y = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;
            int numReg = 0;

            int totala = 0;
            int totalr = 0;

            try
            {
                calendarios.UtilidadesCalendario util = new calendarios.UtilidadesCalendario();

                #region Query

                strSql = "SELECT cd_punto_med as cups15, cd_cups_ext as cups22, fh_lect_registro as fecha,  cd_estado_curva as estado ";
                // HORA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,NM_ENER_AC_H" + j + " as a" + j;

                // HORA REACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,NM_ENER_R1_H" + j + " as r" + j;

                //// FUENTE HORA ACTIVA
                //for (int j = 1; j <= 25; j++)
                //    strSql += " ,CD_FUENTE_HOR_AC_H" + j + " as fh" + j;

                // CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 25; j++)
                    for (int w = 1; w <= 4; w++)
                    {
                        z++;
                        strSql += " ,NM_POT_AC_H" + j + "_CUAD" + w + " as v" + z;
                    }


                // FUENTE CUARTOHORARIA ACTIVA
                //for (int j = 1; j <= 25; j++)
                //    strSql += " ,CD_FUENTE_CUARTH_AC_H" + j + " as fch" + j;



                strSql += " from metra_owner.t_ed_h_curvas "
                    + " WHERE cd_cups_ext_20 in ('" + lista_cups20[0] + "'";

                for (int i = 1; i < lista_cups20.Count; i++)
                    strSql += ",'" + lista_cups20[i] + "'";

                strSql += ") and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
                    + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and cd_estado_curva = '" + estado + "'"
                    + " ORDER BY cd_punto_med, fh_lect_registro;";

                #endregion

                db = new RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    totala = 0;
                    totalr = 0;

                    numReg++;
                    EndesaEntity.medida.CurvaDeCarga c = new EndesaEntity.medida.CurvaDeCarga();

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = util.NumPeriodosHorarios(fechaHora);

                    c.fecha = fechaHora;
                    c.numPeriodosHorarios = numeroPeriodos;
                    c.numPeriodosCuartoHorarios = c.numPeriodosHorarios * 4;
                    
                    if (firstOnly)
                    {
                        sb.Append("replace into eer_cc (cups22, fecha, estado, version, totala, totalr");
                        for (int j = 1; j <= 100; j++)
                            sb.Append(", v" + j);

                        for (int j = 1; j <= 25; j++)
                            sb.Append(", a" + j);

                        for (int j = 1; j <= 25; j++)
                            sb.Append(", r" + j);

                        sb.Append(", f_ult_mod) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(r["cups22"].ToString()).Append("',");
                    sb.Append("'").Append(fechaHora.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(r["estado"].ToString()).Append("'");
                    sb.Append(",1");

                    for (int i = 1; i < 26; i++)
                    {
                        if (r["a" + i] != System.DBNull.Value)
                            totala = totala + Convert.ToInt32(r["a" + i]);

                        if (r["r" + i] != System.DBNull.Value)
                            totalr = totalr + Convert.ToInt32(r["r" + i]);
                    }

                    sb.Append(" ,").Append(totala);
                    sb.Append(" ,").Append(totalr);

                    for (int i = 1; i < 101; i++)
                    {
                        if (r["v" + i] != System.DBNull.Value)
                            sb.Append(" ,").Append(Convert.ToDouble(r["v" + i]));
                        else
                            sb.Append(" ,null");
                    }

                    for (int i = 1; i < 26; i++)
                    {
                        if (r["a" + i] != System.DBNull.Value)
                            sb.Append(" ,").Append(Convert.ToDouble(r["a" + i]));
                        else
                            sb.Append(" ,null");
                    }

                    for (int i = 1; i < 26; i++)
                    {
                        if (r["r" + i] != System.DBNull.Value)
                            sb.Append(" ,").Append(Convert.ToDouble(r["r" + i]));
                        else
                            sb.Append(" ,null");
                    }

                    sb.Append(" ,'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    //EndesaEntity.medida.DicCurva o;
                    //if (!dic.TryGetValue(r["cups15"].ToString().Substring(0, 13), out o))
                    //{
                    //    EndesaEntity.medida.DicCurva d = new EndesaEntity.medida.DicCurva();
                    //    d.dic.Add(c.fecha, c);
                    //    dic.Add(r["cups15"].ToString().Substring(0, 13), d);
                    //}
                    //else
                    //{
                    //    EndesaEntity.medida.CurvaDeCarga x;
                    //    if (!o.dic.TryGetValue(c.fecha, out x))
                    //    {
                    //        o.dic.Add(c.fecha, c);
                    //    }
                    //}

                    // lista.Add(c);
                    if(numReg == 1)
                    {
                        firstOnly = true;
                        dbm = new MySQLDB(MySQLDB.Esquemas.CON);
                        commandm = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbm.con);
                        commandm.ExecuteNonQuery();
                        dbm.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }

                }

                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }
    }
}
