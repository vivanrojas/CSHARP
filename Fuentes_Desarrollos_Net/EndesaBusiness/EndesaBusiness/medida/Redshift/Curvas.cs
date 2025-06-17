using EndesaBusiness.servidores;
using EndesaEntity;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types;

namespace EndesaBusiness.medida.Redshift
{
    public class Curvas
    {

        utilidades.Fechas fechas;
        logs.Log ficheroLog;
        public Dictionary<string, List<CurvaCuartoHoraria>> dic_cc { get; set; }
        public Curvas()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_CurvasGestores");
            fechas = new utilidades.Fechas();
            dic_cc = new Dictionary<string, List<CurvaCuartoHoraria>>();
        }


        public bool GetCurvaGestor(Estados_Curvas estado_curvas, List<string> lista, DateTime fd, DateTime fh, List<string> lista_estados)
        {
            string strSql = "";
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int j = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;

            try
            {
                strSql = "SELECT CD_PUNTO_MED as CUPS15, CD_CUPS_EXT as CUPS22, FH_LECT_REGISTRO as FECHA,";
                // HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_AC_H" + i + " as A" + i + ",";
                }
                // HORA REACTIVA R1
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R1_H" + i + " as R1_" + i + ",";
                }
                // HORA REACTIVA R4
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R4_H" + i + " as R4_" + i + ",";
                }
                // FUENTE HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_HOR_AC_H" + i + " as FH" + i + ",";
                }
                // CUARTOHORARIA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    for (int y = 1; y <= 4; y++)
                    {
                        z++;
                        strSql += " NM_POT_AC_H" + i + "_CUAD" + y + " as V" + z + ",";
                    }
                }
                // FUENTE CUARTOHORARIA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_CUARTH_AC_H" + i + "_CUAD1 as FCH" + i + ",";
                }


                strSql += " CD_SEC_RESUMEN AS VERSION, CD_ESTADO_CURVA"
                + " FROM METRA_OWNER.T_ED_H_CURVAS WHERE"
                + " cd_cups_ext_20 in (";

                for (int y = 0; y < lista.Count; y++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista[y] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista[y] + "'";

                }

                strSql += ") AND"
                + " (FH_LECT_REGISTRO >= " + fd.ToString("yyyyMMdd")
                + " AND FH_LECT_REGISTRO <= " + fh.ToString("yyyyMMdd") + ")"
                + " and CD_ESTADO_CURVA in ('" + lista_estados[0] + "'";

                for (int i = 1; i < lista_estados.Count; i++)
                    strSql += ",'" + lista_estados[i] + "'";

                strSql += ") ORDER BY CD_CUPS_EXT, FH_LECT_REGISTRO, CD_SEC_RESUMEN DESC;";


                ficheroLog.Add(strSql);

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                command.CommandTimeout = 0;
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = fechas.NumPeriodosHorarios(fechaHora);

                    CurvaCuartoHoraria c = new CurvaCuartoHoraria();

                    if (r["CUPS15"] != System.DBNull.Value)
                        c.cups15 = r["CUPS15"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["CD_ESTADO_CURVA"] != System.DBNull.Value)                        
                        c.estado = estado_curvas.GetDescripcion_Estado_Curva(r["CD_ESTADO_CURVA"].ToString());

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    j = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["A" + h] != System.DBNull.Value && r["A" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["A" + h]);

                        if (r["FH" + h] != System.DBNull.Value && r["FH" + h].ToString().Trim() != "")
                            c.fa[h] = r["FH" + h].ToString();

                        if (r["R1_" + h] != System.DBNull.Value && r["R1_" + h].ToString() != "")
                            c.r1[h] = Convert.ToDouble(r["R1_" + h]);

                        if (r["R4_" + h] != System.DBNull.Value && r["R4_" + h].ToString() != "")
                            c.r4[h] = Convert.ToDouble(r["R4_" + h]);

                        if (r["FCH" + h] != System.DBNull.Value && r["FCH" + h].ToString().Trim() != "")
                            c.fc[h] = r["FCH" + h].ToString();
                    }
                    for (int cc = 1; cc <= 100; cc++)
                        if (r["V" + cc] != System.DBNull.Value && r["V" + cc].ToString() != "")
                            c.value[cc] = Convert.ToDouble(r["V" + cc]);

                    

                    List<CurvaCuartoHoraria> o;
                    if (!dic_cc.TryGetValue(c.cups22, out o))
                    {

                        o = new List<CurvaCuartoHoraria>();
                        o.Add(c);
                        dic_cc.Add(c.cups22, o);
                    }
                    else
                    {
                        if (!o.Exists(zz => zz.cups22 == c.cups22 && zz.fecha == c.fecha))
                            o.Add(c);
                    }
                                        

                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }

        public bool CargaMedidaCuartoHoraria(Estados_Curvas estado_curvas, List<string> lista, DateTime fd, DateTime fh, List<string> lista_estados)
        {
            string strSql = "";
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int j = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;

            try
            {
                strSql = "SELECT CD_PUNTO_MED as CUPS15, CD_CUPS_EXT as CUPS22, FH_LECT_REGISTRO as FECHA,";
                // HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_AC_H" + i + " as A" + i + ",";
                }
                // HORA REACTIVA R1
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R1_H" + i + " as R1_" + i + ",";
                }
                // HORA REACTIVA R4
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R4_H" + i + " as R4_" + i + ",";
                }
                // FUENTE HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_HOR_AC_H" + i + " as FH" + i + ",";
                }
                // CUARTOHORARIA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    for (int y = 1; y <= 4; y++)
                    {
                        z++;
                        strSql += " NM_POT_AC_H" + i + "_CUAD" + y + " as V" + z + ",";
                    }
                }
                // FUENTE CUARTOHORARIA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_CUARTH_AC_H" + i + "_CUAD1 as FCH" + i + ",";
                }


                strSql += " CD_SEC_RESUMEN AS VERSION, CD_ESTADO_CURVA"
                + " FROM METRA_OWNER.T_ED_H_CURVAS WHERE"
                + " cd_cups_ext_20 in (";

                for (int y = 0; y < lista.Count; y++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista[y] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista[y] + "'";

                }

                strSql += ") AND"
                + " (FH_LECT_REGISTRO >= " + fd.ToString("yyyyMMdd")
                + " AND FH_LECT_REGISTRO <= " + fh.ToString("yyyyMMdd") + ")"
                + " and CD_ESTADO_CURVA in ('" + lista_estados[0] + "'";

                for (int i = 1; i < lista_estados.Count; i++)
                    strSql += ",'" + lista_estados[i] + "'";

                strSql += ") ORDER BY CD_CUPS_EXT, FH_LECT_REGISTRO, CD_SEC_RESUMEN DESC;";


                ficheroLog.Add(strSql);

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = fechas.NumPeriodosHorarios(fechaHora);

                    CurvaCuartoHoraria c = new CurvaCuartoHoraria();

                    if (r["CUPS15"] != System.DBNull.Value)
                        c.cups15 = r["CUPS15"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["CD_ESTADO_CURVA"] != System.DBNull.Value)
                        c.estado = estado_curvas.GetDescripcion_Estado_Curva(r["CD_ESTADO_CURVA"].ToString());

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    j = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["A" + h] != System.DBNull.Value && r["A" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["A" + h]);

                        if (r["FH" + h] != System.DBNull.Value && r["FH" + h].ToString().Trim() != "")
                            c.fa[h] = r["FH" + h].ToString();

                        if (r["R1_" + h] != System.DBNull.Value && r["R1_" + h].ToString() != "")
                            c.r1[h] = Convert.ToDouble(r["R1_" + h]);

                        if (r["R4_" + h] != System.DBNull.Value && r["R4_" + h].ToString() != "")
                            c.r4[h] = Convert.ToDouble(r["R4_" + h]);

                        if (r["FCH" + h] != System.DBNull.Value && r["FCH" + h].ToString().Trim() != "")
                            c.fc[h] = r["FCH" + h].ToString();
                    }
                    for (int cc = 1; cc <= 100; cc++)
                        if (r["V" + cc] != System.DBNull.Value && r["V" + cc].ToString() != "")
                            c.value[cc] = Convert.ToDouble(r["V" + cc]);



                    List<CurvaCuartoHoraria> o;
                    if (!dic_cc.TryGetValue(c.cups22, out o))
                    {

                        o = new List<CurvaCuartoHoraria>();
                        o.Add(c);
                        dic_cc.Add(c.cups22, o);
                    }
                    else
                    {
                        if (!o.Exists(zz => zz.cups22 == c.cups22 && zz.fecha == c.fecha))
                            o.Add(c);
                    }


                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }

        public bool CargaMedidaHoraria(Estados_Curvas estado_curvas, List<string> lista, DateTime fd, DateTime fh, List<string> lista_estados)
        {
            string strSql = "";
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int j = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;

            try
            {
                strSql = "SELECT CD_PUNTO_MED as CUPS15, CD_CUPS_EXT as CUPS22, FH_LECT_REGISTRO as FECHA,";

                // HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_AC_H" + i + " as A" + i + ",";
                }
                // HORA REACTIVA R1
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R1_H" + i + " as R1_" + i + ",";
                }

                // HORA REACTIVA R2
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R2_H" + i + " as R2_" + i + ",";
                }

                // HORA REACTIVA R3
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R3_H" + i + " as R3_" + i + ",";
                }

                // HORA REACTIVA R4
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R4_H" + i + " as R4_" + i + ",";
                }
                // FUENTE HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_HOR_AC_H" + i + " as FH" + i + ",";
                }
                

                strSql += " CD_SEC_RESUMEN AS VERSION, CD_ESTADO_CURVA"
                + " FROM METRA_OWNER.T_ED_H_CURVAS WHERE"
                + " cd_cups_ext_20 in (";

                for (int y = 0; y < lista.Count; y++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista[y] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista[y] + "'";

                }

                strSql += ") AND"
                + " (FH_LECT_REGISTRO >= " + fd.ToString("yyyyMMdd")
                + " AND FH_LECT_REGISTRO <= " + fh.ToString("yyyyMMdd") + ")"
                + " and CD_ESTADO_CURVA in ('" + lista_estados[0] + "'";

                for (int i = 1; i < lista_estados.Count; i++)
                    strSql += ",'" + lista_estados[i] + "'";

                strSql += ") ORDER BY CD_CUPS_EXT, FH_LECT_REGISTRO, CD_SEC_RESUMEN DESC;";


                ficheroLog.Add(strSql);

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = fechas.NumPeriodosHorarios(fechaHora);

                    CurvaCuartoHoraria c = new CurvaCuartoHoraria();

                    if (r["CUPS15"] != System.DBNull.Value)
                        c.cups15 = r["CUPS15"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["CD_ESTADO_CURVA"] != System.DBNull.Value)
                        c.estado = estado_curvas.GetDescripcion_Estado_Curva(r["CD_ESTADO_CURVA"].ToString());

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    j = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["A" + h] != System.DBNull.Value && r["A" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["A" + h]);

                        if (r["FH" + h] != System.DBNull.Value && r["FH" + h].ToString().Trim() != "")
                            c.fa[h] = r["FH" + h].ToString();

                        if (r["R1_" + h] != System.DBNull.Value && r["R1_" + h].ToString() != "")
                            c.r1[h] = Convert.ToDouble(r["R1_" + h]);

                        if (r["R2_" + h] != System.DBNull.Value && r["R2_" + h].ToString() != "")
                            c.r2[h] = Convert.ToDouble(r["R1_" + h]);

                        if (r["R3_" + h] != System.DBNull.Value && r["R3_" + h].ToString() != "")
                            c.r3[h] = Convert.ToDouble(r["R3_" + h]);

                        if (r["R4_" + h] != System.DBNull.Value && r["R4_" + h].ToString() != "")
                            c.r4[h] = Convert.ToDouble(r["R4_" + h]);

                        
                    }
                    



                    List<CurvaCuartoHoraria> o;
                    if (!dic_cc.TryGetValue(c.cups22, out o))
                    {

                        o = new List<CurvaCuartoHoraria>();
                        o.Add(c);
                        dic_cc.Add(c.cups22, o);
                    }
                    else
                    {
                        if (!o.Exists(zz => zz.cups22 == c.cups22 && zz.fecha == c.fecha))
                            o.Add(c);
                    }


                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }

       

        public bool CargaMedidaHorariaCups22(Estados_Curvas estado_curvas, List<string> lista, DateTime fd, DateTime fh, List<string> lista_estados)
        {
            string strSql = "";
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int j = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;

            try
            {
                strSql = "SELECT CD_PUNTO_MED as CUPS15, CD_CUPS_EXT as CUPS22, FH_LECT_REGISTRO as FECHA,";

                // HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_AC_H" + i + " as A" + i + ",";
                }
                // HORA REACTIVA R1
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R1_H" + i + " as R1_" + i + ",";
                }

                // HORA REACTIVA R2
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R2_H" + i + " as R2_" + i + ",";
                }

                // HORA REACTIVA R3
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R3_H" + i + " as R3_" + i + ",";
                }

                // HORA REACTIVA R4
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " NM_ENER_R4_H" + i + " as R4_" + i + ",";
                }
                // FUENTE HORA ACTIVA
                for (int i = 1; i <= 25; i++)
                {
                    strSql += " CD_FUENTE_HOR_AC_H" + i + " as FH" + i + ",";
                }


                strSql += " CD_SEC_RESUMEN AS VERSION, CD_ESTADO_CURVA"
                + " FROM METRA_OWNER.T_ED_H_CURVAS WHERE"
                + " CD_CUPS_EXT in (";

                for (int y = 0; y < lista.Count; y++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista[y] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += ",'" + lista[y] + "'";

                }

                strSql += ") AND"
                + " (FH_LECT_REGISTRO >= " + fd.ToString("yyyyMMdd")
                + " AND FH_LECT_REGISTRO <= " + fh.ToString("yyyyMMdd") + ")"
                + " and CD_ESTADO_CURVA in ('" + lista_estados[0] + "'";

                for (int i = 1; i < lista_estados.Count; i++)
                    strSql += ",'" + lista_estados[i] + "'";

                strSql += ") ORDER BY CD_CUPS_EXT, FH_LECT_REGISTRO, CD_SEC_RESUMEN DESC;";


                ficheroLog.Add(strSql);

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = fechas.NumPeriodosHorarios(fechaHora);

                    CurvaCuartoHoraria c = new CurvaCuartoHoraria();

                    if (r["CUPS15"] != System.DBNull.Value)
                        c.cups15 = r["CUPS15"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["CD_ESTADO_CURVA"] != System.DBNull.Value)
                        c.estado = estado_curvas.GetDescripcion_Estado_Curva(r["CD_ESTADO_CURVA"].ToString());

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    j = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["A" + h] != System.DBNull.Value && r["A" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["A" + h]);

                        if (r["FH" + h] != System.DBNull.Value && r["FH" + h].ToString().Trim() != "")
                            c.fa[h] = r["FH" + h].ToString();

                        if (r["R1_" + h] != System.DBNull.Value && r["R1_" + h].ToString() != "")
                            c.r1[h] = Convert.ToDouble(r["R1_" + h]);

                        if (r["R2_" + h] != System.DBNull.Value && r["R2_" + h].ToString() != "")
                            c.r2[h] = Convert.ToDouble(r["R1_" + h]);

                        if (r["R3_" + h] != System.DBNull.Value && r["R3_" + h].ToString() != "")
                            c.r3[h] = Convert.ToDouble(r["R3_" + h]);

                        if (r["R4_" + h] != System.DBNull.Value && r["R4_" + h].ToString() != "")
                            c.r4[h] = Convert.ToDouble(r["R4_" + h]);


                    }




                    List<CurvaCuartoHoraria> o;
                    if (!dic_cc.TryGetValue(c.cups22, out o))
                    {

                        o = new List<CurvaCuartoHoraria>();
                        o.Add(c);
                        dic_cc.Add(c.cups22, o);
                    }
                    else
                    {
                        if (!o.Exists(zz => zz.cups22 == c.cups22 && zz.fecha == c.fecha))
                            o.Add(c);
                    }


                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CurvasRedShift.GetCurva: " + e.Message);
                return true;
            }

        }

        public List<EndesaEntity.CurvaCuartoHoraria> GetCurva(string cups22, DateTime fd, DateTime fh)
        {
            
            List<EndesaEntity.CurvaCuartoHoraria> lista = new List<EndesaEntity.CurvaCuartoHoraria>();
            Dictionary<DateTime, List<EndesaEntity.CurvaCuartoHoraria>> d =
                new Dictionary<DateTime, List<CurvaCuartoHoraria>>();

            

            Dictionary<string, List<EndesaEntity.CurvaCuartoHoraria>> matches =
                dic_cc.Where(z => z.Key == cups22).ToDictionary(z => z.Key, z => z.Value);

            foreach (KeyValuePair<string, List<EndesaEntity.CurvaCuartoHoraria>> p in matches)
            {                
                lista = p.Value.Where(z => z.fecha >= fd && z.fecha <= fh).ToList();
            }
             return lista;

        }

        public bool ExisteCurva(string cups22, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.CurvaCuartoHoraria> lista = new List<EndesaEntity.CurvaCuartoHoraria>();
            Dictionary<DateTime, List<EndesaEntity.CurvaCuartoHoraria>> d =
                new Dictionary<DateTime, List<CurvaCuartoHoraria>>();



            Dictionary<string, List<EndesaEntity.CurvaCuartoHoraria>> matches =
                dic_cc.Where(z => z.Key == cups22).ToDictionary(z => z.Key, z => z.Value);

            foreach (KeyValuePair<string, List<EndesaEntity.CurvaCuartoHoraria>> p in matches)
            {
                lista = p.Value.Where(z => z.fecha >= fd && z.fecha <= fh).ToList();
            }
            
            return lista.Count> 0;

        }


        public List<EndesaEntity.CurvaCuartoHoraria> GetCurvaCups20(string cups20, DateTime fd, DateTime fh)
        {           

            List<EndesaEntity.CurvaCuartoHoraria> lista = new List<EndesaEntity.CurvaCuartoHoraria>();
            Dictionary<DateTime, List<EndesaEntity.CurvaCuartoHoraria>> d =
                new Dictionary<DateTime, List<CurvaCuartoHoraria>>();

            Dictionary<string, List<EndesaEntity.CurvaCuartoHoraria>> matches =
                dic_cc.Where(z => z.Key.Substring(0,20) == cups20).ToDictionary(z => z.Key, z => z.Value);

            foreach (KeyValuePair<string, List<EndesaEntity.CurvaCuartoHoraria>> p in matches)                
                foreach(EndesaEntity.CurvaCuartoHoraria pp in p.Value.Where(z => z.fecha >= fd && z.fecha <= fh).ToList())                
                    lista.Add(pp);
               
            
            return lista;

        }

        public double TotalActiva(string cups22, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.CurvaCuartoHoraria> lista = GetCurva(cups22, fd, fh);
            if (lista != null)
                foreach (EndesaEntity.CurvaCuartoHoraria p in lista)
                    for (int i = 1; i < 26; i++)
                        total += p.a[i];
                

            return total;
        }

        public double TotalReactiva(string cups22, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.CurvaCuartoHoraria> lista = GetCurva(cups22, fd, fh);
            if (lista != null)
                foreach (EndesaEntity.CurvaCuartoHoraria p in lista)
                    for (int i = 1; i < 26; i++)
                        total += p.r1[i];

            return total;
        }

        public double TotalActivaCups20(string cups20, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.CurvaCuartoHoraria> lista = GetCurvaCups20(cups20, fd, fh);
            if (lista != null)
                foreach (EndesaEntity.CurvaCuartoHoraria p in lista)
                    for (int i = 1; i < 26; i++)
                        total += p.a[i];


            return total;
        }

        public double TotalReactivaCups20(string cups20, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.CurvaCuartoHoraria> lista = GetCurvaCups20(cups20, fd, fh);
            if (lista != null)
                foreach (EndesaEntity.CurvaCuartoHoraria p in lista)
                    for (int i = 1; i < 26; i++)
                        total += p.r1[i];

            return total;
        }

    }
}
