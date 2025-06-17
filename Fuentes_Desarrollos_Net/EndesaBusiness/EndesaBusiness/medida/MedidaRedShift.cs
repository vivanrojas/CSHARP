using EndesaBusiness.servidores;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class MedidaRedShift
    {
        public Dictionary<DateTime, List<EndesaEntity.medida.P1>> dic { get; set; }
        public List<EndesaEntity.medida.CurvaCuartoHorariaInformes> list_cc { get; set; }
        calendarios.UtilidadesCalendario util;

        public MedidaRedShift()
        {
            util = new calendarios.UtilidadesCalendario();
            list_cc = new List<EndesaEntity.medida.CurvaCuartoHorariaInformes>();
            dic = new Dictionary<DateTime, List<EndesaEntity.medida.P1>>();
        }

        public bool GetCurvaRedShift(string cups13, DateTime fd, DateTime fh, string estado)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;

            int y = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();

            try
            {
                list_cc.Clear();

                #region Query

                strSql = "SELECT cd_punto_med as cups15, cd_cups_ext as cups22, fh_lect_registro as fecha,  cd_estado_curva as estado ";
                // HORA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,NM_ENER_AC_H" + j + " as a" + j;

                // HORA REACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,NM_ENER_R1_H" + j + " as r" + j;

                // FUENTE HORA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,CD_FUENTE_HOR_AC_H" + j + " as fh" + j;

                // CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 25; j++)
                    for (int w = 1; w <= 4; w++)
                    {
                        z++;
                        strSql += " ,NM_POT_AC_H" + j + "_CUAD" + w + " as v" + z;
                    }


                // FUENTE CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 25; j++)
                    //strSql += " ,CD_FUENTE_CUARTH_AC_H" + j + " as fch" + j;
                    strSql += " CD_FUENTE_CUARTH_AC_H" + j + "_CUAD1 as fch" + j + ",";


                if (estado == "F")
                    strSql += " from metra_owner.t_ed_h_curvas "
                    + " WHERE cd_punto_med like '" + cups13 + "%'"
                    + " and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
                    + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and cd_estado_curva = '" + estado + "'"
                    + " ORDER BY cd_punto_med, fh_lect_registro;";
                else
                    strSql += " from metra_owner.t_ed_h_curvas "
                    + " WHERE cd_punto_med like '" + cups13 + "%'"
                    + " and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
                    + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and cd_estado_curva = '" + estado + "'"
                    + " ORDER BY cd_punto_med, fh_lect_registro;";

                #endregion

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = util.NumPeriodosHorarios(fechaHora);

                    EndesaEntity.medida.CurvaCuartoHorariaInformes c = new EndesaEntity.medida.CurvaCuartoHorariaInformes();

                    if (r["cups15"] != System.DBNull.Value)
                        c.cups15 = r["cups15"].ToString();

                    if (r["cups22"] != System.DBNull.Value)
                        c.cups22 = r["cups22"].ToString();

                    if (r["estado"] != System.DBNull.Value)
                        c.estado = r["estado"].ToString() == "F" ? "FACTURADA" : "REGISTRADA";

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    y = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["a" + h] != System.DBNull.Value && r["a" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["a" + h]);

                        if (r["fh" + h] != System.DBNull.Value && r["fh" + h].ToString() != "")
                            c.fa[h] = r["fh" + h].ToString();

                        if (r["r" + h] != System.DBNull.Value && r["r" + h].ToString() != "")
                            c.r[h] = Convert.ToDouble(r["r" + h]);

                        if (r["fch" + h] != System.DBNull.Value && r["fch" + h].ToString() != "")
                            c.fc[h] = r["fch" + h].ToString();
                    }
                    for (int cc = 1; cc <= 100; cc++)
                        if (r["v" + cc] != System.DBNull.Value && r["v" + cc].ToString() != "")
                            c.value[cc] = Convert.ToDouble(r["v" + cc]);

                    if (!list_cc.Exists(zz => zz.cups15 == c.cups15 && zz.fecha == c.fecha))
                        list_cc.Add(c);

                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return true;
            }


        }
        public bool GetCurvaRedShift2(string cups20, DateTime fd, DateTime fh, string estado)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;

            int y = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();

            try
            {
                list_cc.Clear();

                #region Query

                strSql = "SELECT cd_punto_med as cups15, cd_cups_ext as cups22, fh_lect_registro as fecha,  cd_estado_curva as estado ";
                // HORA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,NM_ENER_AC_H" + j + " as a" + j;

                // HORA REACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,NM_ENER_R1_H" + j + " as r" + j;

                // FUENTE HORA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,CD_FUENTE_HOR_AC_H" + j + " as fh" + j;

                // CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 25; j++)
                    for (int w = 1; w <= 4; w++)
                    {
                        z++;
                        strSql += " ,NM_POT_AC_H" + j + "_CUAD" + w + " as v" + z;
                    }


                // FUENTE CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,CD_FUENTE_CUARTH_AC_H" + j + " as fch" + j;


                if (estado == "F")
                    strSql += " from metra_owner.t_ed_h_curvas "
                    + " WHERE cd_cups like '" + cups20 + "%'"
                    + " and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
                    + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and cd_estado_curva = '" + estado + "'"
                    + " ORDER BY cd_punto_med, fh_lect_registro;";
                else
                    strSql += " from metra_owner.t_ed_h_curvas "
                    + " WHERE cd_cups_ext like '" + cups20 + "%'"
                    + " and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
                    + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and cd_estado_curva = '" + estado + "'"
                    + " ORDER BY cd_punto_med, fh_lect_registro;";

                #endregion

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = util.NumPeriodosHorarios(fechaHora);

                    EndesaEntity.medida.CurvaCuartoHorariaInformes c = new EndesaEntity.medida.CurvaCuartoHorariaInformes();

                    if (r["cups15"] != System.DBNull.Value)
                        c.cups15 = r["cups15"].ToString();

                    if (r["cups22"] != System.DBNull.Value)
                        c.cups22 = r["cups22"].ToString();

                    if (r["estado"] != System.DBNull.Value)
                        c.estado = r["estado"].ToString() == "F" ? "FACTURADA" : "REGISTRADA";

                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    y = 0;
                    for (int h = 1; h <= 25; h++)
                    {
                        if (r["a" + h] != System.DBNull.Value && r["a" + h].ToString() != "")
                            c.a[h] = Convert.ToDouble(r["a" + h]);

                        if (r["fh" + h] != System.DBNull.Value && r["fh" + h].ToString() != "")
                            c.fa[h] = r["fh" + h].ToString();

                        if (r["r" + h] != System.DBNull.Value && r["r" + h].ToString() != "")
                            c.r[h] = Convert.ToDouble(r["r" + h]);

                        if (r["fch" + h] != System.DBNull.Value && r["fch" + h].ToString() != "")
                            c.fc[h] = r["fch" + h].ToString();
                    }
                    for (int cc = 1; cc <= 100; cc++)
                        if (r["v" + cc] != System.DBNull.Value && r["v" + cc].ToString() != "")
                            c.value[cc] = Convert.ToDouble(r["v" + cc]);

                    if (!list_cc.Exists(zz => zz.cups15 == c.cups15 && zz.fecha == c.fecha))
                        list_cc.Add(c);

                }

                db.CloseConnection();
                return false;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return true;
            }


        }

        
        private int Estacion(DateTime f)
        {
            int estacion = 0; // Invierno

            if (f > UltimoDomingoMarzo(f) && f < UltimoDomingoOctubre(f))
                estacion = 1; // Verano
            else if (f == UltimoDomingoMarzo(f) && f.Hour > 2)
                estacion = 1;
            else if (f == UltimoDomingoOctubre(f) && f.Hour > 3)
                estacion = 0;
            else
                estacion = 0;

            return estacion;
        }

        private int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }

        private DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }


        //private void P1TelefonicaToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;

        //    List<string> lista_cups = new List<string>();
        //    GO.medida.MedidaRedShift red = new GO.medida.MedidaRedShift();
        //    GO.medida.FormatosExportacionMedida formatos = new GO.medida.FormatosExportacionMedida();

        //    //strSql = "SELECT CUPS FROM telefonica_p1_inventario GROUP BY CUPS";
        //    strSql = "SELECT r.CUPS_CORTO AS CUPS FROM aux1.RELACION_CUPS r WHERE r.CUPS20 IN "
        //        + "('ES0021000006334552SG','ES0021000006919858WY','ES0022000005731647GB','ES0022000005429851SK','ES0031101457795001XL',"
        //        + "'ES0031405997368001PL','ES0031406039804001SF','ES0021000000090637DZ','ES0021000006069435BP','ES0021000008019789PP',"
        //        + "'ES0021000008275465SQ','ES0021000004434850NX','ES0022000004991075MQ','ES0022000005731675MQ','ES0022000005732031KG',"
        //        + "'ES0031102030845001HM','ES0031102462991001KV','ES0031102975953001PD','ES0031104009885001GZ','ES0031300013979002FW',"
        //        + "'ES0031405539809001NY','ES0031406039812001HA','ES0031607342547001BM','ES0021000005297572DW','ES0021000009783742LE',"
        //        + "'ES0021000010185283KY','ES0021000005299863QQ','ES0022000005731740PN','ES0022000005732351NW','ES0031103165524001ZH',"
        //        + "'ES0031104001456001PP','ES0031405143999001XL','ES0031405449944001DP','ES0031406039935001YE','ES0031446437310001BS',"
        //        + "'ES0031607288919001FR','ES0021000006150927NB','ES0021000008087489FL','ES0021000008550944XR','ES0021000009532962HB',"
        //        + "'ES0021000004212816HH','ES0031101465519001JK','ES0031300239076001BA','ES0031406316306001VD','ES0021000006663370MZ',"
        //        + "'ES0021000002007873SQ','ES0022000005731386QA','ES0022000004991127FE','ES0022000005731331JV','ES0031102051574001RW',"
        //        + "'ES0031104010044001VS','ES0031102139371001CE','ES0031500112434001YA','ES0031405997387001YK','ES0031405760981001RG',"
        //        + "'ES0021000003223231AP','ES0021000006195104RM','ES0021000008798251KN','ES0022000005731504KY','ES0022000005731839NL',"
        //        + "'ES0022000007501204YV','ES0022000005731332JH','ES0022000005732092TL','ES0031101973519001RX','ES0031102366799001BC',"
        //        + "'ES0026000000599413DA','ES0031405426335001KW','ES0031405986446001CY','ES0031405997390001EP','ES0031406339748001NH',"
        //        + "'ES0021000002638406ZG','ES0021000005113267EL','ES0021000003239842NJ','ES0021000007668593XE','ES0021000004947846YZ',"
        //        + "'ES0022000004973677PY','ES0031101458680001XW','ES0031103131644001YF','ES0031405438352001PP','ES0031405197728001JF',"
        //        + "'ES0021000009357740JA','ES0021000005037011LP','ES0022000004980866KL','ES0022000005731645GD','ES0022000005731946VB',"
        //        + "'ES0022000007200475HJ','ES0031101583312001SQ','ES0031103269796001GT','ES0031104001483001DY','ES0031104001690001QY',"
        //        + "'ES0031405443788001PG','ES0031405526579001RC','ES0031405711700001YT','ES0275000000013505RZ','ES0031607597812001TB',"
        //        + "'ES0031609298736001YC','ES0031607400307001ZJ','ES0031607358968001RV','ES0031104001467001AZ','ES0031406292074001JG',"
        //        + "'ES0021000006668961HQ','ES0021000000178571ZL','ES0031102403016001SR','ES0021000009665762LD','ES0031446462969001MP',"
        //        + "'ES0031102084830001LA','ES0031609191099001QF','ES0031102503572002ZW')";

        //    //db = new MySQLDB(MySQLDB.Esquemas.MED);
        //    db = new MySQLDB(MySQLDB.Esquemas.AUX);
        //    command = new MySqlCommand(strSql, db.con);
        //    r = command.ExecuteReader();
        //    while (r.Read())
        //    {
        //        lista_cups.Add(r["CUPS"].ToString());
        //    }
        //    db.CloseConnection();
        //    red.GetCurvaRedShift(lista_cups, new DateTime(2019, 12, 14), new DateTime(2019, 12, 31), "F");
        //    formatos.P1(red.dic);
        //    MessageBox.Show("Proceso Terminado", "Telefonica P1", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //}

        //private bool GetCurvaRedShift(string cups13, DateTime fd, DateTime fh, string estado)
        //{
        //    servidores.RedShiftServer db;
        //    OdbcCommand command;
        //    OdbcDataReader r;
        //    string strSql;

        //    int numeroPeriodos = 0;
        //    int year = 0;
        //    int month = 0;
        //    int day = 0;

        //    int y = 0;

        //    int z = 0;
        //    DateTime fechaHora = new DateTime();

        //    try
        //    {
        //        GO.calendarios.UtilidadesCalendario util = new calendarios.UtilidadesCalendario();

        //        #region Query

        //        strSql = "SELECT cd_punto_med as cups15, cd_cups_ext as cups22, fh_lect_registro as fecha,  cd_estado_curva as estado ";
        //        // HORA ACTIVA
        //        for (int j = 1; j <= 25; j++)
        //            strSql += " ,NM_ENER_AC_H" + j + " as a" + j;

        //        // HORA REACTIVA
        //        for (int j = 1; j <= 25; j++)
        //            strSql += " ,NM_ENER_R1_H" + j + " as r" + j;

        //        // FUENTE HORA ACTIVA
        //        for (int j = 1; j <= 25; j++)
        //            strSql += " ,CD_FUENTE_HOR_AC_H" + j + " as fh" + j;

        //        // CUARTOHORARIA ACTIVA
        //        for (int j = 1; j <= 25; j++)
        //            for (int w = 1; w <= 4; w++)
        //            {
        //                z++;
        //                strSql += " ,NM_POT_AC_H" + j + "_CUAD" + w + " as v" + z;
        //            }


        //        // FUENTE CUARTOHORARIA ACTIVA
        //        for (int j = 1; j <= 25; j++)
        //            strSql += " ,CD_FUENTE_CUARTH_AC_H" + j + " as fch" + j;


        //        if (estado == "F")
        //            strSql += " from metra_owner.t_ed_h_curvas "
        //            + " WHERE cd_punto_med like '" + cups13 + "%'"
        //            + " and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
        //            + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
        //            + " and cd_estado_curva = '" + estado + "'"
        //            + " ORDER BY cd_punto_med, fh_lect_registro;";
        //        else
        //            strSql += " from metra_owner.t_ed_h_curvas "
        //            + " WHERE cd_punto_med = '" + cups13 + "'"
        //            + " and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
        //            + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
        //            + " and cd_estado_curva = '" + estado + "'"
        //            + " ORDER BY cd_punto_med, fh_lect_registro;";

        //        #endregion

        //        db = new servidores.RedShiftServer();
        //        command = new OdbcCommand(strSql, db.con);
        //        r = command.ExecuteReader();

        //        while (r.Read())
        //        {

        //            year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
        //            month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
        //            day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

        //            fechaHora = new DateTime(year, month, day, 0, 0, 0);
        //            numeroPeriodos = util.NumPeriodosHorarios(fechaHora);

        //            EndesaEntity.medida.CurvaCuartoHorariaInformes c = new EndesaEntity.medida.CurvaCuartoHorariaInformes();

        //            if (r["cups15"] != System.DBNull.Value)
        //                c.cups15 = r["cups15"].ToString();

        //            if (r["cups22"] != System.DBNull.Value)
        //                c.cups22 = r["cups22"].ToString();

        //            if (r["estado"] != System.DBNull.Value)
        //                c.estado = r["estado"].ToString() == "F" ? "FACTURADA" : "REGISTRADA";

        //            c.fecha = fechaHora;
        //            c.numPeriodos = numeroPeriodos;

        //            y = 0;
        //            for (int h = 1; h <= 25; h++)
        //            {
        //                if (r["a" + h] != System.DBNull.Value && r["a" + h].ToString() != "")
        //                    c.a[h] = Convert.ToDouble(r["a" + h]);

        //                if (r["fh" + h] != System.DBNull.Value && r["fh" + h].ToString() != "")
        //                    c.fa[h] = r["fh" + h].ToString().Trim();

        //                if (r["r" + h] != System.DBNull.Value && r["r" + h].ToString() != "")
        //                    c.r[h] = Convert.ToDouble(r["r" + h]);

        //                if (r["fch" + h] != System.DBNull.Value && r["fch" + h].ToString() != "")
        //                    c.fc[h] = r["fch" + h].ToString().Trim();
        //            }
        //            for (int cc = 1; cc <= 100; cc++)
        //                if (r["v" + cc] != System.DBNull.Value && r["v" + cc].ToString() != "")
        //                    c.value[cc] = Convert.ToDouble(r["v" + cc]);

        //            if (!list_cc.Exists(zz => zz.cups15 == c.cups15 && zz.fecha == c.fecha))
        //                list_cc.Add(c);

        //        }

        //        db.CloseConnection();
        //        return false;

        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return true;
        //    }


        //}
    }
}
