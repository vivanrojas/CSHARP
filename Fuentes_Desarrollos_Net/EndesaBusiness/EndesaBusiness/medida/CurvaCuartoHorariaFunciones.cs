using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class CurvaCuartoHorariaFunciones
    {
        public Dictionary<string, List<EndesaEntity.CurvaCuartoHoraria>> dic_cc { get; set; }

        public CurvaCuartoHorariaFunciones(List<string> lista_cups, DateTime fd, DateTime fh)
        {
            dic_cc = new Dictionary<string, List<EndesaEntity.CurvaCuartoHoraria>>();
            CargaMedidaCuartoHoraria(lista_cups, fd, fh);
        }

        public CurvaCuartoHorariaFunciones()
        {

        }

        private void CargaMedidaCuartoHoraria(List<string> lista_cups, DateTime fd, DateTime fh)
        {
            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            int j = 0;

            try
            {
                for (int i = 0; i < lista_cups.Count; i++)
                {
                    j++;
                    if (firstOnly)
                    {
                        sb.Append("select m.CCOUNIPS, m.Fecha, m.Version, m.Unidad, m.TotalA, m.TotalR");
                        for (int x = 1; x <= 25; x++)
                            sb.Append(" ,A" + x);

                        for (int x = 1; x <= 25; x++)
                            sb.Append(" ,R" + x);

                        sb.Append(" from medida_cuartohoraria_sce m where");
                        sb.Append(" m.CCOUNIPS in (");
                        sb.Append("'").Append(lista_cups[i]).Append("'");
                        firstOnly = false;
                    }
                    else
                    {
                        sb.Append(" ,'").Append(lista_cups[i]).Append("'");
                    }
                    if (j == 250)
                    {
                        sb.Append(")");
                        sb.Append(" and (m.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and m.Fecha <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");

                        j = 0;
                        firstOnly = true;
                        RunQuery(sb.ToString());
                        sb = null;
                        sb = new StringBuilder();
                    }

                }

                if (j > 0)
                {
                    sb.Append(")");
                    sb.Append(" and (m.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and m.Fecha <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                    j = 0;
                    firstOnly = true;
                    RunQuery(sb.ToString());
                    sb = null;

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "CurvaCuartoHorariaFunciones.CargaMedidaCuartoHoraria",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void RunQuery(string q)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(q, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {

                List<EndesaEntity.CurvaCuartoHoraria> c;
                EndesaEntity.CurvaCuartoHoraria cc;
                if (!dic_cc.TryGetValue(r["CCOUNIPS"].ToString(), out c))
                {

                    cc = new EndesaEntity.CurvaCuartoHoraria();
                    cc.totalA = Convert.ToInt32(r["TotalA"]);
                    cc.totalR = Convert.ToInt32(r["TotalR"]);

                    for (int i = 1; i <= 25; i++)
                    {
                        cc.a[i - 1] = Convert.ToInt32(r["A" + i]);
                        cc.r[i - 1] = Convert.ToInt32(r["R" + i]);
                    }

                    dic_cc.Add(r["CCOUNIPS"].ToString(), c);
                }
                else
                {

                }

            }
            db.CloseConnection();
        }

        // Importa extracción de curvas de carga del SCE
        public void ImportarCuartoHorariaPorLinea(string archivo)
        {

            int totalA = 0;
            int totalR = 0;
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            FileInfo file = new FileInfo(archivo);
            System.IO.StreamReader fileStream;
            string line;
            string[] c;
            int i = 0;
            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            string userID = Environment.UserName + "_" + DateTime.Now.ToString("HHmmss");

            forms.FrmProgressBar pb = new forms.FrmProgressBar();

            try
            {
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = this.NumLineas(archivo);
                pb.Text = "Importando " + file.Name;

                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                while ((line = fileStream.ReadLine()) != null)
                {
                    i++;
                    pb.progressBar.Increment(1);
                    pb.Refresh();
                    c = line.Split(';');

                    pb.txtDescripcion.Text = "Leyendo " + c[1] + " - " + FE(c[2]);
                    pb.progressBar.Increment(1);
                    pb.Refresh();



                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO medida_cuartohoraria_sce_temp ");
                        sb.Append("(CCOUNIPS, Fecha, Version, Unidad, TotalA, TotalR");
                        for (int j = 1; j <= 25; j++)
                        {
                            sb.Append(", ").Append("Value" + (j * 4 - 3));
                            sb.Append(", ").Append("Value" + (j * 4 - 2));
                            sb.Append(", ").Append("Value" + (j * 4 - 1));
                            sb.Append(", ").Append("Value" + (j * 4));
                            sb.Append(", ").Append("A" + j);
                            sb.Append(", ").Append("R" + j);
                            sb.Append(", ").Append("TESTME" + j);
                            sb.Append(", ").Append("INDCME" + j);
                        }
                        sb.Append(", Tipo, F_ULT_MOD, UserID) values ");
                        firstOnly = false;
                    }

                    totalA = 0;
                    totalR = 0;
                    for (int j = 1; j <= 25; j++)
                    {
                        int z = j * 9;
                        totalA += Convert.ToInt32(c[z + 4]);
                        totalR += Convert.ToInt32(c[z + 5]);
                    }


                    sb.Append("('").Append(c[1]).Append("', "); // CCOUNIPS
                    sb.Append(FE(c[2])).Append(", ");           // Fecha
                    sb.Append(c[3]).Append(", ");               // Version
                    sb.Append("'").Append(c[5]).Append("', ");  // Unidad
                    sb.Append(totalA).Append(", ");             // TotalA
                    sb.Append(totalR);                          // TotalR


                    for (int j = 1; j <= 25; j++)
                    {
                        int z = j * 9;
                        sb.Append(", ").Append(Convert.ToInt32(c[z]) / 1000);
                        sb.Append(", ").Append(Convert.ToInt32(c[z + 1]) / 1000);
                        sb.Append(", ").Append(Convert.ToInt32(c[z + 2]) / 1000);
                        sb.Append(", ").Append(Convert.ToInt32(c[z + 3]) / 1000);
                        sb.Append(", ").Append(Convert.ToInt32(c[z + 4]));
                        sb.Append(", ").Append(Convert.ToInt32(c[z + 5]));
                        sb.Append(", '").Append(c[z + 7]).Append("'");
                        sb.Append(", '").Append(c[z + 8]).Append("'");
                    }

                    sb.Append(", 1"); // Tipo
                    sb.Append(", '").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'");
                    sb.Append(", '").Append(userID).Append("'),");


                    if (i == 250)
                    {
                        pb.txtDescripcion.Text = "Guardando en Base de datos ...";
                        pb.Refresh();

                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }

                fileStream.Close();

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }


                this.VolcarTemp(userID);

                pb.txtDescripcion.Text = "Borrando tabla temporal";
                pb.Refresh();

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                strSql = "DELETE FROM medida_cuartohoraria_sce_temp where "
                + "UserID = '" + userID + "'";
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                pb.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "CurvaCuartoHorariaFunciones.ImportarCuartoHorariaPorLinea",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            }
        }


        private void VolcarTemp(string userID)
        {
            MySQLDB db;
            MySqlCommand command;

            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand("REPLACE INTO medida_cuartohoraria_sce select " +
                "CCOUNIPS,Fecha,Version,Unidad,TotalA,TotalR," +
                "Value1,Value2,Value3,Value4,Value5,Value6,Value7,Value8,Value9,Value10," +
                "Value11,Value12,Value13,Value14,Value15,Value16,Value17,Value18,Value19,Value20," +
                "Value21,Value22,Value23,Value24,Value25,Value26,Value27,Value28,Value29,Value30," +
                "Value31,Value32,Value33,Value34,Value35,Value36,Value37,Value38,Value39,Value40," +
                "Value41,Value42,Value43,Value44,Value45,Value46,Value47,Value48,Value49,Value50," +
                "Value51,Value52,Value53,Value54,Value55,Value56,Value57,Value58,Value59,Value60," +
                "Value61,Value62,Value63,Value64,Value65,Value66,Value67,Value68,Value69,Value70," +
                "Value71,Value72,Value73,Value74,Value75,Value76,Value77,Value78,Value79,Value80," +
                "Value81,Value82,Value83,Value84,Value85,Value86,Value87,Value88,Value89,Value90," +
                "Value91,Value92,Value93,Value94,Value95,Value96,Value97,Value98,Value99,Value100," +
                "A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17,A18,A19,A20,A21,A22,A23,A24,A25," +
                "R1,R2,R3,R4,R5,R6,R7,R8,R9,R10,R11,R12,R13,R14,R15,R16,R17,R18,R19,R20,R21,R22,R23,R24,R25," +
                "TESTME1,TESTME2,TESTME3,TESTME4,TESTME5,TESTME6,TESTME7,TESTME8,TESTME9,TESTME10," +
                "TESTME11,TESTME12,TESTME13,TESTME14,TESTME15,TESTME16,TESTME17,TESTME18,TESTME19,TESTME20" +
                ",TESTME21,TESTME22,TESTME23,TESTME24,TESTME25,INDCME1,INDCME2,INDCME3,INDCME4,INDCME5,INDCME6," +
                "INDCME7,INDCME8,INDCME9,INDCME10,INDCME11,INDCME12,INDCME13,INDCME14,INDCME15,INDCME16,INDCME17," +
                "INDCME18,INDCME19,INDCME20,INDCME21,INDCME22,INDCME23,INDCME24,INDCME25,Tipo,F_ULT_MOD " +
                "from medida_cuartohoraria_sce_temp where USERID = '" + userID + "';", db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }


        public void ImportarCuartoHorariaSCE(String archivo)
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql, strSql2;
            string userID;
            DateTime inicio = new DateTime();
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            try
            {
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = 4;
                pb.Text = "Importando " + archivo;


                inicio = DateTime.Now;

                userID = Environment.UserName + "_" + inicio.ToString("HHMMSS");
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                strSql = "LOAD DATA LOCAL INFILE '" + archivo.Replace(@"\", "\\\\")
                + "' REPLACE INTO TABLE medida_cuartohoraria_sce_temp "
                + "FIELDS TERMINATED BY ';' LINES TERMINATED BY '\\n' "
                + "(@kK,@CCOUNIPS,@FECHA,@VERSION,@X1,@X2,@X3,@X4";

                for (int i = 1; i <= 25; i++)
                {
                    strSql = strSql + ",@A" + i
                    + ",@AA" + i
                    + ",@AB" + i
                    + ",@AC" + i
                    + ",@AD" + i
                    + ",@TA" + i
                    + ",@TR" + i
                    + ",@B" + i
                    + ",@C" + i;
                }

                strSql = strSql + ",@X5,@X6,@X7,@X8,@X9,@X10,@X11,@X12,@X13,@X14,@X15,@X16,@X17,@X18,@X19,@20) SET "
                + "CCOUNIPS = @CCOUNIPS, "
                + "FECHA = concat(substr(@FECHA,1,4),'-',substr(@FECHA,5,2),'-',substr(@FECHA,7,2)), "
                + "VERSION = @VERSION, "
                + "Unidad =  @X2, "
                + "TotalA = (@TA1+@TA2+@TA3+@TA4+@TA5+@TA6+@TA7+@TA8+@TA9+@TA10+@TA11+@TA12+@TA13+@TA14+@TA15+@TA16+@TA17+@TA18+@TA19+@TA20+@TA21+@TA22+@TA23+@TA24+@TA25), "
                + "TotalR = (@TR1+@TR2+@TR3+@TR4+@TR5+@TR6+@TR7+@TR8+@TR9+@TR10+@TR11+@TR12+@TR13+@TR14+@TR15+@TR16+@TR17+@TR18+@TR19+@TR20+@TR21+@TR22+@TR23+@TR24+@TR25) ";
                for (int i = 1; i <= 25; i++)
                {
                    strSql = strSql + ",TESTME" + i + " = @A" + i
                    + ",INDCME" + i + " = @C" + i
                    + ",Value" + (i * 4 - 3) +
                        " = convert(@AA" + i + ",UNSIGNED INTEGER) / 1000 "
                    + ",Value" + (i * 4 - 2) +
                        " = convert(@AB" + i + ",UNSIGNED INTEGER) / 1000 "
                    + ",Value" + (i * 4 - 1) +
                        " = convert(@AC" + i + ",UNSIGNED INTEGER) / 1000 "
                    + ",Value" + (i * 4) +
                        " = convert(@AD" + i + ",UNSIGNED INTEGER) / 1000 "
                     + ",R" + i +
                        " = convert(@TR" + i + ",UNSIGNED INTEGER) "
                     + ",A" + i +
                        " = convert(@TA" + i + ",UNSIGNED INTEGER) ";
                }

                strSql = strSql + ",USERID = '" + userID + "'";

                pb.txtDescripcion.Text = "Borrando tabla temporal ...";
                pb.progressBar.Increment(1);
                pb.Refresh();

                strSql2 = "DELETE FROM medida_cuartohoraria_sce_temp where "
                + "UserID = '" + userID + "'";
                command = new MySqlCommand(strSql2, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                pb.txtDescripcion.Text = "Ejecutando LOAD DATA INFILE " + archivo;
                pb.progressBar.Increment(1);
                pb.Refresh();

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                pb.txtDescripcion.Text = "Ejecutando REPLACE ";
                pb.progressBar.Increment(1);
                pb.Refresh();

                this.VolcarTemp(userID);

                pb.txtDescripcion.Text = "Borrando tabla temporal";
                pb.progressBar.Increment(1);
                pb.Refresh();

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                strSql2 = "DELETE FROM medida_cuartohoraria_sce_temp where "
                + "UserID = '" + userID + "'";
                command = new MySqlCommand(strSql2, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                pb.Close();
                utilidades.Fichero.BorrarArchivo(archivo);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                 "CurvaCuartoHorariaFunciones.ImportarCuartoHorariaSCE",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

            }
        }

        private int NumLineas(string archivo)
        {
            FileInfo file = new FileInfo(archivo);
            System.IO.StreamReader fileStream;
            string line;
            int i = 0;

            try
            {
                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                while ((line = fileStream.ReadLine()) != null)
                {
                    i++;
                }
                fileStream.Close();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "CurvaCuartoHorariaFunciones.NumLineas",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            }

            return i;

        }
        private string FE(string f)
        {
            DateTime ff = new DateTime();
            if (f == "99999999")
                ff = new DateTime(2099, 12, 31);
            else
                ff = new DateTime(Convert.ToInt32(f.Substring(0, 4)), Convert.ToInt32(f.Substring(4, 2)), Convert.ToInt32(f.Substring(6, 2)));

            return "'" + ff.ToString("yyyy-MM-dd") + "'";

        }
    }
}
