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
    public class CurvaResumenFunciones
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", System.Environment.UserName);
        public Dictionary<string, EndesaEntity.medida.CurvaResumenTabla> dic_curva_resumen { get; set; }
        public medida.CurvaCuartoHorariaFunciones cc { get; set; }

        public CurvaResumenFunciones(List<string> lista_cups15, DateTime fd, DateTime fh)
        {
            dic_curva_resumen = new Dictionary<string, EndesaEntity.medida.CurvaResumenTabla>();
            CargaCurvasConsultaCurvaResumen(lista_cups15, fd, fh);
            // cc = new CurvaCuartoHorariaFunciones(lista_cups15, fd, fh);

        }

        public CurvaResumenFunciones()
        {
            dic_curva_resumen = new Dictionary<string, EndesaEntity.medida.CurvaResumenTabla>();
        }

        public void CargaResumen(string archivo)
        {
            string strSql = "";
            System.IO.StreamReader fileStream;
            string line;
            string[] c;
            int numLineas = 0;

            MySQLDB db;
            MySqlCommand command;

            try
            {
                FileInfo file = new FileInfo(archivo);
                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                while ((line = fileStream.ReadLine()) != null)
                {
                    c = line.Split(';');
                    if (numLineas == 0)
                    {
                        strSql = "REPLACE INTO med_resumen_curvas (CPUNTMED,FFACTDES,FFACTHAS,VSECRCUR,FLECTREG,VTARLTP,VCCRECIB,"
                            + "TESTRCUR,TFUENTEH,TFUENTEC,TESTVALD,VEACONTO,VERCONTO,TINDOBJ,FACTURA) "
                            + "VALUES ('" + c[1] + "', " + FE(c[2]) + ", " + FE(c[3]) + ", " + Convert.ToInt32(c[4])
                            + "," + FE(c[5]) + ", " + Convert.ToInt32(c[6]) + ", " + Convert.ToInt32(c[7]) + ", '" + c[8] + "', ";

                        if (c[9].Trim() != "")
                            strSql += Convert.ToInt32(c[9]) + ", ";
                        else
                            strSql += "null" + ", ";

                        if (c[10].Trim() != "")
                            strSql += Convert.ToInt32(c[10]) + ", ";
                        else
                            strSql += "null" + ", ";

                        if (c[11].Trim() != "")
                            strSql += Convert.ToInt32(c[11]) + ", ";
                        else
                            strSql += "null" + ", ";

                        strSql += Convert.ToDouble(c[12]) + ", " + Convert.ToDouble(c[13]) + ", " + "'" + c[14] + "', " + "'" + c[15] + "')";
                    }
                    else
                    {
                        strSql += ",('" + c[1] + "'," + FE(c[2]) + ", " + FE(c[3]) + ", " + Convert.ToInt32(c[4]) + ", ";
                        strSql += FE(c[5]) + ", " + Convert.ToInt32(c[6]) + ", " + Convert.ToInt32(c[7]) + ",'" + c[8] + "', ";

                        if (c[9].Trim() != "")
                            strSql += Convert.ToInt32(c[9]) + ", ";
                        else
                            strSql += "null" + ", ";


                        if (c[10].Trim() != "")
                            strSql += Convert.ToInt32(c[10]) + ", ";
                        else
                            strSql += "null" + ", ";


                        if (c[11].Trim() != "")
                            strSql += Convert.ToInt32(c[11]) + ", ";
                        else
                            strSql += "null" + ", ";

                        strSql += Convert.ToDouble(c[12]) + ", " + Convert.ToDouble(c[13]) + ", " + "'" + c[14] + "', " + "'" + c[15] + "')";
                    }
                    numLineas++;

                    if (numLineas == 500)
                    {
                        numLineas = 0;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(strSql, db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                    }

                }

                fileStream.Close();
                if (numLineas > 0)
                {
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }

                MessageBox.Show("Carga completada satisfactoriamente.",
                  "Carga Curva Resumen",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "CurvaResumenFunciones.GetCurvas",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);

                ficheroLog.Add("CurvaResumenFunciones.CargaResumen: " + strSql + " --> " + e.Message);
            }

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

        private void ConsultaCurvaResumen(List<string> lista_cups15, DateTime fd, DateTime fh, string estado)
        {

            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            int j = 0;

            try
            {

                for (int i = 0; i < lista_cups15.Count; i++)
                {
                    j++;
                    if (firstOnly)
                    {
                        sb.Append("select r.CPUNTMED,r.FFACTDES,r.FFACTHAS,r.VSECRCUR,r.TESTRCUR,r.VCCRECIB,r.VEACONTO,r.VERCONTO, f.FUENTEMADRE as fuente");
                        sb.Append(" from med.med_resumen_curvas r");
                        sb.Append(" left outer join med.med_fuentesmedida f on");
                        sb.Append(" f.NUM = r.TFUENTEH");
                        sb.Append(" where  r.CPUNTMED in (");
                        sb.Append("'").Append(lista_cups15[i]).Append("'");
                        firstOnly = false;
                    }
                    else
                        sb.Append(",'").Append(lista_cups15[i]).Append("'");


                    if (j == 250)
                    {
                        sb.Append(")");
                        sb.Append(" and (r.FFACTDES >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and r.FFACTHAS <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                        sb.Append(" and r.TESTRCUR = '").Append(estado).Append("'");
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
                    sb.Append(" and (r.FFACTDES >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and r.FFACTHAS <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                    sb.Append(" and r.TESTRCUR = '").Append(estado).Append("'");
                    j = 0;
                    firstOnly = true;
                    RunQuery(sb.ToString());
                    sb = null;
                    sb = new StringBuilder();
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "CurvaResumenFunciones.ConsultaCurvaResumen",
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
                EndesaEntity.medida.CurvaResumenTabla re = new EndesaEntity.medida.CurvaResumenTabla();
                re.cpuntmed = r["CPUNTMED"].ToString();
                re.fecha_desde = Convert.ToDateTime(r["FFACTDES"]);
                re.fecha_hasta = Convert.ToDateTime(r["FFACTHAS"]);
                re.version = Convert.ToInt32(r["VSECRCUR"]);
                re.estado = DescripcionEstadoCurva(r["TESTRCUR"].ToString());
                re.dias = Convert.ToInt32(r["VCCRECIB"]);
                re.activa = Convert.ToInt32(r["VEACONTO"]);
                re.reactiva = Convert.ToInt32(r["VERCONTO"]);
                re.fuente = r["fuente"].ToString();

                EndesaEntity.medida.CurvaResumenTabla rr;
                if (!dic_curva_resumen.TryGetValue(re.cpuntmed, out rr))
                    dic_curva_resumen.Add(re.cpuntmed, re);
            }

            db.CloseConnection();
        }

        private void GetCurvas(DateTime fd, DateTime fh, string estado, bool cupsCorto)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "select r.CPUNTMED,r.FFACTDES,r.FFACTHAS,r.VSECRCUR,r.TESTRCUR,r.VCCRECIB,r.VEACONTO,r.VERCONTO, f.FUENTEMADRE as fuente"
                    + " from med.med_resumen_curvas r"
                    + " left outer join med.med_fuentesmedida f on"
                    + " f.NUM = r.TFUENTEH"
                    + " where";
                if (cupsCorto)
                    strSql += " length(r.CPUNTMED) = 13 and (r.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and r.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')"
                        + " and r.TESTRCUR = '" + estado + "'";
                else
                    strSql += " (r.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and r.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')"
                        + " and r.TESTRCUR = '" + estado + "'";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.CurvaResumenTabla re = new EndesaEntity.medida.CurvaResumenTabla();
                    re.cpuntmed = r["CPUNTMED"].ToString().Substring(0, 13);
                    re.fecha_desde = Convert.ToDateTime(r["FFACTDES"]);
                    re.fecha_hasta = Convert.ToDateTime(r["FFACTHAS"]);
                    re.version = Convert.ToInt32(r["VSECRCUR"]);
                    re.estado = DescripcionEstadoCurva(r["TESTRCUR"].ToString());
                    re.dias = Convert.ToInt32(r["VCCRECIB"]);
                    re.activa = Convert.ToInt32(r["VEACONTO"]);
                    re.reactiva = Convert.ToInt32(r["VERCONTO"]);
                    re.fuente = r["fuente"].ToString();

                    EndesaEntity.medida.CurvaResumenTabla rr;
                    if (!dic_curva_resumen.TryGetValue(re.cpuntmed, out rr))
                        dic_curva_resumen.Add(re.cpuntmed, re);
                }

                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "CurvaResumenFunciones.GetCurvas",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        private string DescripcionEstadoCurva(string e)
        {
            string estado = "";

            switch (e)
            {
                case "F":
                    estado = "FACTURADA";
                    break;
                case "R":
                    estado = "REGISTRADA";
                    break;
                default:
                    estado = "SIN CURVA REGISTRADA";
                    break;
            }

            return estado;
        }

        private void CargaCurvasConsultaCurvaResumen(List<string> listaCups15, DateTime fd, DateTime fh)
        {
            try
            {
                ConsultaCurvaResumen(listaCups15, fd, fh, "F");
                ConsultaCurvaResumen(listaCups15, fd, fh, "R");

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "CurvaResumenFunciones.CargaCurvasConsultaCurvaResumen",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private void CargaCurvas(DateTime fd, DateTime fh)
        {
            try
            {
                GetCurvas(fd, fh, "F", true);
                GetCurvas(fd, fh, "R", true);
                GetCurvas(fd, fh, "F", false);
                GetCurvas(fd, fh, "R", false);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "CurvaResumenFunciones.CargaCurvas",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }

        }

        public EndesaEntity.medida.CurvaResumenTabla GetCurvaResumen(string cups)
        {
            EndesaEntity.medida.CurvaResumenTabla rr = new EndesaEntity.medida.CurvaResumenTabla();
            try
            {
                if (dic_curva_resumen.TryGetValue(cups, out rr))
                {

                }


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "CurvaResumenFunciones.GetCurvaResumen",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
            return rr;
        }
    }
}
