using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class CurvaResumenSCE : EndesaEntity.medida.CurvaResumenTabla
    {
        Dictionary<string, List<EndesaEntity.medida.CurvaResumenTabla>> dic;

        public CurvaResumenSCE(List<string> lista_cups15, DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, List<EndesaEntity.medida.CurvaResumenTabla>>();
            ConsultaCurvaResumen(lista_cups15, fd, fh);
        }


        private void ConsultaCurvaResumen(List<string> listaCups15, DateTime fd, DateTime fh)
        {
            try
            {
                BuscaCurvaResumenFacturado(listaCups15, fd, fh);
                BuscaCurvaResumenRegistrado(listaCups15, fd, fh);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "CurvaHorariaSCE.CargaCurvasConsultaCurvaResumen",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private void BuscaCurvaResumenFacturado(List<string> lista_cups15, DateTime fd, DateTime fh)
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
                        sb.Append("select r.CPUNTMED, min(r.FFACTDES) as FFACTDES ,max(r.FFACTHAS) as FFACTHAS, r.VSECRCUR, r.TESTRCUR, sum(r.VCCRECIB) as VCCRECIB,");
                        sb.Append("sum(r.VEACONTO) as VEACONTO, sum(r.VERCONTO) as VERCONTO, f.FUENTEMADRE as fuente");
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
                        sb.Append(" and r.TESTRCUR = '").Append("F").Append("'  group by r.CPUNTMED, MONTH(r.FFACTDES)");
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
                    sb.Append(" and r.TESTRCUR = '").Append("F").Append("' group by r.CPUNTMED, MONTH(r.FFACTDES)");
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
                "CurvaHorariaSCE.ConsultaCurvaResumen",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        private void BuscaCurvaResumenRegistrado(List<string> lista_cups15, DateTime fd, DateTime fh)
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
                        sb.Append("select r.CPUNTMED, min(r.FFACTDES) as FFACTDES ,max(r.FFACTHAS) as FFACTHAS, r.VSECRCUR, r.TESTRCUR, sum(r.VCCRECIB) as VCCRECIB,");
                        sb.Append("sum(r.VEACONTO) as VEACONTO, sum(r.VERCONTO) as VERCONTO, f.FUENTEMADRE as fuente");
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
                        sb.Append(" and r.TESTRCUR = '").Append("R").Append("' group by r.CPUNTMED, MONTH(r.FFACTDES)");
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
                    sb.Append(" and r.TESTRCUR = '").Append("R").Append("' group by r.CPUNTMED, MONTH(r.FFACTDES) ");
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
                "CurvaHorariaSCE.ConsultaCurvaResumen",
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

                List<EndesaEntity.medida.CurvaResumenTabla> rr;
                if (!dic.TryGetValue(re.cpuntmed, out rr))
                {
                    List<EndesaEntity.medida.CurvaResumenTabla> lista = new List<EndesaEntity.medida.CurvaResumenTabla>();
                    lista.Add(re);
                    dic.Add(re.cpuntmed, lista);
                }
                else
                    rr.Add(re);

            }

            db.CloseConnection();
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

        public void GetCurva(string cups15, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.medida.CurvaResumenTabla> lista;
            if (dic.TryGetValue(cups15, out lista))
            {
                List<EndesaEntity.medida.CurvaResumenTabla> sublista = lista.Where(z => (z.fecha_desde >= fd && z.fecha_hasta <= fh)).ToList();
                if (sublista.Count > 0)
                    for (int i = 0; i < sublista.Count(); i++)
                    {
                        this.existe_curva = true;
                        this.estado = sublista[i].estado;
                        this.fecha_desde = sublista[i].fecha_desde;
                        this.fecha_hasta = sublista[i].fecha_hasta;
                        this.activa = sublista[i].activa;
                        this.reactiva = sublista[i].reactiva;
                        this.fuente = sublista[i].fuente;
                        this.dias = sublista[i].dias;

                    }


            }
            else
            {

                this.existe_curva = false;

                this.dias = 0;

            }


        }

        public DataTable GetCurvaDGV(string cups13, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            try
            {
                strSql = "select r.CPUNTMED,r.FFACTDES,r.FFACTHAS,r.VSECRCUR,r.TESTRCUR,r.VCCRECIB,r.VEACONTO,r.VERCONTO, f.FUENTEMADRE as fuente"
                    + " from med.med_resumen_curvas r"
                    + " left outer join med.med_fuentesmedida f on"
                    + " f.NUM = r.TFUENTEH"
                    + " where r.CPUNTMED like '" + cups13 + "%' and"
                    + " (r.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " r.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "');";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                MySqlDataAdapter da = new MySqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                return dt;
            }
            catch(Exception e)
            {
                return null;
            }
 
        }

    }
}
