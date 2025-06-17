using EndesaBusiness.medida.Redshift;
using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.adif
{
    public class NuevaMedidaADIF : EndesaEntity.medida.CurvaResumenTabla
    {

        utilidades.Fechas fechas;
        logs.Log ficheroLog;
        public Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> dic { get; set; }

        public Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> dic_activa { get; set; }
        public Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>> dic_reactiva { get; set; }
        public Dictionary<string, List<CurvaCuartoHoraria>> dic_cc { get; set; }

        public NuevaMedidaADIF()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_MedidaADIF");
            dic_cc = new Dictionary<string, List<CurvaCuartoHoraria>>();
            fechas = new Fechas();
        }

        public NuevaMedidaADIF(List<string> listacups20, DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>>();
            dic_activa = new Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>>();
            dic_reactiva = new Dictionary<string, List<EndesaEntity.medida.CurvaDeCarga>>();

            ConsultaMedidaProporcionadaPorAdif("A", fd, fh, listacups20);
            ConsultaMedidaProporcionadaPorAdif("R", fd, fh, listacups20);

        }

       


        public bool CargaMedidaHoraria(List<string> lista, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            int numeroPeriodos = 0;
            int num_horas2 = 0;
           
           
            DateTime fechaHora = new DateTime();
            bool firstOnly = true;


            // PUNTOS DE MEDIDA            
            EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes pm = 
                new EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes(lista);


            try
            {
                strSql = "SELECT CUPSREE, FECHAHORA, ESTACION, AE, AES, R1, R2, R3, R4, FechaCarga"
                    + " FROM adif_PO1011"                   
                    + " WHERE CUPSREE in (";

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
                + " (FECHAHORA > '" + fd.ToString("yyyy-MM-dd") + "'"
                + " AND FECHAHORA <= '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "')"
                + " ORDER BY CUPSREE, FECHAHORA, ESTACION DESC, FechaCarga";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);                
                r = command.ExecuteReader();

                while (r.Read())
                {

                    CurvaCuartoHoraria c = new CurvaCuartoHoraria();

                    fechaHora = Convert.ToDateTime(r["FECHAHORA"]);

                    /* Si hay 2 horas 2 a la segunda hora se pone hora 3 */
                    if (fechaHora.Hour == 2)
                        num_horas2++;
                    else
                        num_horas2 = 0;

                    if (num_horas2 != 2)
                    {
                        fechaHora = fechaHora.AddHours(-1);
                        c.estacion = Convert.ToInt32(r["ESTACION"]);
                    }
                    else
                        c.estacion = 2;


                    numeroPeriodos = fechas.NumPeriodosHorarios(fechaHora);
                                                           

                    if (r["CUPSREE"] != System.DBNull.Value)
                    {
                        c.cups20 = r["CUPSREE"].ToString();
                        c.cups22 = pm.GetCUPS22(r["CUPSREE"].ToString());
                    }
                    
                    c.fecha = fechaHora;
                    c.numPeriodos = numeroPeriodos;

                    
                    c.AE = Convert.ToDouble(r["AE"]);
                    c.AES = Convert.ToDouble(r["AES"]);
                    c.R1 = Convert.ToDouble(r["R1"]);
                    c.R2 = Convert.ToDouble(r["R2"]);
                    c.R3 = Convert.ToDouble(r["R3"]);
                    c.R4 = Convert.ToDouble(r["R4"]);
                    c.fecha_carga = Convert.ToDateTime(r["FechaCarga"]);
                                        
                    List<CurvaCuartoHoraria> o;
                    if (!dic_cc.TryGetValue(c.cups20, out o))                    
                    {

                        o = new List<CurvaCuartoHoraria>();
                        o.Add(c);
                        dic_cc.Add(c.cups20, o);
                    }
                    else
                    {
                        if (!o.Exists(zz => zz.cups20 == c.cups20 && zz.fecha == c.fecha && zz.estacion == c.estacion))
                            o.Add(c);
                        else
                        {
                            for (int i = 0; i < o.Count; i++)
                            {
                                if (o[i].fecha == c.fecha && o[i].estacion == c.estacion && o[i].fecha_carga < c.fecha_carga)
                                    o[i] = c;
                            }
                        }
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

        private void ConsultaMedidaProporcionadaPorAdif(string tipoEnergia, DateTime fd, DateTime fh, List<string> listacups20)
        {

            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();

            int j = 0;

            try
            {
                for (int i = 0; i < listacups20.Count; i++)
                {
                    j++;
                    if (firstOnly)
                    {                        
                        sb.Append("select a.CUPS20 , a.Total as total ,a.FechaCarga, a.Fecha");

                        for (int x = 1; x <= 25; x++)
                            sb.Append(" ,a.Value" + x);

                        sb.Append(" from adif_medida_horaria_adif a");                        
                        sb.Append(" inner join (select cups20, max(FechaCarga)FechaCarga, Fecha from adif_medida_horaria_adif b where");
                        sb.Append(" (b.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and b.Fecha <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                        sb.Append(" and b.TipoEnergia = '").Append(tipoEnergia).Append("' group by b.cups20, month(Fecha)) b on");
                        sb.Append(" b.cups20 = a.cups20 and");
                        sb.Append(" b.FechaCarga = a.FechaCarga");
                        sb.Append(" where");
                        sb.Append(" a.CUPS20 in (");
                        sb.Append("'").Append(listacups20[i]).Append("'");

                        firstOnly = false;
                    }
                    else
                        sb.Append(" ,'").Append(listacups20[i]).Append("'");

                    if (j == 250)
                    {
                        sb.Append(")");
                        sb.Append(" and (a.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' AND a.Fecha <= '");
                        sb.Append(fh.ToString("yyyy-MM-dd")).Append("') AND a.TipoEnergia = '").Append(tipoEnergia).Append("'");
                        sb.Append(" group by a.cups20, a.Fecha order by a.cups20, a.Fecha");
                        j = 0;
                        firstOnly = true;
                        this.RunQuery(fd, fh, tipoEnergia, sb.ToString());
                        sb = null;
                        sb = new StringBuilder();
                    }

                }

                if (j > 0)
                {
                    sb.Append(")");
                    sb.Append(" and (a.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' AND a.Fecha <= '");
                    sb.Append(fh.ToString("yyyy-MM-dd")).Append("') AND a.TipoEnergia = '").Append(tipoEnergia).Append("'");
                    sb.Append(" group by a.cups20, a.Fecha order by a.cups20, a.Fecha");
                    j = 0;
                    firstOnly = true;
                    this.RunQuery(fd, fh, tipoEnergia, sb.ToString());
                    sb = null;
                    sb = new StringBuilder();
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                 "InformeErse.ConsultaMedida",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }
        }

        private void RunQuery(DateTime fd, DateTime fh, string tipoEnergia, string q)
        {
            bool esUltimoDomingoMarzo = false;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            utilidades.Fechas util_fecha = new utilidades.Fechas();

            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(q, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {

                EndesaEntity.medida.CurvaDeCarga c = new EndesaEntity.medida.CurvaDeCarga();
                c.fecha = Convert.ToDateTime(r["fecha"]);
                c.numPeriodosHorarios = util_fecha.NumPeriodosHorarios(c.fecha);
                c.numPeriodosCuartoHorarios = c.numPeriodosHorarios * 4;


                esUltimoDomingoMarzo = c.fecha == util_fecha.UltimoDomingoMarzo(c.fecha);

                if (tipoEnergia == "A")
                {
                    if (r["Total"] != System.DBNull.Value)
                        c.total_activa = Convert.ToDouble(r["Total"]);

                    for (int i = 1; i < 26; i++)
                    {

                        if (esUltimoDomingoMarzo && i == 2)
                        {
                            if (r["Value3"] != System.DBNull.Value)
                                c.horaria_activa[i - 1] = Convert.ToDouble(r["Value3"]);
                        }
                        else if (esUltimoDomingoMarzo && i == 3)
                        {
                            c.horaria_activa[i - 1] = 0;
                        }
                        else
                        {
                            if (r["Value" + i] != System.DBNull.Value)
                                c.horaria_activa[i - 1] = Convert.ToDouble(r["Value" + i]);
                        }

                    }
                }
                else
                {
                    if (r["Total"] != System.DBNull.Value)
                        c.total_reactiva = Convert.ToDouble(r["Total"]);

                    for (int i = 1; i < 26; i++)
                    {
                        if (r["Value" + i] != System.DBNull.Value)
                            c.horaria_reactiva[i - 1] = Convert.ToDouble(r["Value" + i]);
                    }
                }



                List<EndesaEntity.medida.CurvaDeCarga> o;
                if (!dic.TryGetValue(r["CUPS20"].ToString(), out o))
                {
                    List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
                    lista.Add(c);
                    dic.Add(r["cups20"].ToString(), lista);
                }
                else
                {
                    o.Add(c);
                }

                if(tipoEnergia == "A")
                {
                    if (!dic_activa.TryGetValue(r["CUPS20"].ToString(), out o))
                    {
                        List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
                        lista.Add(c);
                        dic_activa.Add(r["cups20"].ToString(), lista);
                    }
                    else
                    {
                        o.Add(c);
                    }
                }
                else
                {
                    if (!dic_reactiva.TryGetValue(r["CUPS20"].ToString(), out o))
                    {
                        List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
                        lista.Add(c);
                        dic_reactiva.Add(r["cups20"].ToString(), lista);
                    }
                    else
                    {
                        o.Add(c);
                    }
                }



            }
            db.CloseConnection();
        }
        public List<EndesaEntity.medida.CurvaDeCarga> GetCurva(string cups, DateTime fd, DateTime fh)
        {


            List<EndesaEntity.medida.CurvaDeCarga> o;
            if (dic.TryGetValue(cups, out o))
            {
                List<EndesaEntity.medida.CurvaDeCarga> matches = o.Where(z => z.fecha >= fd && z.fecha <= fh).ToList();
                return matches;
            }
            else
                return null;
        }


        public List<CurvaCuartoHoraria> GetCurvaPO1011(string cups, DateTime fd, DateTime fh)
        {
            List<CurvaCuartoHoraria> o;
            if (dic_cc.TryGetValue(cups, out o))
            {
                List<CurvaCuartoHoraria> matches = o.Where(z => z.fecha >= fd && z.fecha < fh.AddDays(1)).ToList();
                return matches;
            }
            else
                return null;
        }

        public List<CurvaCuartoHoraria> GetCurvaPO1011_Neteada(string cups, DateTime fd, DateTime fh)
        {
            
            List<CurvaCuartoHoraria> matches = new List<CurvaCuartoHoraria>();

            List<CurvaCuartoHoraria> o;
            if (dic_cc.TryGetValue(cups, out o))            
                matches = o.Where(z => z.fecha >= fd && z.fecha < fh.AddDays(1)).ToList();


            if (matches.Count > 0)
            {

                for (int i = 0; i < matches.Count(); i++)
                {
                    if (matches[i].AE - matches[i].AES < 0)
                        matches[i].NETEADA = 0;
                    else
                        matches[i].NETEADA = matches[i].AE - matches[i].AES;
                }

                return matches;   
                

            } else
                return null;
            
                
        }

        public void BuscaCurva(string cups, DateTime fd, DateTime fh)
        {
            this.existe_curva = false;
            this.activa = 0;
            this.reactiva = 0;

            List<EndesaEntity.medida.CurvaDeCarga> o;
            if (dic.TryGetValue(cups, out o))
            {
                this.existe_curva = false;
                List<EndesaEntity.medida.CurvaDeCarga> matches = o.Where(z => z.fecha >= fd && z.fecha <= fh).ToList();
                for (int i = 0; i < matches.Count(); i++)
                {
                    this.existe_curva = true;
                    this.activa += Convert.ToInt32(matches[i].total_activa);
                    this.reactiva += Convert.ToInt32(matches[i].total_reactiva);
                }

            }
        }



        public Double TotalActiva(string cups, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.medida.CurvaDeCarga> o;
            if (dic.TryGetValue(cups, out o))
            {
                this.existe_curva = true;
                List<EndesaEntity.medida.CurvaDeCarga> matches = o.Where(z => z.fecha >= fd && z.fecha <= fh).ToList();
                for (int i = 0; i < matches.Count(); i++)
                    total += matches[i].total_activa;
            }

            return total;
        }

               

        public double TotalActivaEntrante(string cups, DateTime fd, DateTime fh)
        {
            double total = 0;
            this.existe_curva = false;
            List<EndesaEntity.CurvaCuartoHoraria> o;
            if (dic_cc.TryGetValue(cups, out o))
            {
                this.existe_curva = true;
                List<EndesaEntity.CurvaCuartoHoraria> matches = o.Where(z => z.fecha >= fd && z.fecha < fh.AddDays(1)).ToList();
                for (int i = 0; i < matches.Count(); i++)
                    total += matches[i].AE;
            }

            return total;
        }

        

        public double TotalActivaSaliente(string cups, DateTime fd, DateTime fh)
        {
            double total = 0;
            this.existe_curva = false;
            List<EndesaEntity.CurvaCuartoHoraria> o;
            if (dic_cc.TryGetValue(cups, out o))
            {
                this.existe_curva = true;
                List<EndesaEntity.CurvaCuartoHoraria> matches = o.Where(z => z.fecha >= fd && z.fecha < fh.AddDays(1)).ToList();
                for (int i = 0; i < matches.Count(); i++)
                    total += matches[i].AES;
            }

            return total;
        }

        public double TotalReactivaR4(string cups, DateTime fd, DateTime fh)
        {
            double total = 0;
            this.existe_curva = false;
            List<EndesaEntity.CurvaCuartoHoraria> o;
            if (dic_cc.TryGetValue(cups, out o))
            {
                this.existe_curva = true;
                List<EndesaEntity.CurvaCuartoHoraria> matches = o.Where(z => z.fecha >= fd && z.fecha < fh.AddDays(1)).ToList();
                for (int i = 0; i < matches.Count(); i++)
                    total += matches[i].R4;
            }

            return total;
        }

        public double Neteo(string cups, DateTime fd, DateTime fh)
        {
            double total = 0;
            this.existe_curva = false;
            List<EndesaEntity.CurvaCuartoHoraria> o;
            if (dic_cc.TryGetValue(cups, out o))
            {
                this.existe_curva = true;
                List<EndesaEntity.CurvaCuartoHoraria> matches = o.Where(z => z.fecha >= fd && z.fecha < fh.AddDays(1)).ToList();
                for (int i = 0; i < matches.Count(); i++)
                {
                    if (matches[i].AE - matches[i].AES < 0)
                        total += 0;
                    else
                        total += matches[i].AE - matches[i].AES;
                }
                    
            }

            return total;
        }

        public Double TotalReactiva(string cups, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.medida.CurvaDeCarga> o;
            if (dic.TryGetValue(cups, out o))
            {
                List<EndesaEntity.medida.CurvaDeCarga> matches = o.Where(z => z.fecha >= fd && z.fecha <= fh).ToList();
                for (int i = 0; i < matches.Count(); i++)
                    total += matches[i].total_reactiva;
            }

            return total;
        }
    }
}
