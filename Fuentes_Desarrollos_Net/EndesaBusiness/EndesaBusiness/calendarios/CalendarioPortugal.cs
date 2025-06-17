using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.calendarios
{
    public class CalendarioPortugal
    {
        Dictionary<DateTime, DateTime> dic_festivos;
        List<EndesaEntity.facturacion.Fact_pt_calendarios_Tabla> calendario;
        public Dictionary<DateTime, List<EndesaEntity.facturacion.FechaHoraPeriodo>> dic_periodos_tarifarios { get; set; }


        int numDias = 0;
        public CalendarioPortugal(int tipo_calendario, DateTime fd, DateTime fh)
        {
            calendario = new List<EndesaEntity.facturacion.Fact_pt_calendarios_Tabla>();
            CargaDefinicionCalendario(tipo_calendario, fd, fh);

            dic_festivos = new Dictionary<DateTime, DateTime>();
            CargaListaFestivos(fd, fh);

            numDias = Convert.ToInt32((fh - fd).TotalDays);
            dic_periodos_tarifarios = new Dictionary<DateTime, List<EndesaEntity.facturacion.FechaHoraPeriodo>>();
            CargaListaPeriodosTarifarios(fd, fh);

        }

        private void CargaListaPeriodosTarifarios(DateTime fd, DateTime fh)
        {
            DateTime dia_hora = new DateTime();
            DateTime d = fd;
            int numPeriodosCuartoHorarios = 0;
            bool firstOnly92 = true;
            bool firstOnly100 = true;

            for (int y = 0; y <= (fh - fd).TotalDays; y++)
            {
                dia_hora = d;
                numPeriodosCuartoHorarios = NumPeriodosCuartoHorarios(d);

                List<EndesaEntity.facturacion.FechaHoraPeriodo> lista_cuarto_horaria = new List<EndesaEntity.facturacion.FechaHoraPeriodo>();
                for (int j = 0; j < numPeriodosCuartoHorarios; j++)
                {

                    EndesaEntity.facturacion.FechaHoraPeriodo c = new EndesaEntity.facturacion.FechaHoraPeriodo();
                    c.fechahora = dia_hora;
                    c.periodoTarifario = GetPeriodoTarifa(dia_hora);
                    lista_cuarto_horaria.Add(c);
                    if (numPeriodosCuartoHorarios == 92 && (dia_hora.Hour == 1 && dia_hora.Minute == 45) && firstOnly92)
                    {
                        dia_hora = dia_hora.AddHours(1);
                        dia_hora = dia_hora.AddMinutes(15);
                        firstOnly92 = false;
                    }
                    else if (numPeriodosCuartoHorarios == 100 && (dia_hora.Hour == 2 && dia_hora.Minute == 45) && firstOnly100)
                    {
                        dia_hora = dia_hora.AddMinutes(15);
                        dia_hora = dia_hora.AddHours(-1);
                        firstOnly100 = false;
                    }
                    else
                        dia_hora = dia_hora.AddMinutes(15);
                }

                dic_periodos_tarifarios.Add(d, lista_cuarto_horaria);
                d = d.AddDays(1);
            }


        }


        public int GetPeriodoTarifa(DateTime f)
        {

            int dia = 0;
            EndesaEntity.facturacion.Fact_pt_calendarios_Tabla c;
            int hora = Convert.ToInt32(f.ToString("HH:mm").Replace(":", ""));

            if (EsFestivo(f) || (int)f.DayOfWeek == 0)
                dia = 0;
            else if ((int)f.DayOfWeek >= 1 && (int)f.DayOfWeek <= 5)
                dia = 1;
            else
                dia = 6;

            c = calendario.Find(z => z.dia == dia && z.estacion == Estacion(f)
                  && (z.hora_desde <= hora && z.hora_hasta >= (hora + 15)));

            return c.periodo;

        }

        private bool EsLaborable(DateTime f)
        {
            return (f.DayOfWeek >= DayOfWeek.Monday && f.DayOfWeek <= DayOfWeek.Friday);
        }

        private int NumPeriodosCuartoHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 92 : 96 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 100 : 96 : 96);
        }

        private void CargaDefinicionCalendario(int tipo_calendario, DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            

            try
            {
                strSql = "select dia, estacion, periodo, hora_desde, hora_hasta from fact_pt_calendarios"
                    + " where calendario_id = " + tipo_calendario + " and"
                    + " fd <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fh >= '" + fh.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.Fact_pt_calendarios_Tabla c = new EndesaEntity.facturacion.Fact_pt_calendarios_Tabla();
                    c.dia = Convert.ToInt32(r["dia"]);
                    c.estacion = r["estacion"].ToString();
                    c.periodo = Convert.ToInt32(r["periodo"]);
                    c.hora_desde = Convert.ToInt32(r["hora_desde"].ToString().Substring(0, 5).Replace(":", ""));
                    if (Convert.ToInt32(r["hora_hasta"].ToString().Substring(0, 5).Replace(":", "")) == 0)
                        c.hora_hasta = 2400;
                    else
                        c.hora_hasta = Convert.ToInt32(r["hora_hasta"].ToString().Substring(0, 5).Replace(":", ""));

                    calendario.Add(c);
                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "CargaDefinicionCalendario",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

        private bool EsFestivo(DateTime f)
        {
            DateTime o;
            return dic_festivos.TryGetValue(f.Date, out o);
        }

        private string Estacion(DateTime f)
        {
            string estacion = "";

            if (f > UltimoDomingoMarzo(f) && f < UltimoDomingoOctubre(f))
            {
                estacion = "VERANO";
            }
            else
            {
                estacion = "INVIERNO";
            }


            return estacion;
        }

        private DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private void CargaListaFestivos(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                strSql = "select fechafestivo from fact_pt_festivos where"
                + " FechaFestivo >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " FechaFestivo <= '" + fh.ToString("yyyy-MM-dd") + "';";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    dic_festivos.Add(Convert.ToDateTime(r["fechafestivo"]), Convert.ToDateTime(r["fechafestivo"]));
                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "CargaListaFestivos",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }


    }
}
