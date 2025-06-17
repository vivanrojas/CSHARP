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
public class CargaCalendarios
    {
        List<EndesaEntity.facturacion.TarifaTerritorio> lTT = new List<EndesaEntity.facturacion.TarifaTerritorio>();
        List<DateTime> lf = new List<DateTime>(); // lista festivos
        List<EndesaEntity.Calendario> lc = new List<EndesaEntity.Calendario>();
        List<EndesaEntity.facturacion.CalendarioTarifario> lpc = new List<EndesaEntity.facturacion.CalendarioTarifario>();

        public CargaCalendarios()
        {
            DateTime fd = new DateTime(2014, 01, 01);
            DateTime fh = new DateTime(2099, 12, 31);
            CargaListaTarifas(fd, fh);
            CargaListaFestivos(fd, fh);
            CargaDefinicionCalendario(fd, fh);
            CargaListaPeriodosTarifarios(fd, fh);
        }

        public CargaCalendarios(DateTime fromDate, DateTime toDate)
        {
            CargaListaTarifas(fromDate, toDate);
            CargaListaFestivos(fromDate, toDate);
            CargaDefinicionCalendario(fromDate, toDate);
            CargaListaPeriodosTarifarios(fromDate, toDate);
        }

        public CargaCalendarios(DateTime fromDate, DateTime toDate, string tarifa, string territorio)
        {
            CargaListaTarifas(fromDate, toDate);
            CargaListaFestivos(fromDate, toDate);
            CargaDefinicionCalendario(fromDate, toDate);
            CargaListaPeriodosTarifarios(fromDate, toDate);
        }


        private void CargaListaPeriodosTarifarios(DateTime fd, DateTime fh)
        {
            int z = 0;

            for (int i = 0; i < lTT.Count(); i++)
            {
                EndesaEntity.facturacion.CalendarioTarifario ct = new EndesaEntity.facturacion.CalendarioTarifario();
                ct.tarifa = lTT[i].tarifa;
                ct.territorio = lTT[i].territorio;

                z = 0;
                DateTime d = fd;
                for (int y = 0; y <= (fh - fd).TotalDays; y++)
                {
                    z++;
                    EndesaEntity.FechaTarifa ft = new EndesaEntity.FechaTarifa();
                    ft.f = d;
                    for (int j = 1; j <= NumPeriodosHorarios(d); j++)
                    {
                        ft.pt[j] = GetPeriodoTarifa(ct.tarifa, ct.territorio, d, j);
                    }

                    ct.ft.Add(ft);

                    d = d.AddDays(1);
                }

                lpc.Add(ct);
            }
        }

        public int NumPeriodosTarifarios(DateTime fd, DateTime fh, string tarifa, string cups13, int pt)
        {
            List<EndesaEntity.facturacion.CalendarioTarifario> vlpc = new List<EndesaEntity.facturacion.CalendarioTarifario>();
            List<EndesaEntity.FechaTarifa> ft = new List<EndesaEntity.FechaTarifa>();
            int p = 0;

            vlpc = lpc.FindAll(z => z.tarifa == tarifa && z.territorio == TerritorioDesdeCUPS13(cups13));
            for (int i = 0; i <= vlpc.Count(); i++)
            {
                ft = vlpc[i].ft.FindAll(z => z.f >= fd && z.f <= fh);
                for (int j = 0; i <= ft.Count(); j++)
                {
                    for (int x = 1; x <= 25; x++)
                    {
                        if (ft[j].pt[x] == pt)
                        {
                            p++;
                        }
                    }

                }
            }

            return p;
        }

        private int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }
        public int GetPeriodoTarifa(string tarifa, string territorio, DateTime f, int ph)
        {
            bool buscarPeriodoTarifario = false;
            bool laborable = false;

            int p = 0;
            EndesaEntity.Calendario c = new EndesaEntity.Calendario();


            
            switch (tarifa)
            {
                case "2.0TD":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "3.0TD":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "6.1TD":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "6.2TD":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "6.3TD":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "6.4TD":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "3.0TDVE":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
                case "6.1TDVE":
                    buscarPeriodoTarifario = true;
                    laborable = true;
                    break;
            }

            if (buscarPeriodoTarifario)
            {
                // Para las tarifas 2, 3 y 6 si es festivo o fin de semana es P6
                if ((tarifa.ToUpper().Contains("6") || tarifa.ToUpper().Contains("3")) && 
                    (!EsLaborable(f) || EsFestivo(f)))
                {
                    p = 6;
                }
                else if (tarifa.Substring(0,1) == "2" && !EsLaborable(f) || EsFestivo(f))
                {
                    p = 3;
                }
                else
                {
                    c = lc.Find(z => z.tarifa.ToUpper() == tarifa.ToUpper() &&
                                     z.territorio == territorio &&
                                     z.estacion == Estacion(f, tarifa.ToUpper(), territorio) &&
                                     z.laborable == laborable &&
                                     z.fd <= f && z.fh >= f);
                    if (c != null)
                    {
                        if (f == UltimoDomingoOctubre(f) && ph > 3)
                        {
                            p = c.periodos[ph - 1];
                        }
                        else if (f == UltimoDomingoMarzo(f) && ph > 2)
                        {
                            p = c.periodos[ph + 1];
                        }
                        else
                        {
                            p = c.periodos[ph];
                        }

                    } // if (c != null)
                    else
                    {
                        MessageBox.Show("Error a la hora de buscar la tarifa " + tarifa.ToUpper()
                            + System.Environment.NewLine
                            + territorio,
                         "Error a la hora de buscar la tarifa",
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Error);
                    }
                } // if(tarifa == "6.x" && (!EsLaborable(f) || EsFestivo(f)))
            } // if (buscarPeriodoTarifario)

            return p;

        }

        private bool EsLaborable(DateTime f)
        {
            return (f.DayOfWeek >= DayOfWeek.Monday && f.DayOfWeek <= DayOfWeek.Friday);
        }

        private bool EsFestivo(DateTime f)
        {
            DateTime fecha;
            fecha = lf.Find(z => z.Equals(f.Date));
            return fecha.Equals(f.Date);
        }

        private void CargaDefinicionCalendario(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";
            

            try
            {

                strSql = "select * from fact.fact_calendarios"
                    + " where FechaDesde <= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " and FechaHasta >= '" + fh.ToString("yyyy-MM-dd") + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.Calendario c = new EndesaEntity.Calendario();
                    c.tarifa = reader["Tarifa"].ToString();
                    c.territorio = reader["Territorio"].ToString();
                    c.estacion = reader["Estacion"].ToString();
                    c.laborable = (reader["Laborable"].ToString() == "S" ? true : false);
                    c.fd = Convert.ToDateTime(reader["FechaDesde"]);
                    c.fh = Convert.ToDateTime(reader["FechaHasta"]);
                    for (int i = 1; i <= 24; i++)
                    {
                        c.periodos[i] = Convert.ToInt32(reader["Value" + i]);
                    }

                    lc.Add(c);
                }

                reader.Close();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Error en la carga de calendarios",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

        private string TerritorioDesdeCUPS13(string cups13)
        {
            string territorio = "";
            if (cups13.Substring(0, 3) == "XCA")
            {
                territorio = "MELILLA";
            }
            else if (cups13.Substring(0, 3) == "XAA")
            {
                territorio = "CEUTA";
            }
            else
            {
                switch (cups13.Substring(0, 1))
                {
                    case "G":
                        territorio = "BALEARES";
                        break;
                    case "U":
                        territorio = "CANARIAS";
                        break;
                    default:
                        territorio = "PENINSULA";
                        break;
                }
            }

            return territorio;
        }

        private string Estacion(DateTime f, string tarifa, string zona)
        {
            string estacion = "";
            if (tarifa.Substring(0, 1) == "6" || tarifa.Substring(0, 1) == "3")
            {
                estacion = ConvierteMes(f, zona);
            }
            else
            {
                if (f > UltimoDomingoMarzo(f) && f < UltimoDomingoOctubre(f))
                {
                    estacion = "VERANO";
                }
                else
                {
                    estacion = "INVIERNO";
                }
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

        private string ConvierteMes(DateTime f, string z)
        {
            string mes = "";
            if (f.Month == 6 && z == "PENINSULA")
            {
                if (f >= new DateTime(2021, 06, 01))
                    mes = "JUNIO";
                else
                    mes = (f.Day < 16 ? "JUNIO-1" : "JUNIO-2");
            }
            else
            {
                switch (f.Month)
                {
                    case 1:
                        mes = "ENERO";
                        break;
                    case 2:
                        mes = "FEBRERO";
                        break;
                    case 3:
                        mes = "MARZO";
                        break;
                    case 4:
                        mes = "ABRIL";
                        break;
                    case 5:
                        mes = "MAYO";
                        break;
                    case 6:
                        mes = "JUNIO";
                        break;
                    case 7:
                        mes = "JULIO";
                        break;
                    case 8:
                        mes = "AGOSTO";
                        break;
                    case 9:
                        mes = "SEPTIEMBRE";
                        break;
                    case 10:
                        mes = "OCTUBRE";
                        break;
                    case 11:
                        mes = "NOVIEMBRE";
                        break;
                    case 12:
                        mes = "DICIEMBRE";
                        break;
                }
            }

            return mes;
        }

        private void CargaListaFestivos(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";

            try
            {
                strSql = "select FechaFestivo from fact.fact_diasfestivos where"
                    + " FechaFestivo >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FechaFestivo <= '" + fh.ToString("yyyy-MM-dd") + "';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    lf.Add(Convert.ToDateTime(reader["FechaFestivo"]));
                }

                reader.Close();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la carga de festivos electricos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void CargaListaTarifas(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";
            

            try
            {
                strSql = "select tarifa, territorio from fact.fact_calendarios where" 
                    + " Tarifa <> '8.X' and Tarifa not LIKE '%TDVE'"
                    + " and (FechaDesde <= '" + fd.ToString("yyyy-MM-dd") + "'" 
                    + " and FechaHasta >= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + "group by tarifa, territorio";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.facturacion.TarifaTerritorio tt = new EndesaEntity.facturacion.TarifaTerritorio();

                    tt.tarifa = reader["tarifa"].ToString().ToUpper();
                    tt.territorio = reader["territorio"].ToString().ToUpper();
                    lTT.Add(tt);
                }

                reader.Close();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la carga de lista tarifas",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }

        private void CargaListaTarifas(DateTime fd, DateTime fh, string tarifa)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";


            try
            {
                strSql = "select tarifa, territorio from fact.fact_calendarios where Tarifa <> '8.X'"
                    + " and (FechaDesde <= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " and FechaHasta >= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + "group by tarifa, territorio";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.facturacion.TarifaTerritorio tt = new EndesaEntity.facturacion.TarifaTerritorio();

                    tt.tarifa = reader["tarifa"].ToString().ToUpper();
                    tt.territorio = reader["territorio"].ToString().ToUpper();
                    lTT.Add(tt);
                }

                reader.Close();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la carga de lista tarifas",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }

        public int NumPeriodosEntreFechas(DateTime fd, DateTime fh)
        {
            int i = 0;

            for (DateTime d = fd; fd <= fh; d.AddDays(1))
            {
                i = i + NumPeriodosHorarios(d);
            }

            return i;
        }
    }
}
