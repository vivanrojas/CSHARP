using MySql.Data.MySqlClient;
using Procesos.servidores;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Procesos.business
{
    class Fechas
    {

        List<DateTime> festivos = new List<DateTime>();
        
        public Fechas()
        {
            CargaFestivos();
        }


        private void CargaFestivos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "Select FechaFestivo from fact_diasfestivos;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    festivos.Add(Convert.ToDateTime(reader["FechaFestivo"]));
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Fechas - CargaFestivos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public bool EsLaborable(DateTime d)
        {
            return (!EsFestivo(d) && (int)d.DayOfWeek != 0 && (int)d.DayOfWeek != 6);
        }

        private Boolean EsFestivo(DateTime f)
        {
            return festivos.Exists(element => element == f);
        }

        public DateTime PrimerDiaHabilDelMes()
        {
            DateTime hoy = new DateTime();
            hoy = DateTime.Now;
            DateTime primerDiaMes = new DateTime(hoy.Year, hoy.Month, 1);
            

            for (int i = 1; i > 8; i++)
            {
                if ((int)primerDiaMes.DayOfWeek != 0 && (int)primerDiaMes.DayOfWeek != 6 && HayPaseBatch(primerDiaMes))
                {
                    return primerDiaMes.Date;
                }
                primerDiaMes = primerDiaMes.AddDays(1);
            }
            return primerDiaMes.Date;
        }

        public DateTime UltimoDiaHabilDelMes()
        {
            DateTime hoy = new DateTime();
            hoy = DateTime.Now;
            DateTime ultimodiaDelMes = new DateTime(hoy.Year, hoy.Month, DateTime.DaysInMonth(hoy.Year, hoy.Month));
            if (hoy.Date == ultimodiaDelMes.Date)
            {
                hoy = ultimodiaDelMes.AddDays(1);

            }

            for (int i = 9; i > 0; i--)
            {
                if ((int)ultimodiaDelMes.DayOfWeek != 0 && (int)ultimodiaDelMes.DayOfWeek != 6 && HayPaseBatch(ultimodiaDelMes))
                {
                    return ultimodiaDelMes.Date;
                }
                ultimodiaDelMes = ultimodiaDelMes.AddDays(-1);
            }
            return ultimodiaDelMes.Date;
        }

        public DateTime UltimoDiaHabil()
        {
            DateTime hoy = new DateTime();
            try
            {
                hoy = DateTime.Now;
                hoy = hoy.AddDays(-1);

                for (int i = 9; i > 0; i--)
                {
                    if ((int)hoy.DayOfWeek != 0 && (int)hoy.DayOfWeek != 6 && HayPaseBatch(hoy))
                    {
                        return hoy;
                    }
                    hoy = hoy.AddDays(-1);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error en último día habil " + e.Message);
            }
            return hoy;
        }

        public DateTime UltimoDiaHabilAnterior(DateTime f)
        {
            DateTime hoy = f;
            try
            {

                hoy = hoy.AddDays(-1);

                for (int i = 9; i > 0; i--)
                {
                    if ((int)hoy.DayOfWeek != 0 && (int)hoy.DayOfWeek != 6 && HayPaseBatch(hoy))
                    {
                        return hoy;
                    }
                    hoy = hoy.AddDays(-1);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error en último día habil " + e.Message);
            }
            return hoy;
        }

        public bool HayPaseBatch(DateTime fecha)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            bool haypase = true;

            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select FechaFestivo from fact.fact_diasfestivos where " +
                "FechaFestivo = '" + fecha.ToString("yyyy-MM-dd") + "';";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    haypase = false;
                }
                db.CloseConnection();
                return haypase;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


            return haypase;
        }

        public void SalvarFechaProceso(string nombre_proceso, DateTime fecha)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                strSql = "replace into go_fechas_procesos (codigo_proceso, fecha_fin_proceso) values ("
                    + "'" + nombre_proceso + "',"
                    + "'" + fecha.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public int Estacion(DateTime f)
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


        public int Estacion(DateTime f, bool hora_repetida)
        {
            int estacion = 0; // Invierno

            if ((f > UltimoDomingoMarzo(f).AddHours(2) && f < UltimoDomingoOctubre(f).AddHours(3)) && !hora_repetida)
                estacion = 1; // Verano            
            else if (f.Date == UltimoDomingoOctubre(f) && f.Hour == 2 && hora_repetida)
                estacion = 0;

            return estacion;
        }


        public string DD_MM_YYYY_YYYY_MM_DD(string f)
        {
            string fecha = null;

            if (f != "00/00/0000")
            {
                fecha = f.Substring(6, 4) + "-" + f.Substring(3, 2) + "-" + f.Substring(0, 2);
            }

            return fecha;
        }

        public string ConvierteMes_a_Letra(DateTime f)
        {
            string mes = "";

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


            return mes;
        }

        public int NumHoras(DateTime fd, DateTime fh)
        {
            return (Convert.ToInt32((fh - fd).TotalDays + 1) * 24) + CorreccionCambioHorario(fd, fh);
        }

        private int CorreccionCambioHorario(DateTime fd, DateTime fh)
        {

            DateTime ultimoDomingoMarzo = new DateTime();
            DateTime ultimoDomingoOctubre = new DateTime();

            try
            {
                if (fh.Month < fd.Month)
                {
                    ultimoDomingoMarzo = UltimoDomingoMarzo(fh);
                    ultimoDomingoOctubre = UltimoDomingoOctubre(fd);
                }
                else
                {
                    ultimoDomingoMarzo = UltimoDomingoMarzo(fd);
                    ultimoDomingoOctubre = UltimoDomingoOctubre(fh);
                }

                for (DateTime d = fd; d <= fh; d = d.AddDays(1))
                {
                    if (d == ultimoDomingoMarzo)
                        return -1;

                    if (d == ultimoDomingoOctubre)
                        return 1;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "ClaseCalendario - CorreccionCambioHorario",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

            return 0;
        }

        public DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        public DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        public int NumPeriodosCuartoHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 92 : 96 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 100 : 96 : 96);
        }

        public int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }

        public bool EsAnioBisiesto(DateTime f)
        {
            return DateTime.IsLeapYear(f.Year);
        }

        public DateTime UltimoDiaHabilMes()
        {
            DateTime hoy = new DateTime();
            try
            {
                hoy = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                hoy = hoy.AddMonths(1);
                hoy = hoy.AddDays(-1);

                for (int i = 9; i > 0; i--)
                {
                    if ((int)hoy.DayOfWeek != 0 && (int)hoy.DayOfWeek != 6 && HayPaseBatch(hoy))
                    {
                        return hoy;
                    }
                    hoy = hoy.AddDays(-1);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error en último día habil " + e.Message);
            }
            return hoy;
        }

        public DateTime MediadosDeMes()
        {
            DateTime d = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 15);
            // Domingo
            for (int i = 1; i > 9; i++)
            {
                if ((int)d.DayOfWeek != 0 && (int)d.DayOfWeek != 6 && HayPaseBatch(d))
                    return d.Date;
                else
                    d = d.AddDays(1);
            }
            return d.Date;



        }
    }
}
