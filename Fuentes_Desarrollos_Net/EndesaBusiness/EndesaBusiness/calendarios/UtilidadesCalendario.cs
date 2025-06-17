using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.calendarios
{
    public class UtilidadesCalendario
    {

        Dictionary<DateTime, DateTime> dic;


        public UtilidadesCalendario()
        {
            dic = Carga();
        }


        private Dictionary<DateTime, DateTime> Carga()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Dictionary<DateTime, DateTime> d = new Dictionary<DateTime, DateTime>();
            DateTime fechaFestivo = new DateTime();

            try
            {

                strSql = "SELECT FechaFestivo from fact_diasfestivos";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["FechaFestivo"] != System.DBNull.Value)
                        fechaFestivo = Convert.ToDateTime(r["FechaFestivo"]);
                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                return null;
            }
            
        }

        public int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }

        public int NumPeriodosCuartoHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 92 : 96 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 100 : 96 : 96);
        }
        private DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        public string MesLetra(DateTime d)
        {
            string mes = "";
            switch(d.Month)
            {
                case 1:
                    mes = "Enero";
                    break;
                case 2:
                    mes = "Febrero";
                    break;
                case 3:
                    mes = "Marzo";
                    break;
                case 4:
                    mes = "Abril";
                    break;
                case 5:
                    mes = "Mayo";
                    break;
                case 6:
                    mes = "Junio";
                    break;
                case 7:
                    mes = "Julio";
                    break;
                case 8:
                    mes = "Agosto";
                    break;
                case 9:
                    mes = "Septiembre";
                    break;
                case 10:
                    mes = "Octubre";
                    break;
                case 11:
                    mes = "Noviembre";
                    break;
                case 12:
                    mes = "Diciembre";
                    break;

            }

            return mes;
        }

        public int CorreccionCambioHorario(DateTime fd, DateTime fh)
        {

            DateTime ultimoDomingoMarzo = new DateTime();
            DateTime ultimoDomingoOctubre = new DateTime();
            int numDias = 0;

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
                        numDias += -1;

                    if (d == ultimoDomingoOctubre)
                        numDias += 1;
                }

                return numDias;
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

        public DateTime UltimoDiaHabilDelMes(DateTime f)
        {            
            DateTime ultimoDiaMes = 
                new DateTime(f.Year, f.Month, DateTime.DaysInMonth(f.Year, f.Month));

            for (int i = 9; i > 0; i--)
            {
                if ((int)ultimoDiaMes.DayOfWeek != 0 
                    && (int)ultimoDiaMes.DayOfWeek != 6 
                    && !EsFestivo(ultimoDiaMes))
                {
                    
                    return ultimoDiaMes;
                }
                ultimoDiaMes = ultimoDiaMes.AddDays(-1);
            }

            return ultimoDiaMes;
        }


        private bool EsFestivo(DateTime f)
        {
            DateTime o;
            if (dic.TryGetValue(f, out o))
                return true;
            else
                return false;
        }

    }
}
