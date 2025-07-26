using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class RellenaTablaPlanificacionDiasHabiles
    {

        EndesaBusiness.utilidades.Fechas f;
        public RellenaTablaPlanificacionDiasHabiles()
        {
            f = new utilidades.Fechas();
        }

        public void InsertaDias(int anio)
        {
            MySqlCommand command;
            MySQLDB db;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            DateTime fecha_hasta = new DateTime();
            int num_registros = 0;
            int total_registros = 0;
            try
            {

                fecha_hasta = new DateTime(anio, 12, 31);

                for(DateTime d = UltimoDiaInsertado().AddDays(1); d <= fecha_hasta; d = d.AddDays(1))
                {
                    if (firstOnly)
                    {
                        sb.Append("replace into global_planificaciondiashabiles (fecha_ejecucion) ");
                        sb.Append(" values ");
                        firstOnly = false;
                    }
                    if (f.EsLaborable(d))
                    {
                        num_registros++;
                        total_registros++;
                        sb.Append("('").Append(d.ToString("yyyy-MM-dd")).Append("'),");

                        if(num_registros == 10)
                        {
                            Console.CursorLeft = 0;
                            Console.Write("Insertando " + total_registros + " registros ...");
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.AUX);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            num_registros = 0;
                        }
                    }
                    
                }

                if (num_registros > 0)
                {
                    Console.WriteLine("Insertando " + total_registros + " registros ...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;                    
                    num_registros = 0;
                }

            }
            catch(Exception e)
            {

            }
        }



        private DateTime UltimoDiaInsertado()
        {
            MySqlCommand command;
            MySqlDataReader r;
            MySQLDB db;
            string strSql = "";
            DateTime d = new DateTime();

            try
            {
                strSql = "select max(fecha_ejecucion) fecha from global_planificaciondiashabiles";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);                
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    d = Convert.ToDateTime(r["fecha"]);
                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                return DateTime.MinValue;
            }
        }
    }
}
