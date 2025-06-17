using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.utilidades
{
    public class Cargas
    {

        Dictionary<int, DateTime> dic_ultima_factura;


        public Cargas()
        {
            dic_ultima_factura = UltimaFechaFactura();
        }


        private Dictionary<int, DateTime> UltimaFechaFactura()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();
            int empresa_id = 0;

            Dictionary<int, DateTime> d =
                new Dictionary<int, DateTime>();

            try
            {
                strSql = "select empresa_id, FFACTURA from fact.fo_ultima_factura";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["empresa_id"] != System.DBNull.Value)
                        empresa_id = Convert.ToInt32(r["empresa_id"]);
                    if (r["FFACTURA"] != System.DBNull.Value)
                        fecha = Convert.ToDateTime(r["FFACTURA"]);


                    d.Add(empresa_id, fecha);                    
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_Alarmas",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);

                return null;
            }
        }

        public DateTime GetFechaUltimaFactura(int empresa_id)
        {
            DateTime o;
            if (dic_ultima_factura.TryGetValue(empresa_id, out o))
                return o;
            else
                return DateTime.MinValue;
        }


        public DateTime Fecha_UltimaActualizacionMedida(string tabla, string campo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();

            try
            {
                strSql = "select max(" + campo + ") f_ult_mod from med." + tabla;

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["f_ult_mod"] != System.DBNull.Value)
                        fecha = Convert.ToDateTime(r["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_tablas_kee",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return fecha;

            }
        }

        public DateTime FechaActualizacion_tablas_kee(string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();
            try
            {
                strSql = "select max(f_ult_mod) f_ult_mod from med." + tabla;

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["f_ult_mod"] != System.DBNull.Value)
                        fecha = Convert.ToDateTime(r["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_tablas_kee",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }

        public DateTime FechaActualizacion_CurvasDatamart()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime fecha = new DateTime();
            try
            {
                strSql = "select max(f_ult_mod) f_ult_mod from med.dt_cc_m";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    fecha = Convert.ToDateTime(reader["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_Alarmas",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }

        public DateTime FechaActualizacion_contratosPS()
        {
            DateTime fecha = new DateTime();
            try
            {

                utilidades.Param p = new utilidades.Param("contratos_ps_param", MySQLDB.Esquemas.CON);
                fecha = p.LastUpdateParameter("ContratosPS_MD5");

                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_Alarmas",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }

        public DateTime FechaActualizacion_Alarmas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime fecha = new DateTime();
            try
            {
                strSql = "select max(F_ULT_MOD) f_ult_mod from fact.fact_alarmas";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    fecha = Convert.ToDateTime(reader["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_Alarmas",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }

        public  DateTime FechaActualizacion_cont_ps_at_mt()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime fecha = new DateTime();
            try
            {
                strSql = "select max(f_ult_mod) f_ult_mod from cont.contratos_ps_mt";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    fecha = Convert.ToDateTime(reader["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_cont_ps_at_mt",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }


        public DateTime FechaActualizacion_Contratos_Agora()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime fecha = new DateTime();
            try
            {
                strSql = "select max(f_ult_mod) f_ult_mod from fact.fo_agora";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    fecha = Convert.ToDateTime(reader["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechaActualizacion_Alarmas",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }

        public  DateTime UltimaActualizacionExtraccionAutoconsumo()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime fecha = new DateTime();

            try
            {
                strSql = "SELECT MAX(f_ult_mod) f_ult_mod FROM cont.ps_autoconsumos";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    fecha = Convert.ToDateTime(reader["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "UltimaActualizacionExtraccionAutoconsumo",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1899, 01, 01);

            }
        }

        public  DateTime UltimaActualizacionCuadroMando(string nombreProceso)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();

            try
            {
                strSql = "select f_ult_mod from cm_fechas_procesos where"
                    + " proceso = '" + nombreProceso + "'";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["f_ult_mod"] != System.DBNull.Value)
                        fecha = Convert.ToDateTime(r["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "UltimaActualizacionCuadroMando: " + nombreProceso,
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1999, 12, 31);

            }
        }

        public int NumRegTabla(string tabla, servidores.MySQLDB.Esquemas esquema)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int NumReg = 0;

            strSql = "select count(*) totalRegistros from " + tabla;

            db = new MySQLDB(esquema);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["totalRegistros"] != System.DBNull.Value)
                    NumReg = Convert.ToInt32(r["totalRegistros"]);
            }
            db.CloseConnection();
            return NumReg;
        }
    }
}
