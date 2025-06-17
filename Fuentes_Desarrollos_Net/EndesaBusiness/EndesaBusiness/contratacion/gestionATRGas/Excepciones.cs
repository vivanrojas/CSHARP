using EndesaBusiness.servidores;
using EndesaEntity;
using EndesaEntity.cnmc.V21_2019_12_17;
using Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class Excepciones
    {
        logs.Log ficheroLog;

        //Listado de las excepciones (tabla atrgas_excepcion_tramitacion_xml + estado)
        public List<Table_atrgas_excepcion_tramitacion_xml> lista_excepcion_tramitacion { get; set; }

        public Excepciones() 
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ExcepcionTramitacion");
            lista_excepcion_tramitacion = new List<Table_atrgas_excepcion_tramitacion_xml>();
            

            GetAll();
        }
        public void ActualizaEstadoExcepcion(Int32 id, string estado)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                //strSql = "update atrgas_excepcion_tramitacion_xml set estado = '" + estado + "',last_update_by ='" + System.Environment.UserName + "'  where id='" + id +"';";
                //01/07/2024 GUS: modificamos eliminado actualizacion last_update_by cuando se trata del estado, únicamente cuando NO es "Cancelada", ya que se ejecuta en la carga del formulario
                
                if(estado == "Cancelada")
                    strSql = "update atrgas_excepcion_tramitacion_xml set estado = '" + estado + "',last_update_by ='" + System.Environment.UserName + "'  where id='" + id + "';";
                else
                    strSql = "update atrgas_excepcion_tramitacion_xml set estado = '" + estado + "' where id='" + id + "';";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Error actualizando estado de la excepción de tramitación ATR Gas: " + e.Message);
            }
        }

        // Obtenemos el contenido completo de la tabla atrgas_excepcion_tramitacion_xml
        // Estableciendo el estado según la fechora actual: Programada, En ejecución, Finalizada o Cancelada
        private void GetAll()
        {
            DateTime fechora_actual;
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                fechora_actual = DateTime.Now;

                strSql = "select id, nombre, tramitacion, fecha_desde, fecha_hasta,"
                + " created_by, created_date, last_update_by, last_update_date, estado"
                + " from atrgas_excepcion_tramitacion_xml";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Table_atrgas_excepcion_tramitacion_xml c = new Table_atrgas_excepcion_tramitacion_xml();

                    c.id = Convert.ToInt32(r["id"]);
                    c.nombre = r["nombre"].ToString();
                    c.tramitacion = r["tramitacion"].ToString();
                    c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);

                    c.created_by = r["created_by"].ToString();
                    c.creation_date = Convert.ToDateTime(r["created_date"]);
                    c.last_update_by = r["last_update_by"].ToString();
                    c.last_update_date = Convert.ToDateTime(r["last_update_date"]);

                    c.estado = (r["estado"].ToString());
                    //Actualizamos el estado de la excepción si no está cancelada comparando la fecha_desde y fecha_hasta con la fechora actual

                    if (c.estado != "Cancelada")
                    {
                        if (fechora_actual >= c.fecha_desde && fechora_actual <= c.fecha_hasta)
                        {
                            c.estado = "En ejecución";
                        }
                        else if (fechora_actual < c.fecha_desde)
                        {
                            c.estado = "Programada";  
                        }
                        else if (fechora_actual > c.fecha_hasta)
                        {
                            c.estado = "Finalizada";
                        }
                        else
                        {
                            c.estado = "Cancelada";
                        }
                        //Actualizar estado registro
                        ActualizaEstadoExcepcion(c.id, c.estado);
                    }
                    
                    lista_excepcion_tramitacion.Add(c);
                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Error obteniendo las excepciones de tramitación ATR Gas: " + e.Message);
            }

        }

        public static void Add(string nombre_grupo_distribuidora, DateTime fd, DateTime fh) 
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "insert into atrgas_excepcion_tramitacion_xml (nombre, fecha_desde,"
                    + "fecha_hasta, created_by, created_date, last_update_by) values ";

                strSql += "('" + nombre_grupo_distribuidora + "',"
                    + "'" + fd.ToString("yyyy-MM-dd HH:mm") + "',"
                    + "'" + fh.ToString("yyyy-MM-dd HH:mm") + "',"
                    + "'" + System.Environment.UserName + "',"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "',"
                    + "'" + System.Environment.UserName + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error al añadir nueva excepción de tramitación",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
        public static void Update(Int32 id, string nombre_grupo_distribuidora, DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update atrgas_excepcion_tramitacion_xml set nombre='" + nombre_grupo_distribuidora + "', fecha_desde='" +
                    fd.ToString("yyyy-MM-dd HH:mm") + "', fecha_hasta='" +
                    fh.ToString("yyyy-MM-dd HH:mm") + "', last_update_by='" + System.Environment.UserName + "'" +
                    " where id='" + id + "';";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error al modificar excepción de tramitación",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
        public List<string> GetNombresDistribuidoras()
        {
            List<string> lista_grupo_distribuidoras = new List<string>();
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "select distinct nombre"
                + " from atrgas_distribuidoras where tramitacion <> 'Mail';";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    lista_grupo_distribuidoras.Add(r["nombre"].ToString());  
                }
                db.CloseConnection();
            }    
            catch (Exception e)
            {
                ficheroLog.AddError("Error obteniendo los grupos de distribuidoras ATR Gas: " + e.Message);
            }

            return lista_grupo_distribuidoras;
        }
    }
}
