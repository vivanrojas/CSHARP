using EndesaBusiness.servidores;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Office.Core;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class Seguimiento_Procesos
    {
        logs.Log ficheroLog;
        //Dictionary<string, EndesaEntity.global.Global_fechas_procesos> dic;
        public Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos> dic;
        //public List<EndesaEntity.herramientas.Seguimiento_Procesos> lista { get; set; }
        //public List<EndesaEntity.herramientas.Seguimiento_Procesos> procesos { get; set; }
        //public List<EndesaEntity.herramientas.Seguimiento_Procesos> procesos_detalle { get; set; }
        public List<string> lista_areas { get; set; }
        public List<string> lista_procesos { get; set; }

        utilidades.Fechas utilFechas;

        public Seguimiento_Procesos()
        {
            utilFechas = new Fechas();
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "GBL_SS_PP");
            dic = Carga();
            //lista = CargaLista();
            lista_areas = CargaListaAreas();
            lista_procesos = CargaListaProcesos();
            //procesos = CargaProcesos();
            //procesos_detalle = CargaProcesosDetalle();
        }

        private List<string> CargaListaAreas()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;


            List<string> lista = new List<string>();

            try
            {

                strSql = "select area, proceso, paso, descripcion, fecha_inicio, fecha_fin"
                    + " from ss_pp_ejecucion where habilitado = 'S' group by area";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    lista.Add(r["area"].ToString());

                }
                db.CloseConnection();
                return lista;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("CargaListaAreas: " + ex.Message);
                return null;
            }


        }
        private List<string> CargaListaProcesos()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;


            List<string> lista = new List<string>();

            try
            {

                strSql = "select area, proceso, paso, descripcion, fecha_inicio, fecha_fin"
                    + " from ss_pp_ejecucion where habilitado = 'S' group by proceso";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {


                    lista.Add(r["proceso"].ToString());

                }
                db.CloseConnection();
                return lista;
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("CargaListaProcesos: " + ex.Message);
                return null;
            }


        }



        //private List<EndesaEntity.herramientas.Seguimiento_Procesos> CargaProcesos()
        //{
        //    servidores.MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;

        //    Dictionary<string, EndesaEntity.global.Global_fechas_procesos> d =
        //        new Dictionary<string, EndesaEntity.global.Global_fechas_procesos>();


        //    try
        //    {

        //        strSql = "select area, proceso, descripcion, fecha_inicio, fecha_fin"
        //            + " from ss_pp_procesos where habilitado = 'S'";
        //        db = new MySQLDB(MySQLDB.Esquemas.AUX);
        //        command = new MySqlCommand(strSql, db.con);
        //        r = command.ExecuteReader();
        //        while (r.Read())
        //        {
        //            EndesaEntity.global.Global_fechas_procesos c = new EndesaEntity.global.Global_fechas_procesos();
        //            c.area = r["area"].ToString();
        //            c.proceso = r["proceso"].ToString();                    

        //            if (r["descripcion"] != System.DBNull.Value)
        //                c.descripcion = r["descripcion"].ToString();

        //            if (r["fecha_inicio"] != System.DBNull.Value)
        //                c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

        //            if (r["fecha_fin"] != System.DBNull.Value)
        //                c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

        //            EndesaEntity.global.Global_fechas_procesos o;
        //            if (!d.TryGetValue(c.area + "_" + c.proceso + "_" + c.paso, out o))
        //                d.Add(c.area + "_" + c.proceso + "_" + c.paso, c);

        //        }
        //        db.CloseConnection();
        //        return d;
        //    }
        //    catch (Exception ex)
        //    {
        //        ficheroLog.AddError("Carga: " + ex.Message);
        //        return null;
        //    }


        //}

        //private List<EndesaEntity.herramientas.Seguimiento_Procesos> CargaProcesosDetalle()
        //{
        //    servidores.MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;

        //    Dictionary<string, EndesaEntity.global.Global_fechas_procesos> d =
        //        new Dictionary<string, EndesaEntity.global.Global_fechas_procesos>();


        //    try
        //    {

        //        strSql = "select area, proceso, descripcion, fecha_inicio, fecha_fin"
        //            + " from ss_pp_procesos where habilitado = 'S'";
        //        db = new MySQLDB(MySQLDB.Esquemas.AUX);
        //        command = new MySqlCommand(strSql, db.con);
        //        r = command.ExecuteReader();
        //        while (r.Read())
        //        {
        //            EndesaEntity.global.Global_fechas_procesos c = new EndesaEntity.global.Global_fechas_procesos();
        //            c.area = r["area"].ToString();
        //            c.proceso = r["proceso"].ToString();
        //            c.paso = r["paso"].ToString();

        //            if (r["descripcion"] != System.DBNull.Value)
        //                c.descripcion = r["descripcion"].ToString();

        //            if (r["fecha_inicio"] != System.DBNull.Value)
        //                c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

        //            if (r["fecha_fin"] != System.DBNull.Value)
        //                c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

        //            EndesaEntity.global.Global_fechas_procesos o;
        //            if (!d.TryGetValue(c.area + "_" + c.proceso + "_" + c.paso, out o))
        //                d.Add(c.area + "_" + c.proceso + "_" + c.paso, c);

        //        }
        //        db.CloseConnection();
        //        return d;
        //    }
        //    catch (Exception ex)
        //    {
        //        ficheroLog.AddError("Carga: " + ex.Message);
        //        return null;
        //    }


        //}


        private Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos> Carga()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos> d =
                new Dictionary<string, EndesaEntity.herramientas.Seguimiento_Procesos>();
                
            
            try
            {
                
                strSql = "select area, proceso,  descripcion, fecha_inicio, fecha_fin,"
                    + " comentario, ejecucion, tarea, contacto, habilitado, hora_inicio"
                    + " from ss_pp_procesos where habilitado = 'S'";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.herramientas.Seguimiento_Procesos c = new EndesaEntity.herramientas.Seguimiento_Procesos();

                    c.area = r["area"].ToString();
                    c.proceso = r["proceso"].ToString();                    

                    if (r["descripcion"] != System.DBNull.Value)
                        c.descripcion = r["descripcion"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["hora_inicio"] != System.DBNull.Value)
                        c.hora_inicio = Convert.ToDateTime(r["hora_inicio"]);

                    if (r["comentario"] != System.DBNull.Value)
                        c.comentario = r["comentario"].ToString();

                    if (r["ejecucion"] != System.DBNull.Value)
                        c.ejecucion = r["ejecucion"].ToString();

                    if (r["tarea"] != System.DBNull.Value)
                        c.tarea = r["tarea"].ToString();

                    if (r["contacto"] != System.DBNull.Value)
                        c.contacto = r["contacto"].ToString();

                    EndesaEntity.herramientas.Seguimiento_Procesos o;
                    if (!d.TryGetValue(c.area + "_" + c.proceso, out o))
                    {
                        d.Add(c.area + "_" + c.proceso, c);
                    }
                      
                    
                }
                db.CloseConnection();

                strSql = "select area, proceso, paso, descripcion, fecha_inicio, fecha_fin,"
                    + " comentario, ejecucion, tarea, contacto, habilitado"
                    + " from ss_pp_pasos where habilitado = 'S'";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.herramientas.Seguimiento_ProcesosDetalle c = 
                        new EndesaEntity.herramientas.Seguimiento_ProcesosDetalle();

                    c.area = r["area"].ToString();
                    c.proceso = r["proceso"].ToString();
                    c.paso = r["paso"].ToString();

                    if (r["descripcion"] != System.DBNull.Value)
                        c.descripcion = r["descripcion"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["comentario"] != System.DBNull.Value)
                        c.comentario = r["comentario"].ToString();

                    EndesaEntity.herramientas.Seguimiento_Procesos o;
                    if (d.TryGetValue(c.area + "_" + c.proceso, out o))
                    {
                        o.detalle.Add(c);
                    }

                }
                db.CloseConnection();

                return d;





            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Carga: " + ex.Message);                
                return null;
            }

            
        }

        //private List<EndesaEntity.herramientas.Seguimiento_Procesos> CargaLista()
        //{
        //    servidores.MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;

        //    List<EndesaEntity.herramientas.Seguimiento_Procesos> lista =
        //        new List<EndesaEntity.herramientas.Seguimiento_Procesos>();

            
        //    try
        //    {

        //        strSql = "select area, proceso, paso, descripcion, fecha_inicio, fecha_fin,"
        //            + " ejecucion, tarea, contacto, comentario"
        //            + " from ss_pp_pasos where habilitado = 'S'";
        //        ficheroLog.Add(strSql);
        //        db = new MySQLDB(MySQLDB.Esquemas.AUX);
        //        command = new MySqlCommand(strSql, db.con);
        //        r = command.ExecuteReader();
        //        while (r.Read())
        //        {
        //            EndesaEntity.herramientas.Seguimiento_Procesos c = new EndesaEntity.herramientas.Seguimiento_Procesos();
        //            c.area = r["area"].ToString();
        //            c.proceso = r["proceso"].ToString();
        //            c.paso = r["paso"].ToString();

        //            if (r["descripcion"] != System.DBNull.Value)
        //                c.descripcion = r["descripcion"].ToString();

        //            if (r["fecha_inicio"] != System.DBNull.Value)
        //                c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

        //            if (r["fecha_fin"] != System.DBNull.Value)
        //                c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

        //            if (r["ejecucion"] != System.DBNull.Value)
        //                c.ejecucion = r["ejecucion"].ToString();

        //            if (r["tarea"] != System.DBNull.Value)
        //                c.tarea = r["tarea"].ToString();

        //            if (r["contacto"] != System.DBNull.Value)
        //                c.contacto = r["contacto"].ToString();

        //            if (r["comentario"] != System.DBNull.Value)
        //                c.comentario = r["comentario"].ToString();


        //            lista.Add(c);

        //        }
        //        db.CloseConnection();
        //        return lista;
        //    }
        //    catch (Exception ex)
        //    {
        //        ficheroLog.AddError("Carga: " + ex.Message);                
        //        return null;
        //    }


        //}

        

        public void Update_Fecha_Inicio(string area, string proceso, string paso)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            
            try
            {

                if(ExisteProceso(area, proceso, paso))
                {
                    strSql = "update ss_pp_pasos set"
                       + " fecha_inicio = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                       + " ,fecha_fin = null"
                       + " ,comentario = 'En ejecución'"
                       + " where area = '" + area + "' and"
                       + " proceso = '" + proceso + "' and"
                       + " paso = '" + paso + "'";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    strSql = "update ss_pp_procesos set"
                       + " fecha_inicio = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                       + " ,fecha_fin = null"
                       + " ,comentario = 'En ejecución'"
                       + " where area = '" + area + "' and"
                       + " proceso = '" + proceso + "'";
                       
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                }
                else
                {
                    strSql = "replace into ss_pp_pasos "
                        + " (area, proceso, paso, fecha_inicio, fecha_fin, comentario) values"
                        + " ('" + area + "',"
                        + " '" + proceso + "',"
                        + " '" + paso + "',"
                        + " '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                        + " null,"
                        + " 'En ejecución')";
                        
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    strSql = "replace into ss_pp_procesos "
                        + " (area, proceso, fecha_inicio, fecha_fin, comentario) values"
                        + " ('" + area + "',"
                        + " '" + proceso + "',"                        
                        + " '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                        + " null,"
                        + " 'En ejecución')";

                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();


                }

                

            }
            catch(Exception ex)
            {
                ficheroLog.AddError("Update_Fecha_Inicio: " + ex.Message);
                
            }

           
            
        }

        public void Update_Fecha_Fin(string area, string proceso, string paso)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            bool actualiza_proceso = true;
            string comentario = "";
            try
            {
                strSql = "update ss_pp_pasos set"
                + " fecha_fin = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                + " ,comentario = 'OK'"
                + " where area = '" + area + "' and"
                + " proceso = '" + proceso + "' and"
                + " paso = '" + paso + "'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Comprobamos que todos los pasos
                // sean ok para poner ok al proceso padre
                
                strSql = "Select comentario from ss_pp_pasos"
                    + " where area = '" + area + "' and"
                    + " proceso = '" + proceso + "' and"
                    + " comentario <> 'OK'";

                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    actualiza_proceso = false;
                    comentario = r["comentario"].ToString();
                }
                //actualiza_proceso = !r.Read();
                db.CloseConnection();

                if (actualiza_proceso)
                {
                    strSql = "update ss_pp_procesos set"
                        + " fecha_fin = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                        + " comentario = 'OK'"
                        + " where area = '" + area + "' and"
                        + " proceso = '" + proceso + "'";

                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }
                else
                {
                    strSql = "update ss_pp_procesos set"
                        + " fecha_fin = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                        + " comentario = '" + comentario + "'"
                        + " where area = '" + area + "' and"
                        + " proceso = '" + proceso + "'";

                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }

                

            }
            catch (Exception ex)
            {
                ficheroLog.AddError("Update_Fecha_Fin: " + ex.Message);
                
            }
            
            
        }

        public void Update_Comentario(string area, string proceso, string paso, string comentario)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            bool actualiza_proceso = false;

            try
            {
                strSql = "update ss_pp_pasos set"
                    + " comentario = '" + comentario + "'"
                    + " where area = '" + area + "' and"
                    + " proceso = '" + proceso + "' and"
                    + " paso = '" + paso + "'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "update ss_pp_procesos set"
                    + " comentario = '" + comentario + "'"
                    + " where area = '" + area + "' and"
                    + " proceso = '" + proceso + "'";
                    
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


               


            }
            catch (Exception ex)
            {
                ficheroLog.AddError("Update_Comentario: " + ex.Message);
            }


        }

        private bool ExisteProceso(string area, string proceso, string paso)
        {
            EndesaEntity.herramientas.Seguimiento_Procesos o;
            if (dic.TryGetValue(area + "_" + proceso, out o))            
                return o.detalle.Exists(c => c.area == area && c.proceso == proceso && c.paso == paso);
            
            else
                return false;
                
        }

        public DateTime GetFecha_InicioProceso(string area, string proceso, string paso)
        {
            EndesaEntity.herramientas.Seguimiento_Procesos o;
            DateTime f = DateTime.MinValue;
            if (dic.TryGetValue(area + "_" + proceso, out o))
            {
                EndesaEntity.herramientas.Seguimiento_ProcesosDetalle p =
                    o.detalle.Find(c => c.area == area && c.proceso == proceso && c.paso == paso);
                if (p != null)
                    f = p.fecha_inicio;
            }

            return f;


        }

        public DateTime GetFecha_FinProceso(string area, string proceso, string paso)
        {
            EndesaEntity.herramientas.Seguimiento_Procesos o;
            DateTime f = DateTime.MinValue;
            if (dic.TryGetValue(area + "_" + proceso, out o))
            {
                EndesaEntity.herramientas.Seguimiento_ProcesosDetalle p =
                    o.detalle.Find(c => c.area == area && c.proceso == proceso && c.paso == paso);
                if (p != null)
                    f = p.fecha_fin;
            }

            return f;
        }

        public string GetComentarioProceso(string area, string proceso, string paso)
        {
            EndesaEntity.herramientas.Seguimiento_Procesos o;
            string f = "";
            if (dic.TryGetValue(area + "_" + proceso, out o))
            {
                EndesaEntity.herramientas.Seguimiento_ProcesosDetalle p =
                    o.detalle.Find(c => c.area == area && c.proceso == proceso && c.paso == paso);
                if (p != null)
                    f = p.comentario;
            }

            return f;
        }

        public void Reset()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            string ejecucion = "";


            try
            {
                ficheroLog.Add("=====");
                ficheroLog.Add("RESET");
                ficheroLog.Add("=====");

                if (utilFechas.EsLaborable())
                {
                    strSql = "update ss_pp_pasos set comentario = ''"
                    + " where ejecucion = 'L-V'"
                    + " and habilitado = 'S'";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    strSql = "update ss_pp_procesos set comentario = ''"
                    + " where ejecucion = 'L-V'"
                    + " and habilitado = 'S'";
                    ficheroLog.Add(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                }

                switch ((int)DateTime.Now.DayOfWeek)
                {
                    case 1:
                        ejecucion = "L";
                        break;
                    case 2:
                        ejecucion = "M";
                        break;
                    case 3:
                        ejecucion = "X";
                        break;
                    case 4:
                        ejecucion = "J";
                        break;
                    case 5:
                        ejecucion = "V";
                        break;
                }
                strSql = "update ss_pp_pasos set comentario = ''"
                    + " where ejecucion = '" + ejecucion + "'"
                    + " and habilitado = 'S'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.AUX);                
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "update ss_pp_procesos set comentario = ''"
                + " where ejecucion = '" + ejecucion + "'"
                + " and habilitado = 'S'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }catch(Exception ex)
            {
                ficheroLog.addError(ex.Message);
            }

            

        }

    }
}
