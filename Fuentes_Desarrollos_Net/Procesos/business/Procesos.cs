using MySql.Data.MySqlClient;
using Procesos.servidores;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Procesos.business
{
    class Procesos
    {
        business.Parametricas param;
        business.Fechas fechas;
        public Dictionary<int, List<EndesaEntity.cola.ProcesoCola>> dic;
        MSAccess access;
        Log ficheroLog;
        
        public Procesos(string cola)
        {
            ficheroLog = new Log(Environment.CurrentDirectory, "logs", cola);
            access = new MSAccess();
            fechas = new Fechas();
            param = new business.Parametricas();
            dic = CargaCola(cola);
            
        }

        private Dictionary<int, List<EndesaEntity.cola.ProcesoCola>> CargaCola(string cola)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<int, List<EndesaEntity.cola.ProcesoCola>> d
                = new Dictionary<int, List<EndesaEntity.cola.ProcesoCola>>();


            try
            {
                strSql = "select cola, grupo, proceso, id_p_proceso, ruta, "
                    + " bbdd, nombre_proceso, activo, obligatorio, una_vez,"
                    + " descripcion, fecha_ultima_ejec_ok, fecha_ultimo_lanzamiento,"
                    + " id_p_periodicidad, id_p_parametro, parametro, mensaje_error"
                    + " from aux1.q_procesos "
                    + " where cola = '" + cola + "'";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cola.ProcesoCola c = new EndesaEntity.cola.ProcesoCola();

                    if (r["cola"] != System.DBNull.Value)
                        c.cola = r["cola"].ToString();

                    if (r["grupo"] != System.DBNull.Value)
                        c.grupo = Convert.ToInt32(r["grupo"]);

                    if (r["proceso"] != System.DBNull.Value)
                        c.proceso = Convert.ToInt32(r["proceso"]);

                    if (r["ruta"] != System.DBNull.Value)
                        c.ruta = r["ruta"].ToString();

                    if (r["id_p_proceso"] != System.DBNull.Value)
                        c.id_p_proceso = Convert.ToInt32(r["id_p_proceso"]);

                    if (r["bbdd"] != System.DBNull.Value)
                        c.bbdd = r["bbdd"].ToString();

                    if (r["nombre_proceso"] != System.DBNull.Value)
                        c.nombre_proceso = r["nombre_proceso"].ToString();

                    c.activo = r["activo"].ToString() == "S";
                    c.obligatorio = r["obligatorio"].ToString() == "S";
                    c.una_vez = r["una_vez"].ToString() == "S";

                    if (r["descripcion"] != System.DBNull.Value)
                        c.descripcion = r["descripcion"].ToString();

                    if (r["fecha_ultima_ejec_ok"] != System.DBNull.Value)
                        c.fecha_ultima_ejec_ok = Convert.ToDateTime(r["fecha_ultima_ejec_ok"]);

                    if (r["fecha_ultimo_lanzamiento"] != System.DBNull.Value)
                        c.fecha_ultimo_lanzamiento = Convert.ToDateTime(r["fecha_ultimo_lanzamiento"]);

                    if (r["id_p_periodicidad"] != System.DBNull.Value)
                        c.id_p_periodicidad = Convert.ToInt32(r["id_p_periodicidad"]);

                    if (r["id_p_parametro"] != System.DBNull.Value)
                        c.id_p_parametro = Convert.ToInt32(r["id_p_parametro"]);

                    if (r["mensaje_error"] != System.DBNull.Value)
                        c.mensaje_error = r["mensaje_error"].ToString();

                    if (r["parametro"] != System.DBNull.Value)
                        c.parametro = r["parametro"].ToString();

                        List<EndesaEntity.cola.ProcesoCola> o;
                    if (!d.TryGetValue(c.grupo, out o))
                    {
                        o = new List<EndesaEntity.cola.ProcesoCola>();
                        o.Add(c);
                        d.Add(c.grupo, o);
                    }
                    else
                        o.Add(c);


                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public EndesaEntity.cola.ProcesoCola GetProceso(int grupo, int proceso)
        {
            List<EndesaEntity.cola.ProcesoCola> o;
            if (dic.TryGetValue(grupo, out o))
            {
                for (int i = 0; i < o.Count; i++)
                    if (o[i].proceso == proceso)
                        return o[i];

                return null;
            }
            else
                return null;

        }

        public void EjecutaCola()
        {
            bool hayerror = false;
            bool firstOnly = true;

            foreach (KeyValuePair<int, List<EndesaEntity.cola.ProcesoCola>> p in dic)
            {
                hayerror = false;
                firstOnly = true;

                foreach (EndesaEntity.cola.ProcesoCola c in p.Value)
                {
                    // si hay un proceso obligatorio y falla
                    // el resto de procesos del nodo no se ejecutarán
                    if(c.activo && Lanzar(c.id_p_periodicidad))
                    {
                        if (c.obligatorio && firstOnly)
                        {
                            hayerror = EjecutaNodo(c, false);
                            firstOnly = false;

                            ficheroLog.Add("Ejecutando: "
                               + " Grupo: " + c.grupo
                               + " Proceso: " + c.proceso
                               + " bbdd: " + c.bbdd
                               + " nombre_proceso: " + c.nombre_proceso
                               + " obligatorio: " + c.obligatorio
                               + " Hay Error: " + hayerror);

                        }
                        else
                        {
                            if (!hayerror)
                            {
                                EjecutaNodo(c, false);
                                firstOnly = true;

                                ficheroLog.Add("Ejecutando: "
                               + " Grupo: " + c.grupo
                               + " Proceso: " + c.proceso
                               + " bbdd: " + c.bbdd
                               + " nombre_proceso: " + c.nombre_proceso
                               + " obligatorio: " + c.obligatorio
                               + " Hay Error: " + hayerror);
                            }

                        }
                    }
                                        
                }
                    
            }
                
        }


        public bool EjecutaNodo(EndesaEntity.cola.ProcesoCola c, bool ejecutaAhora)
        {

            bool hay_error = true;
            try
            {
                if ((c.activo && Lanzar(c.id_p_periodicidad)) || ejecutaAhora)
                {
                    switch (param.GetProceso(c.id_p_proceso))
                    {
                        case "Proceso Batch":
                            c = access.ProcesoBatch(c);                            
                            UpdateProceso(c);
                            hay_error = c.mensaje_error != "";
                            break;
                        case "Macro":
                            c = access.EjecutaMacro(c);
                            hay_error = c.mensaje_error != "";
                            UpdateProceso(c);                            
                            break;
                        case "Procedimiento":                            
                            hay_error = false;
                            break;
                        case "Envío Correo":                            
                            hay_error = false;
                            break;
                        case "Ejecuta Parámetro":                            
                            hay_error = false;
                            break;
                        case "Guarda Correo":                            
                            hay_error = false;
                            break;
                    }
                    return hay_error;
                }
                else
                    return false;
                    
            }
            catch(Exception e)
            {
                ficheroLog.AddError("Ejecutando: "
                               + " Grupo: " + c.grupo
                               + " Proceso: " + c.proceso
                               + " bbdd: " + c.bbdd
                               + " nombre_proceso: " + c.nombre_proceso
                               + " obligatorio: " + c.obligatorio
                               + " Hay Error: " + e.Message);

                c.mensaje_error = e.Message;
                UpdateProceso(c);
                return true;
            }
        }

        private bool Lanzar(int id_periodicidad)
        {
            bool ejecuta = false;
            switch (param.GetPeriodicidad(id_periodicidad))
            {
                case "DIARIO":
                    ejecuta = true;
                    break;
                case "QUINCENAL":
                    ejecuta = DateTime.Now.Date == fechas.MediadosDeMes();
                    break;
                case "PRIMER DÍA HÁBIL DEL MES":
                    ejecuta = DateTime.Now.Date == fechas.PrimerDiaHabilDelMes();
                    break;
                case "ÚLTIMO DÍA HÁBIL DEL MES":
                    ejecuta = DateTime.Now.Date == fechas.UltimoDiaHabilDelMes();
                    break;
                default:
                    ejecuta = param.GetPeriodicidad(id_periodicidad) == DiaDeLaSemana();
                    break;

            }
            return ejecuta;
        }

        private string DiaDeLaSemana()
        {

            int dd = (int)DateTime.Now.DayOfWeek;
            string d = "";
            switch (dd)
            {
                case 1:
                    d = "LUNES";
                    break;
                case 2:
                    d = "MARTES";
                    break;
                case 3:
                    d = "MIÉRCOLES";
                    break;
                case 4:
                    d = "JUEVES";
                    break;
                case 5:
                    d = "VIERNES";
                    break;
                case 6:
                    d = "SÁBADO";
                    break;
                case 0:
                    d = "DOMINDO";
                    break;
            }
            return d;
        }

        public Color ColorNodo(EndesaEntity.cola.ProcesoCola c)
        {
            
            DateTime hoy = new DateTime();
            hoy = DateTime.Now;
            hoy = hoy.Date;

            if (c.fecha_ultimo_lanzamiento > hoy && c.fecha_ultima_ejec_ok > c.fecha_ultimo_lanzamiento)
                return Color.Blue;

            if (c.fecha_ultimo_lanzamiento > hoy && c.fecha_ultima_ejec_ok < hoy)
                return Color.Red;

            if (c.activo && Lanzar(c.id_p_periodicidad))
                return Color.Green;

            

            return Color.Black;
        }

        public void GuardaDatos(EndesaEntity.cola.ProcesoCola c)
        {
            if (ExisteProceso(c))
                UpdateProceso(c);
            else
                NewProceso(c);
        }

        private void NewProceso(EndesaEntity.cola.ProcesoCola c)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            sb.Append("insert into q_procesos");
            sb.Append(" (cola, grupo, proceso, id_p_proceso, ruta, bbdd, nombre_proceso, activo, obligatorio, descripcion,");
            sb.Append(" id_p_periodicidad, id_p_parametro, parametro) values ");

            sb.Append("('").Append(c.cola).Append("',");
            sb.Append(c.grupo).Append(",");
            sb.Append(c.proceso).Append(",");
            sb.Append(c.id_p_proceso).Append(",");
            sb.Append("'").Append(c.ruta.Replace(@"\", "\\\\")).Append("',");
            sb.Append("'").Append(c.bbdd).Append("',");
            sb.Append("'").Append(c.nombre_proceso).Append("',");
            sb.Append("'").Append(c.activo ? "S" : "N").Append("',");
            sb.Append("'").Append(c.obligatorio ? "S" : "N").Append("',");

            if (c.descripcion != "")
                sb.Append("'").Append(c.descripcion).Append("',");
            else
                sb.Append("null,");

            sb.Append(c.id_p_periodicidad).Append(",");
            sb.Append(c.id_p_parametro).Append(",");

            if (c.descripcion != "")
                sb.Append("'").Append(c.parametro).Append("')");
            else
                sb.Append("null)");


            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(sb.ToString(), db.con);
            command.ExecuteNonQuery();
        }

        private void UpdateProceso(EndesaEntity.cola.ProcesoCola c)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "update q_procesos set"
                + " id_p_proceso = " + c.id_p_proceso + ","
                + " ruta = '" + c.ruta.Replace(@"\", "\\\\") + "',"
                + " bbdd = '" + c.bbdd + "',"
                + " nombre_proceso = '" + c.nombre_proceso + "',"
                + " activo = '" + (c.activo ? "S" : "N") + "',"
                + " obligatorio = '" + (c.obligatorio ? "S" : "N") + "',"
                + " una_vez = '" + (c.una_vez ? "S" : "N") + "',"
                + " descripcion = '" + c.descripcion + "',"
                + " fecha_ultima_ejec_ok = '" + c.fecha_ultima_ejec_ok.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                + " fecha_ultimo_lanzamiento = '" + c.fecha_ultimo_lanzamiento.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                + " id_p_periodicidad = " + c.id_p_periodicidad + ","
                + " id_p_parametro = " + c.id_p_parametro;

                if(c.mensaje_error != null)
                    strSql += ", mensaje_error = '" + utilidades.FuncionesTexto.ArreglaAcentos(c.mensaje_error) + "'";

                strSql += " where cola = '" + c.cola + "' and"
                + " grupo = " + c.grupo + " and"
                + " proceso = " + c.proceso;
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();

        }

        public void EliminaNodo(EndesaEntity.cola.ProcesoCola c)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "delete from q_procesos where"
                + " cola = '" + c.cola + "' and"
                + " grupo = " + c.grupo + " and"
                + " proceso = " + c.proceso;

            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
        }

        public bool ExisteProceso(EndesaEntity.cola.ProcesoCola c)
        {
            List<EndesaEntity.cola.ProcesoCola> o;
            if (dic.TryGetValue(c.grupo, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].grupo == c.grupo && o[i].proceso == c.proceso)
                        return true;
                }
            }
            return false;
        }

     
        public void ImportarCola()
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "MS Access accdb|*.accdb;|MS Access viejuno|*.mdb";
            d.Multiselect = false;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    CopiaCola(fileName);
                }
            }
        }

        private void CopiaCola(string rutaAccess)
        {
            OleDbDataReader r;
            servidores.AccessDB acc;
            string strSql = "";

            List<EndesaEntity.cola.ProcesoCola> lista = new List<EndesaEntity.cola.ProcesoCola>();
            try
            {
                strSql = "select cd_Grupo, cd_Proceso, Tipo_Proceso, Ruta,"
                    + " BD, Nombre_Proceso, Activo, Obligatorio, UnaVez, Descripcion,"
                    + " Fecha_Ultima_Ejec_ok, Fecha_Ultimo_Lanza, Periodicidad,"
                    + " Tipo_Parametro, Parametro, MsgBreakPoint, PeriodoEjecucion"
                    + " from procesos";
                acc = new servidores.AccessDB(rutaAccess);
                OleDbCommand cmd = new OleDbCommand(strSql, acc.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cola.ProcesoCola c = new EndesaEntity.cola.ProcesoCola();
                    c.grupo = Convert.ToInt32(r["cd_Grupo"]);
                    c.proceso = Convert.ToInt32(r["cd_Proceso"]);
                    c.id_p_proceso = Convert.ToInt32(r["Tipo_Proceso"]);
                    c.ruta = r["Ruta"].ToString();
                    
                    if (r["BD"] != System.DBNull.Value)
                        c.bbdd = r["BD"].ToString();

                    if (r["Nombre_Proceso"] != System.DBNull.Value)
                        c.nombre_proceso = r["Nombre_Proceso"].ToString();




                }
                acc.CloseConnection();

                
            }
            catch(Exception e)
            {
                
            }
        }

    }
}
