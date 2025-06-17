using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EndesaBusiness.eer
{
    public class CurvasAccess
    {

        utilidades.Param param;
        logs.Log ficheroLog;
        public CurvasAccess()
        {
            param = new utilidades.Param("eer_param", MySQLDB.Esquemas.CON);
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_Curvas_EER");
        }


        public void Proceso()
        {

            //CopiaCurvas();
            CopiaDatosPeaje();
            TrataCurvas();

        }


        private void CopiaCurvas()
        {
            int i = 0;
            int x = 0;
            string strSql = "";
            servidores.AccessDB ac;
            OleDbDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;

            MySQLDB db;
            MySqlCommand command;
            int total_registros = 0;

            try
            {

                strSql = "delete from cont.eer_curvas_cuarto_horarias_tmp";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                strSql = "select count(*) as total_registros from [EER_CURVAS CUARTO HORARIAS]";
                ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas", DateTime.Now, DateTime.Now));
                OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                    total_registros = Convert.ToInt32(r["total_registros"]);

                ac.CloseConnection();

                Console.WriteLine("Total registros encontrados en: "
                    + param.GetValue("ruta_access_curvas",DateTime.Now, DateTime.Now)
                    + " --> "  + string.Format("{0:#,##0}", total_registros));


                strSql = "select  cups20, fuente, fecha, hora, estacion, ae, r1, r4"
                     + " from [EER_CURVAS CUARTO HORARIAS]";                     

                Console.WriteLine(strSql);

                ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas",DateTime.Now, DateTime.Now));
                cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {
                    

                    if (firstOnly)
                    {
                        sb.Append("replace into eer_curvas_cuarto_horarias_tmp");
                        sb.Append(" (cups20, fuente, fecha, hora, estacion,");
                        sb.Append("ae, r1, r4) values ");
                        firstOnly = false;
                    }

                    if (r["cups20"] != System.DBNull.Value &&
                        r["fecha"] != System.DBNull.Value &&
                        r["hora"] != System.DBNull.Value &&
                        r["estacion"] != System.DBNull.Value &&
                        r["ae"] != System.DBNull.Value &&
                        r["r1"] != System.DBNull.Value)
                    {

                        i++;
                        x++;

                        sb.Append("('").Append(r["cups20"].ToString()).Append("',");

                        if (r["fuente"] != System.DBNull.Value)
                            sb.Append("'").Append(r["fuente"].ToString()).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append("'").Append(Convert.ToDateTime(r["fecha"]).ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(Convert.ToDateTime(r["hora"]).ToString("HH:mm")).Append("',");

                        if (r["estacion"].ToString().Trim() != "")
                            sb.Append(r["estacion"].ToString()).Append(",");
                        else
                            sb.Append("0,");

                        if (r["ae"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["ae"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");


                        if (r["r1"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["r1"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");

                        if (r["r4"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["r4"]).ToString().Replace(",", "."));
                        else
                            sb.Append("null");


                        sb.Append("),");

                        if (i == 100)
                        {
                            Console.CursorLeft = 0;
                            Console.Write(string.Format("{0:#,##0}", x) + " de " + string.Format("{0:#,##0}", total_registros));

                            i = 0;
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.CON);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                        }
                    }
                }

                if (i > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write(x + " de " + total_registros);
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;                    
                }

                ac.CloseConnection();

                //MessageBox.Show("Copia completada con éxito", "Copia de Curva Access", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception e)
            {
                // MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(e.Message);
            }
            
        }
        public  void CopiaDatosPeaje()
        {
            Dictionary<string, List<EndesaEntity.DatosPeaje>> d_peajes_access;
            Dictionary<string, List<EndesaEntity.DatosPeaje>> d_peajes_mysql;            

            try
            {

                d_peajes_access = CargaDatosPeajeMedidaAccess();
                d_peajes_mysql = CargaDatosPeajesMySQL();

                foreach(KeyValuePair<string, List<EndesaEntity.DatosPeaje>> p in d_peajes_access)
                {
                    for (int i = 0; i < p.Value.Count(); i++)
                    {
                        List<EndesaEntity.DatosPeaje> o;
                        if (!d_peajes_mysql.TryGetValue(p.Key, out o))
                        {
                            InsertPeajeMySQL(p.Value[i], 1);
                        }
                        else
                        {
                            EndesaEntity.DatosPeaje datosPeaje = new EndesaEntity.DatosPeaje();
                            for (int x = 0; x < o.Count(); x++)
                            {

                                if(o[x].fecha_desde == p.Value[i].fecha_desde &&
                                   o[x].fecha_hasta == p.Value[i].fecha_hasta)
                                  
                                {
                                    datosPeaje = o[x];
                                    break;
                                }

                            }

                            if(datosPeaje.cups20 != null)
                            {
                                if (datosPeaje.importe_termino_potencia != p.Value[i].importe_termino_potencia)
                                    InsertPeajeMySQL(p.Value[i], datosPeaje.version + 1); 
                            }
                            else
                            {
                                if(p.Value[i].fecha_desde <= p.Value[i].fecha_hasta)
                                    InsertPeajeMySQL(p.Value[i], 1);
                            }
                           

                        }
                    }

                        

                }


            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }
        public void CopiaCurvasCuartoHorariasAccess()
        {
            int i = 0;
            string strSql = "";
            servidores.AccessDB ac;
            OleDbDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();


            MySQLDB db;
            MySqlCommand command;


            EndesaBusiness.medida.CurvasEER cc = new medida.CurvasEER();
            List<string> lista_cups20 = new List<string>();
            EndesaBusiness.eer.Inventario inventario = new EndesaBusiness.eer.Inventario(DateTime.MinValue, DateTime.MaxValue);

            try
            {
                strSql = "SELECT cups20, min(fecha) as min_fecha, max(fecha) as max_fecha"
                    + " from [EER_CURVAS CUARTO HORARIAS]"
                    + " group by cups20, YEAR(fecha), ";

                ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas", DateTime.Now, DateTime.Now));
                OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {

                    if (firstOnly)
                    {
                        fd = Convert.ToDateTime(r["min_fecha"]);
                        fh = Convert.ToDateTime(r["max_fecha"]);
                        firstOnly = false;
                    }

                    if (fd > Convert.ToDateTime(r["min_fecha"]))
                        fd = Convert.ToDateTime(r["min_fecha"]);

                    if (fh > Convert.ToDateTime(r["max_fecha"]))
                        fh = Convert.ToDateTime(r["max_fecha"]);

                    lista_cups20.Add(r["cups20"].ToString());
                }
                ac.CloseConnection();

                for (int j = 0; j < lista_cups20.Count(); j++)
                {
                    i = 0;
                    strSql = "select cups20, fuente, hora, ae,  from [EER_CURVAS CUARTO HORARIAS] where "
                        + " cups20 = '" + lista_cups20[j] + "'"
                        + "(fecha >= #" + fd.ToString("MM/dd/yyyy") + "# and fecha <= #" + fh.ToString("MM/dd/yyyy") + "#)"
                        + " order by cups20, fecha, hora";
                    ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas", DateTime.Now, DateTime.Now));
                    cmd = new OleDbCommand(strSql, ac.con);
                    r = cmd.ExecuteReader();
                    while (r.Read())
                    {
                        i++;
                        if (firstOnly)
                        {


                            //cc = new medida.CurvasEER(r["cups20"].ToString(), fd, fh);
                            cc = new medida.CurvasEER(inventario.GetPS(r["cups20"].ToString(), fd, fh), fd, fh);
                            firstOnly = false;
                        }

                        if (r["fuente"] != System.DBNull.Value)
                            cc.curvaCuartoHorariaFuente[i] = r["fuente"].ToString();

                        cc.curvaCuartoHorariaDias[i] = Convert.ToDateTime(r["fecha"]);
                        cc.curvaCuartoHorariaDias[i] = cc.curvaCuartoHorariaDias[i].AddHours(Convert.ToDateTime(r["hora"]).Hour);
                        cc.curvaCuartoHorariaDias[i] = cc.curvaCuartoHorariaDias[i].AddMinutes(Convert.ToDateTime(r["hora"]).Minute);

                        if (r["ae"] != System.DBNull.Value)
                        {
                            cc.curvaCuartoHorariaActiva[i] = Convert.ToDouble(r["ae"]);
                            cc.curvaCuartoHorariaPotencias[i] = Convert.ToDouble(r["ae"]) * 4;
                        }

                        if (r["r1"] != System.DBNull.Value)
                            cc.curvaCuartoHorariaReactiva[i] = Convert.ToDouble(r["r1"]);

                    }                                     

                    ac.CloseConnection();


                }

                MessageBox.Show("Copia completada con éxito", "Copia de Curva Access", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void TrataCurvas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            EndesaBusiness.eer.Inventario inventario;
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();

            calendarios.Calendario cal;
            EndesaBusiness.medida.CurvasEER curva;
            EndesaBusiness.medida.CurvaResumenEER cr;


            try
            {
                

                strSql = "select min(fecha) as min_fecha, max(fecha) as max_fecha"
                    + " from cont.eer_curvas_cuarto_horarias_tmp";
                    // + " where cups20 = 'ES0031408599774021XS'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fd = Convert.ToDateTime(r["min_fecha"]);
                    fh = Convert.ToDateTime(r["max_fecha"]);
                }
                db.CloseConnection();
                //inventario = new EndesaBusiness.eer.Inventario(DateTime.MinValue, DateTime.MaxValue);
                inventario = new EndesaBusiness.eer.Inventario(fd, fh);
                //fd = new DateTime(fd.Year, fd.Month, 01);
                //fh = fh.AddMonths(1).AddDays(-1);

                cr = new medida.CurvaResumenEER(fd, fh);

                strSql = "select t.cups20, min(t.fecha) as min_fecha, max(t.fecha) as max_fecha, "
                    + " sum(t.ae) as total_ae, sum(t.r1) as total_r1,"
                    + " count(*) as total_registros"
                    + " FROM eer_inventario i"
                    + " INNER JOIN eer_curvas_cuarto_horarias_tmp t ON"
                    + " t.cups20 = i.cups20 AND"
                    + " (t.fecha >= i.fecha_inicio AND t.fecha <= i.fecha_fin)"
                    // + " where i.cups20 = 'ES0031408456675001KA'"
                    + " GROUP BY i.cups20, i.fecha_inicio";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                   
                    fd = Convert.ToDateTime(r["min_fecha"]);                    
                    fh = Convert.ToDateTime(r["max_fecha"]);

                    Console.WriteLine("Buscando datos en eer_curvas_cuarto_horarias_tmp para "
                       + r["cups20"].ToString() + " " + fd.ToString("dd-MM-yyyy") + " ~ " + fh.ToString("dd-MM-yyyy"));

                    // EndesaEntity.punto_suministro.PuntoSuministro ps = inventario.GetPS(r["cups20"].ToString());
                    EndesaEntity.punto_suministro.PuntoSuministro ps = 
                        inventario.GetPS(r["cups20"].ToString(), fd, fh);
                    if (ps != null)
                    {
                        cal = new calendarios.Calendario(fd, fh);
                        curva = new medida.CurvasEER(ps, fd, fh, "R", true, true);
                        if (curva.curvaCompleta)
                        {

                            EndesaEntity.medida.CurvaResumen c = new EndesaEntity.medida.CurvaResumen();
                            c.cups20 = r["cups20"].ToString();
                            c.fd = fd;
                            c.fh = fh;
                            c.activa = curva.totalEnergiaActiva;
                            c.reactiva = curva.totalEnergiaReactiva;
                            c.tarifa = ps.tarifa.tarifa;
                            c.origen = "C";
                            c.completa = true;
                            c.num_periodos = Convert.ToInt32(r["total_registros"]);

                            cal.CalculaDatosMedida(ps,
                                curva.curvaCuartoHorariaActiva,
                                curva.curvaCuartoHorariaReactiva,
                                curva.curvaCuartoHorariaPotencias,
                                curva.curvaCuartoHorariaDias);

                            for (int pt = 1; pt < cal.energiaActivaPorPeriodo.Count(); pt++)
                                if (cal.energiaActivaPorPeriodo[pt] != 0)
                                    c.activa_periodo[pt] = cal.energiaActivaPorPeriodo[pt];

                            for (int pt = 1; pt < cal.energiaReactivaPorPeriodo.Count(); pt++)
                                if (cal.energiaReactivaPorPeriodo[pt] != 0)
                                    c.reactiva_periodo[pt] = cal.energiaReactivaPorPeriodo[pt];


                            for (int pt = 1; pt < cal.potenciasMaximasRegistradas.Count(); pt++)
                                if (cal.potenciasMaximasRegistradas[pt] != 0)
                                    c.potencias_maximas[pt] = cal.potenciasMaximasRegistradas[pt];

                            cr.Save(c);
                            CopiaCurva_Tmp_MySQL(c.cups20, c.fd, c.fh, cr.version_curva);
                            // BorrarCurvaAccess(c.cups20, c.fd, c.fh);

                        }
                        else
                        {
                            EndesaEntity.medida.CurvaResumen c = new EndesaEntity.medida.CurvaResumen();
                            c.cups20 = r["cups20"].ToString();
                            c.fd = Convert.ToDateTime(r["min_fecha"]);
                            c.fh = Convert.ToDateTime(r["max_fecha"]);
                            c.activa = Convert.ToDouble(r["total_ae"]);
                            c.reactiva = Convert.ToDouble(r["total_r1"]);
                            c.tarifa = ps.tarifa.tarifa;
                            c.origen = "C";
                            c.completa = false;
                            c.num_periodos = Convert.ToInt32(r["total_registros"]);
                            cr.Save(c);
                        }
                    }
                   

                }
                db.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void CopiaCurva_Tmp_MySQL(string cups20, DateTime fd, DateTime fh, int version)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            
            try
            {
                strSql = "replace into eer_curvas_cuarto_horarias"
                    + " SELECT cups20, fuente, fecha, hora, estacion, "
                    + version + " ,estado, ae, r1, r4,"
                    + " NULL AS codigo_factura, f_ult_mod"
                    + " FROM cont.eer_curvas_cuarto_horarias_tmp"
                    + " where cups20 = '" + cups20 + "' and"
                    + " (fecha >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha <= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        
        private void BorrarCurvaAccess(string cups20, DateTime fd, DateTime fh)
        {
            string strSql = "";
            servidores.AccessDB ac;
            OleDbDataReader r;
            OleDbCommand cmd;

            try
            {
                strSql = "delete from [EER_CURVAS CUARTO HORARIAS] where "
                        + " cups20 = '" + cups20 + "' and"
                        + " (fecha >= #" + fd.ToString("MM/dd/yyyy") + "# and fecha <= #" + fh.ToString("MM/dd/yyyy") + "#)";
                
                ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas", DateTime.Now, DateTime.Now));
                cmd = new OleDbCommand(strSql, ac.con);
                cmd.ExecuteNonQuery();
                ac.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private Dictionary<string, List<EndesaEntity.DatosPeaje>> CargaDatosPeajeMedidaAccess()
        {

            servidores.AccessDB ac;
            OleDbDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.DatosPeaje>> d =
                new Dictionary<string, List<EndesaEntity.DatosPeaje>>();

            try
            {

                strSql = "select  CUPS, [FECHA DESDE] as FD, [FECHA HASTA] as FH,"
                     + " [TERMINO POTENCIA A FACTURAR] as importe_termino_potencia,"
                     + " [EXC_POT (Tarifa 6X)] as importe_excesos_potencia,"
                     + " [EXCESOS DE REACTIVA] as importe_excesos_reactiva,"
                     + " AP1, AP2, AP3, AP4, AP5, AP6, RP1, RP2, RP3, RP4, RP5, RP6,"
                     + " PMAX1, PMAX2, PMAX3, PMAX4, PMAX5, PMAX6"
                     + " from [CONSUMOS PERIODOS]";

                Console.WriteLine(strSql);                
                ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas", DateTime.Now, DateTime.Now));
                OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.DatosPeaje c = new EndesaEntity.DatosPeaje();
                    c.cups20 = r["CUPS"].ToString();
                    c.fecha_desde = Convert.ToDateTime(r["FD"]);
                    c.fecha_hasta = Convert.ToDateTime(r["FH"]);
                    
                    if (r["importe_termino_potencia"] != System.DBNull.Value)
                        c.importe_termino_potencia = Convert.ToDouble(r["importe_termino_potencia"]);

                    if (r["importe_excesos_potencia"] != System.DBNull.Value)
                        c.importe_excesos_potencia = Convert.ToDouble(r["importe_excesos_potencia"]);

                    if (r["importe_excesos_reactiva"] != System.DBNull.Value)
                        c.importe_excesos_reactiva = Convert.ToDouble(r["importe_excesos_reactiva"]);

                    for (int i = 1; i <= 6; i++)
                        if (r["AP" + i] != System.DBNull.Value)
                            c.activa[i] = Convert.ToDouble(r["AP" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["RP" + i] != System.DBNull.Value)
                            c.reactiva[i] = Convert.ToDouble(r["RP" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["PMAX" + i] != System.DBNull.Value)
                            c.potmax[i] = Convert.ToDouble(r["PMAX" + i]);

                    List<EndesaEntity.DatosPeaje> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.DatosPeaje>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    else
                        o.Add(c);

                }
                ac.CloseConnection();
                return d;

            }catch(Exception e)
            {

                ficheroLog.AddError("CargaDatosPeajeMedidaAccess: " + e.Message);
                return null;
            }

        }

        private Dictionary<string, List<EndesaEntity.DatosPeaje>> CargaDatosPeajesMySQL()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.DatosPeaje>> d =
                new Dictionary<string, List<EndesaEntity.DatosPeaje>>();

            try
            {
                strSql = "select cups20, fd, fh, version, cod_fiscal, fecha_factura,"
                    + " a_p1, a_p2, a_p3, a_p4, a_p5, a_p6,"
                    + " r_p1, r_p2, r_p3, r_p4, r_p5, r_p6,"
                    + " potmax_1, potmax_2, potmax_3, potmax_4, potmax_5, potmax_6,"
                    + " importe_termino_potencia, importe_excesos_potencia, importe_excesos_reactiva"
                    + " from cont.eer_datos_medida";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.DatosPeaje c = new EndesaEntity.DatosPeaje();
                    c.cups20 = r["cups20"].ToString();
                    if (r["fd"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["fd"]);
                    if (r["fh"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["fh"]);

                    c.version = Convert.ToInt32(r["version"]);

                    for (int i = 1; i <= 6; i++)
                        if (r["a_p" + i] != System.DBNull.Value)
                            c.activa[i] = Convert.ToDouble(r["a_p" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["r_p" + i] != System.DBNull.Value)
                            c.reactiva[i] = Convert.ToDouble(r["r_p" + i]);

                    for (int i = 1; i <= 6; i++)
                        if (r["potmax_" + i] != System.DBNull.Value)
                            c.potmax[i] = Convert.ToDouble(r["potmax_" + i]);

                    if (r["importe_termino_potencia"] != System.DBNull.Value)
                        c.importe_termino_potencia = Convert.ToDouble(r["importe_termino_potencia"]);

                    if (r["importe_excesos_potencia"] != System.DBNull.Value)
                        c.importe_excesos_potencia = Convert.ToDouble(r["importe_excesos_potencia"]);

                    if (r["importe_excesos_reactiva"] != System.DBNull.Value)
                        c.importe_excesos_reactiva = Convert.ToDouble(r["importe_excesos_reactiva"]);

                    List<EndesaEntity.DatosPeaje> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.DatosPeaje>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    else
                        o.Add(c);
                }
                db.CloseConnection();
                return d;
            }
            
            catch (Exception e)
            {
                ficheroLog.AddError("CargaDatosPeajesMySQL: " + e.Message);
                return null;
            }
        }

        private void InsertPeajeMySQL(EndesaEntity.DatosPeaje p, int version)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "replace  into eer_datos_medida"
                    + " (cups20, fd, fh, version,"
                    + " a_p1, a_p2, a_p3, a_p4, a_p5, a_p6,"
                    + " r_p1, r_p2, r_p3, r_p4, r_p5, r_p6,"
                    + " potmax_1, potmax_2, potmax_3, potmax_4, potmax_5, potmax_6,"
                    + " importe_termino_potencia, importe_excesos_potencia, importe_excesos_reactiva)"
                    + " values";

                strSql += " ('" + p.cups20 + "',";
                strSql += "'" + p.fecha_desde.ToString("yyyy-MM-dd") + "',";
                strSql += "'" + p.fecha_hasta.ToString("yyyy-MM-dd") + "',";
                strSql += version + ",";

                for (int i = 1; i <= 6; i++)
                    strSql += p.activa[i].ToString().Replace(",",".") + ",";

                for (int i = 1; i <= 6; i++)
                    strSql += p.reactiva[i].ToString().Replace(",", ".") + ",";

                for (int i = 1; i <= 6; i++)
                    if (p.potmax[i] != 0)
                        strSql += p.potmax[i].ToString().Replace(",", ".") + ",";
                    else
                        strSql += "null,";

                if (p.importe_termino_potencia != 0)
                    strSql += p.importe_termino_potencia.ToString().Replace(",", ".") + ",";
                else
                    strSql += "null,";

                if (p.importe_excesos_potencia != 0)
                    strSql += p.importe_excesos_potencia.ToString().Replace(",", ".") + ",";
                else
                    strSql += "null,";

                if (p.importe_excesos_reactiva != 0)
                    strSql += p.importe_excesos_reactiva.ToString().Replace(",", ".") + ");";
                else
                    strSql += "null);";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }catch(Exception e)
            {
                ficheroLog.AddError("InsertPeajeMySQL: " + e.Message);
            }
        }

        //private bool BuscaPeaje(string cups20, DateTime fd, DateTime fh, )

    }
}
