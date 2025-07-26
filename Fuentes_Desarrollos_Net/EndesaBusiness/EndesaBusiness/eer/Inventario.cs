using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.eer
{
    public class Inventario: EndesaEntity.Inventario
    {
        Dictionary<string, EndesaEntity.Inventario> dic;
        Dictionary<string, EndesaEntity.Direccion> dic_dir_envio; // direccion envío
        Dictionary<string, EndesaEntity.Direccion> dic_dir_ps; // direccion punto de suminsitro
        Dictionary<string, EndesaEntity.Direccion> dic_dir_fact; // direccion facturacion        

        public List<EndesaEntity.eer.InventarioPuntosFacturador> lista_inventario { get; set; }
        PuntosMedida pm;
        EndesaBusiness.punto_suministro.Tarifa tarifa;
        EndesaBusiness.eer.PreciosEnergia precios_energia;

        public Inventario(DateTime fd, DateTime fh)
        {
            dic_dir_envio = Carga_Direccion_Envio(fd, fh);
            dic_dir_ps = Carga_Direccion_PS(fd, fh);
            pm = new PuntosMedida();
            //precios_energia = new PreciosEnergia(fd, fh);
            tarifa = new punto_suministro.Tarifa();
            dic = Carga(fd, fh);
            lista_inventario = CargaLista();

        }


        private List<EndesaEntity.eer.InventarioPuntosFacturador> CargaLista()
        {
            List<EndesaEntity.eer.InventarioPuntosFacturador> l
                = new List<EndesaEntity.eer.InventarioPuntosFacturador>();


            foreach(KeyValuePair<string, EndesaEntity.Inventario> p in dic)
            {
                foreach(KeyValuePair<string, EndesaEntity.punto_suministro.PuntoSuministro> pp in p.Value.dic_puntos_suministro)
                {
                    EndesaEntity.eer.InventarioPuntosFacturador c = new EndesaEntity.eer.InventarioPuntosFacturador();
                    c.nif = p.Value.nif;
                    c.cliente = p.Value.nombre_cliente;
                    c.fecha_consumo_desde = pp.Value.fecha_inicio;
                    c.fecha_consumo_hasta = pp.Value.fecha_fin;
                    c.cups20 = pp.Value.cups20;
                    c.tarifa = pp.Value.tarifa.tarifa;
                    c.tipo_punto_medida = pp.Value.tipo_punto_medida;
                    l.Add(c);
                }
            }


            return l;
        }

        public EndesaEntity.punto_suministro.PuntoSuministro GetPS(string nif, string cups20, DateTime fd, DateTime fh)
        {
            EndesaEntity.Inventario o;
            if (dic.TryGetValue(nif, out o))
            {
                foreach (KeyValuePair<string, EndesaEntity.punto_suministro.PuntoSuministro> p in o.dic_puntos_suministro)
                {
                    if (p.Value.cups20 == cups20 && (p.Value.fecha_inicio <= fh && p.Value.fecha_fin >= fd))
                        return p.Value;
                }
            }
            return null;
           
        }

        public EndesaEntity.punto_suministro.PuntoSuministro GetPS(string cups20, DateTime fd, DateTime fh)
        {
            string clave = "";

            clave = cups20 + "|" + fd.ToString("yyyyMMdd") + "|" + fh.ToString("yyyyMMdd");

            foreach (KeyValuePair<string, EndesaEntity.Inventario> p in dic)
            {
                EndesaEntity.punto_suministro.PuntoSuministro o;
                if (p.Value.dic_puntos_suministro.TryGetValue(clave, out o))
                    return o;
            }
            return null;
        }

        private Dictionary<string, EndesaEntity.Direccion> Carga_Direccion_Envio(DateTime fd, DateTime fh)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.Direccion> d
                = new Dictionary<string, EndesaEntity.Direccion>();


            try
            {
                strSql = "SELECT cups20, fecha_inicio, fecha_fin, dir_envio," 
                    + " created_by, created_date, last_update_by, last_update_date"
                    + " FROM cont.eer_direccion_envio "
                    + " where (fecha_inicio <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_fin >= '" + fh.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Direccion c = new EndesaEntity.Direccion();
                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["dir_envio"] != System.DBNull.Value)
                        c.direccion_completa = r["dir_envio"].ToString();

                    EndesaEntity.Direccion o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }
        private Dictionary<string, EndesaEntity.Direccion> Carga_Direccion_PS(DateTime fd, DateTime fh)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.Direccion> d
                = new Dictionary<string, EndesaEntity.Direccion>();


            try
            {
                strSql = "SELECT cups20, fecha_inicio, fecha_fin, dir_ps, codigo_postal,"
                    + " created_by, created_date, last_update_by, last_update_date"
                    + " FROM cont.eer_direccion_ps                "
                    + " where (fecha_inicio <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_fin >= '" + fh.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Direccion c = new EndesaEntity.Direccion();
                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["dir_ps"] != System.DBNull.Value)
                        c.direccion_completa = r["dir_ps"].ToString();

                    if (r["codigo_postal"] != System.DBNull.Value)
                        c.codigo_postal = r["codigo_postal"].ToString();

                    

                    EndesaEntity.Direccion o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public string GetDireccion_Envio(string cups20)
        {
            EndesaEntity.Direccion o;
            if (dic_dir_envio.TryGetValue(cups20, out o))
                return o.direccion_completa;
            else
                return "";
        }

        private string GetDireccion_PS(string cups20)
        {
            EndesaEntity.Direccion o;
            if (dic_dir_ps.TryGetValue(cups20, out o))
                return o.direccion_completa;
            else
                return "";
        }


        public void ActualizaCuotaPendiente(string cups20)
        {
            // Una vez se registra la factura asignado a la misma el codigo de factura y la fecha factura
            // se procede a la actualizacíon de la cuota pdte restando en 1 su valor
        }

        private Dictionary<string, EndesaEntity.Inventario> Carga(DateTime fd, DateTime fh)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime fecha_inicio = new DateTime();
            DateTime fecha_fin = new DateTime();

            string clave = "";

            Dictionary<string, EndesaEntity.Inventario> d
                = new Dictionary<string, EndesaEntity.Inventario>();

            try
            {
                strSql = "SELECT cif, nombre_cliente, cups20, fecha_inicio, fecha_fin, fecha_alta, fecha_baja,"
                    + " tarifa_acceso, potencia_1, potencia_2, potencia_3, potencia_4, potencia_5, potencia_6,"
                    + " medida_en_baja, porcentaje_perdidas, potencia_trafo, tipo_contrato, tipo_autoconsumo, cnae,"
                    + " sstt, importe_sstt, cuotas_sstt, cuotas_pdtes_sstt, alquiler, ser_gest_prefer, dto_te,"
                    + " exencion_ise, porcentaje_exencion_ise, tipo_punto_medida"
                    + " FROM cont.eer_inventario ";         
                
                if(fd != DateTime.MinValue && fh != DateTime.MaxValue)
                {
                    strSql += "where (fecha_inicio <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                        + " fecha_fin >= '" + fd.ToString("yyyy-MM-dd") + "')";

                    
                }
                    
                
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (Convert.ToDateTime(r["fecha_inicio"]) < fd)
                        fecha_inicio = fd;
                    else
                        fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (Convert.ToDateTime(r["fecha_fin"]) > fh)
                        fecha_fin = fh;
                    else
                        fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    clave = (r["cups20"].ToString() 
                        + "|" + fecha_inicio.ToString("yyyyMMdd") 
                        + "|" + fecha_fin.ToString("yyyyMMdd"));


                    EndesaEntity.Inventario o;
                    if (!d.TryGetValue(r["cif"].ToString(), out o))
                    {
                        EndesaEntity.Inventario c = new EndesaEntity.Inventario();

                        if (r["cif"] != System.DBNull.Value)
                            c.nif = r["cif"].ToString();

                        if (r["nombre_cliente"] != System.DBNull.Value)
                            c.nombre_cliente = r["nombre_cliente"].ToString();

                        if (r["fecha_inicio"] != System.DBNull.Value)
                            c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                        if (r["fecha_fin"] != System.DBNull.Value)
                            c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                        if (r["fecha_alta"] != System.DBNull.Value)
                            c.fecha_alta = Convert.ToDateTime(r["fecha_alta"]);

                        if (r["fecha_baja"] != System.DBNull.Value)
                            c.fecha_baja = Convert.ToDateTime(r["fecha_baja"]);

                        if (r["tipo_contrato"] != System.DBNull.Value)
                            c.tipo_contrato = r["tipo_contrato"].ToString();

                        if (r["tipo_autoconsumo"] != System.DBNull.Value)
                            c.tipo_autoconsumo = r["tipo_autoconsumo"].ToString();

                        if (r["cnae"] != System.DBNull.Value)
                            c.cnae = r["cnae"].ToString();


                            EndesaEntity.Direccion dir_envio;
                        if (dic_dir_envio.TryGetValue(clave, out dir_envio))
                            c.direccion_envio = dir_envio.direccion_completa;

                        if (r["cups20"] != System.DBNull.Value)
                        {
                            EndesaEntity.punto_suministro.PuntoSuministro ps = new EndesaEntity.punto_suministro.PuntoSuministro();                            
                            ps.cups20 = r["cups20"].ToString();
                            ps.lista_puntos_medida_principales = pm.GetPuntosMedida(ps.cups20);

                            if (Convert.ToDateTime(r["fecha_inicio"]) < fd)
                                ps.fecha_inicio = fd;
                            else
                                ps.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                            if (Convert.ToDateTime(r["fecha_fin"]) > fh)
                                ps.fecha_fin = fh;
                            else
                                ps.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                            if (r["sstt"] != System.DBNull.Value)
                                ps.sstt = r["sstt"].ToString() == "S";

                            if (ps.sstt)
                            {
                                if (r["importe_sstt"] != System.DBNull.Value)
                                    ps.importe_sstt = Convert.ToDouble(r["importe_sstt"]);

                                if (r["cuotas_sstt"] != System.DBNull.Value)
                                    ps.cuotas = Convert.ToInt32(r["cuotas_sstt"]);

                                if (r["cuotas_pdtes_sstt"] != System.DBNull.Value)
                                    ps.cuotas_pdtes = Convert.ToInt32(r["cuotas_pdtes_sstt"]);

                            }


                            if (r["tipo_punto_medida"] != System.DBNull.Value)
                                ps.tipo_punto_medida = Convert.ToInt32(r["tipo_punto_medida"]);

                            if (r["alquiler"] != System.DBNull.Value)
                                ps.alquiler = Convert.ToDouble(r["alquiler"]);

                            if (r["tarifa_acceso"] != System.DBNull.Value)                            
                                ps.tarifa = tarifa.GetTarifa(r["tarifa_acceso"].ToString());

                            if (r["ser_gest_prefer"] != System.DBNull.Value)
                                ps.servicio_gestion_preferente = Convert.ToDouble(r["ser_gest_prefer"]);

                            if (r["dto_te"] != System.DBNull.Value)
                                ps.dto_te = Convert.ToDouble(r["dto_te"]);


                            // Asignamos al punto los precios de la energia
                            //ps.precios_energia = precios_energia.GetPrecio(ps.cups20, );

                            for (int i = 1; i <= 6; i++)
                            {
                                if (r["potencia_" + i] != System.DBNull.Value)
                                    ps.potecias_contratadas[i] = Convert.ToInt32(r["potencia_" + i]);
                            }

                            if (r["medida_en_baja"] != System.DBNull.Value)
                                ps.medida_en_baja = r["medida_en_baja"].ToString() == "S";
                            if (ps.medida_en_baja)
                            {
                                if (r["porcentaje_perdidas"] != System.DBNull.Value)
                                    ps.porcetaje_perdidas = Convert.ToDouble(r["porcentaje_perdidas"]);

                                if (r["potencia_trafo"] != System.DBNull.Value)
                                    ps.potencia_trafo = Convert.ToDouble(r["potencia_trafo"]);
                            }

                            #region ISE
                            if(r["exencion_ise"].ToString() == "S")
                            {                                
                                if (r["porcentaje_exencion_ise"] != System.DBNull.Value)
                                {
                                    ps.exencion_ise = true;
                                    ps.porcentaje_exencion = Convert.ToDouble(r["porcentaje_exencion_ise"]);
                                }
                                    

                            }
                            #endregion

                            EndesaEntity.Direccion oo;
                            if (dic_dir_ps.TryGetValue(ps.cups20, out oo))
                                ps.direccion = oo;
                            

                            EndesaEntity.punto_suministro.PuntoSuministro pso;
                            if (!c.dic_puntos_suministro.TryGetValue(clave, out pso))
                                c.dic_puntos_suministro.Add(clave, ps);                            
                        }


                        d.Add(c.nif, c);
                    }
                    else
                    {
                        if (r["cups20"] != System.DBNull.Value)
                        {


                            EndesaEntity.punto_suministro.PuntoSuministro pso;
                            if (!o.dic_puntos_suministro.TryGetValue(clave, out pso))
                            {
                                EndesaEntity.punto_suministro.PuntoSuministro ps = new EndesaEntity.punto_suministro.PuntoSuministro();

                                ps.cups20 = r["cups20"].ToString();

                                if (Convert.ToDateTime(r["fecha_inicio"]) < fd)
                                    ps.fecha_inicio = fd;
                                else
                                    ps.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                                if (Convert.ToDateTime(r["fecha_fin"]) > fh)
                                    ps.fecha_fin = fh;
                                else
                                    ps.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);


                                ps.lista_puntos_medida_principales = pm.GetPuntosMedida(ps.cups20);

                                if (r["tarifa_acceso"] != System.DBNull.Value)
                                    ps.tarifa = tarifa.GetTarifa(r["tarifa_acceso"].ToString());

                                // Asignamos al punto los precios de la energia
                                //ps.precios_energia = precios_energia.GetPrecio(ps.cups20);

                                for (int i = 1; i <= 6; i++)
                                {
                                    if (r["potencia_" + i] != System.DBNull.Value)
                                        ps.potecias_contratadas[i] = Convert.ToInt32(r["potencia_" + i]);
                                }

                                if (r["medida_en_baja"] != System.DBNull.Value)
                                    ps.medida_en_baja = r["medida_en_baja"].ToString() == "S";
                                if (ps.medida_en_baja)
                                {
                                    if (r["porcentaje_perdidas"] != System.DBNull.Value)
                                        ps.porcetaje_perdidas = Convert.ToDouble(r["porcentaje_perdidas"]);

                                    if (r["potencia_trafo"] != System.DBNull.Value)
                                        ps.potencia_trafo = Convert.ToDouble(r["potencia_trafo"]);
                                }

                                if (r["sstt"] != System.DBNull.Value)
                                    ps.sstt = r["sstt"].ToString() == "S";

                                if (ps.sstt)
                                {
                                    if (r["importe_sstt"] != System.DBNull.Value)
                                        ps.importe_sstt = Convert.ToDouble(r["importe_sstt"]);

                                    if (r["cuotas_sstt"] != System.DBNull.Value)
                                        ps.cuotas = Convert.ToInt32(r["cuotas_sstt"]);

                                    if (r["cuotas_pdtes_sstt"] != System.DBNull.Value)
                                        ps.cuotas_pdtes = Convert.ToInt32(r["cuotas_pdtes_sstt"]);

                                }

                                if (r["alquiler"] != System.DBNull.Value)
                                    ps.alquiler = Convert.ToDouble(r["alquiler"]);


                                if (r["ser_gest_prefer"] != System.DBNull.Value)
                                    ps.servicio_gestion_preferente = Convert.ToDouble(r["ser_gest_prefer"]);

                                if (r["dto_te"] != System.DBNull.Value)
                                    ps.dto_te = Convert.ToDouble(r["dto_te"]);

                                if (r["tipo_punto_medida"] != System.DBNull.Value)
                                    ps.tipo_punto_medida = Convert.ToInt32(r["tipo_punto_medida"]);


                                EndesaEntity.Direccion oo;
                                if (dic_dir_ps.TryGetValue(ps.cups20, out oo))
                                    ps.direccion = oo;                                                              
                                
                                o.dic_puntos_suministro.Add(clave, ps);
                            }
                        }


                    }
                    
                    

                }
                db.CloseConnection();

                return d;
            }
            catch(Exception e)
            {
                return null;
            }
        }

        public EndesaEntity.Inventario GetCliente(string cups20)
        {
            foreach(KeyValuePair<string, EndesaEntity.Inventario> p in dic)
            {
                EndesaEntity.punto_suministro.PuntoSuministro ps;
                if (p.Value.dic_puntos_suministro.TryGetValue(cups20, out ps))
                    return p.Value;
            }

            return null;

        }

        public EndesaEntity.Inventario GetCliente(string cups20, DateTime fd, DateTime fh)
        {
            string clave = "";

            clave = cups20 + "|" + fd.ToString("yyyyMMdd") + "|" + fh.ToString("yyyyMMdd");
            foreach (KeyValuePair<string, EndesaEntity.Inventario> p in dic)
            {
                EndesaEntity.punto_suministro.PuntoSuministro ps;
                if (p.Value.dic_puntos_suministro.TryGetValue(clave, out ps))
                    return p.Value;
            }

            return null;

        }

        


        //public void GetCustomer_From_CUPS20(string cups20)
        //{
        //    EndesaEntity.Inventario o;
        //    if(dic.TryGetValue(cups20, out o))
        //    {
        //        this.nif = o.nif;
        //        this.nombre_cliente = o.nombre_cliente;
        //        this.cups20 = o.cups20;
        //        this.tarifa = o.tarifa;
        //        this.potecias_contratadas = o.potecias_contratadas;
        //    }
        //}

    }
}

