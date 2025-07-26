using EndesaBusiness.global;
using EndesaBusiness.servidores;
using EndesaEntity.cnmc;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types;

namespace EndesaBusiness.cnmc
{
    public class CNMC
    {
        
        Dictionary<string, string> dic_tipo_producto;
        List<EndesaEntity.cnmc.PeajeGas> lista_tipo_peajes;
        Dictionary<string, EndesaEntity.cnmc.PeajeGas> dic_peajes_gas;

        public Dictionary<string, string> dic_solicitud;
        public Dictionary<string, string> dic_ambito_validez_cie;
        public Dictionary<string, string> dic_cnae;
        public Dictionary<string, string> dic_direccion_fiscal;
        public Dictionary<string, string> dic_escalera;
        public Dictionary<string, string> dic_indicativo_activacion;
        public Dictionary<string, string> dic_indicativo_sino;
        public Dictionary<string, string> dic_marca;
        public Dictionary<string, string> dic_modelo_fecha_efecto;
        public Dictionary<string, string> dic_modo_control_potencia;
        public Dictionary<string, string> dic_nacionalidad;
        public Dictionary<string, string> dic_piso;
        public Dictionary<string, string> dic_procesos;
        public Dictionary<string, string> dic_propiedad_equipo;
        public Dictionary<string, string> dic_puerta;
        public Dictionary<string, string> dic_sujetos;
        public Dictionary<string, string> dic_tarifa_atr;
        public Dictionary<string, string> dic_tensiones;
        public Dictionary<string, string> dic_aclarador_finca;
        public Dictionary<string, string> dic_tipo_aparato;
        public Dictionary<string, string> dic_autoconsumo;
        public Dictionary<string, string> dic_contrato_atr;
        public Dictionary<string, string> dic_documentacion;
        public Dictionary<string, string> dic_documento;
        public Dictionary<string, string> dic_equipo_medida;
        public Dictionary<string, string> dic_identificador;
        public Dictionary<string, string> dic_peajes;
        public Dictionary<string, string> dic_persona;
        public Dictionary<string, string> dic_producto;
        public Dictionary<string, string> dic_solicitud_producto;
        public Dictionary<string, string> dic_tipo_suministro;
        public Dictionary<string, string> dic_via;
        public Dictionary<string, string> dic_provincias;
        public Dictionary<string, string> dic_distribuidoras;
        public Dictionary<string, string> dic_tipo_activacion_prevista;
        public Dictionary<string, string> dic_municipios;
        public Dictionary<string, EndesaEntity.cnmc.Motivo_Rechazo> dic_motivos_rechazo;

        public CNMC()
        {
            dic_tipo_producto = Carga_Tipo_Producto();
            lista_tipo_peajes = Carga_Tipo_Peajes();

            dic_solicitud = Carga_Tabla_CNMC("cnmc_p_tipo_solicitud");
            dic_ambito_validez_cie = Carga_Tabla_CNMC("cnmc_p_ambito_validez_cie");
            dic_cnae = Carga_Tabla_CNMC_Codigos("cnmc_p_cnae");
            dic_direccion_fiscal = Carga_Tabla_CNMC("cnmc_p_direccion_fiscal");
            dic_escalera = Carga_Tabla_CNMC("cnmc_p_escalera");
            dic_indicativo_activacion = Carga_Tabla_CNMC("cnmc_p_indicativo_activacion");
            dic_indicativo_sino = Carga_Tabla_CNMC("cnmc_p_indicativo_sino");
            dic_marca = Carga_Tabla_CNMC("cnmc_p_marca");
            // dic_modelo_fecha_efecto = Carga_Tabla_CNMC("cnmc_p_modelo_fecha_efecto");
            dic_modo_control_potencia = Carga_Tabla_CNMC("cnmc_p_modo_control_potencia");
            dic_nacionalidad = Carga_Tabla_CNMC("cnmc_p_nacionalidad");
            dic_piso = Carga_Tabla_CNMC("cnmc_p_piso");
            dic_procesos = Carga_Tabla_CNMC("cnmc_p_procesos");
            dic_propiedad_equipo = Carga_Tabla_CNMC("cnmc_p_propiedad_equipo");
            dic_puerta = Carga_Tabla_CNMC("cnmc_p_puerta");
            //dic_sujetos = Carga_Tabla_CNMC("cnmc_p_sujetos");
            dic_tarifa_atr = Carga_Tabla_CNMC("cnmc_p_tarifa_atr");
            dic_tensiones = Carga_Tabla_CNMC("cnmc_p_tensiones");
            dic_aclarador_finca = Carga_Tabla_CNMC("cnmc_p_tipo_aclarador_finca");
            dic_tipo_aparato = Carga_Tabla_CNMC("cnmc_p_tipo_aparato");
            dic_autoconsumo = Carga_Tabla_CNMC("cnmc_p_tipo_autoconsumo");
            dic_contrato_atr = Carga_Tabla_CNMC("cnmc_p_tipo_contrato_atr");
            dic_documentacion = Carga_Tabla_CNMC("cnmc_p_tipo_documentacion");
            dic_documento = Carga_Tabla_CNMC("cnmc_p_tipo_documento");
            dic_equipo_medida = Carga_Tabla_CNMC("cnmc_p_tipo_equipo_medida");
            dic_identificador = Carga_Tabla_CNMC("cnmc_p_tipo_identificador");
            // dic_peajes = Carga_Tabla_CNMC("cnmc_p_tipo_peajes");
            dic_persona = Carga_Tabla_CNMC("cnmc_p_tipo_persona");
            dic_producto = Carga_Tabla_CNMC("cnmc_p_tipo_producto");
            dic_solicitud_producto = Carga_Tabla_CNMC("cnmc_p_tipo_solicitud_producto");
            dic_via = Carga_Tabla_CNMC("cnmc_p_tipo_via");
            dic_provincias = Carga_Tabla_CNMC("cnmc_p_provincias");
            dic_distribuidoras = Carga_Tabla_CNMC("cnmc_p_distribuidoras");
            dic_tipo_activacion_prevista = Carga_Tabla_CNMC("cnmc_p_tipo_activacion_prevista");
            dic_municipios = Carga_Municipios();
            dic_motivos_rechazo = Carga_Motivos_Rechazo();
        }
        public CNMC(DateTime fd, DateTime fh)
        {
            dic_tipo_producto = Carga_Tipo_Producto();
            lista_tipo_peajes = Carga_Tipo_Peajes();
            dic_peajes_gas = Carga_Tipo_Peajes(fd, fh);
        }


        private Dictionary<string, string> Carga_Tabla_CNMC_Codigos(string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string codigo = "";
            string descripcion = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select codigo, descripcion from " + tabla
                    + " where (fecha_baja is null or fecha_baja > '"
                    + DateTime.Now.ToString("yyyy-MM-dd") + "')"
                    + " group by descripcion";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    codigo = r["codigo"].ToString();
                    descripcion = r["descripcion"].ToString();
                    d.Add(codigo, descripcion);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }
        private Dictionary<string, string> Carga_Tabla_CNMC(string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string codigo = "";
            string descripcion = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select codigo, descripcion from " + tabla
                    + " where (fecha_baja is null or fecha_baja > '"
                    + DateTime.Now.ToString("yyyy-MM-dd") + "')"
                    + " group by descripcion";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    codigo = r["codigo"].ToString();
                    descripcion = r["descripcion"].ToString();
                    d.Add(descripcion, codigo);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {                
                return null;
            }
        }

        public Dictionary<string, string> Carga_Municipios()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string cpro = "";
            string nombre = "";
            string nombre_sin_tilde = "";
            string codigo_municipio = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select cpro, nombre, codigo_municipio from cnmc_p_municipios";                    
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {                    

                    string[] separado = r["nombre"].ToString().ToLower().Split('/');

                    for(int i = 0; i < separado.Length; i++)
                    {
                        nombre = separado[i];
                        nombre_sin_tilde = utilidades.FuncionesTexto.SinTildes(separado[i]);

                        cpro = r["cpro"].ToString();
                        codigo_municipio = r["codigo_municipio"].ToString();

                        string o;
                        if (!d.TryGetValue(cpro + "|" + nombre, out o))
                            d.Add(cpro + "|" + nombre, codigo_municipio);

                        if (!d.TryGetValue(cpro + "|" + nombre_sin_tilde, out o))
                            d.Add(cpro + "|" + nombre_sin_tilde, codigo_municipio);

                    }
                    
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public List<string> GetLista_Motivos_Rechazo(string proceso)
        {
            List<string> lista = new List<string>();

            foreach(KeyValuePair<string, EndesaEntity.cnmc.Motivo_Rechazo> p in dic_motivos_rechazo)
            {
                switch (proceso.ToUpper())
                {
                    case "A3":
                        if(p.Value.a3)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "B1":
                        if (p.Value.b1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "C1":
                        if (p.Value.c1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "C2":
                        if (p.Value.c2)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "M1":
                        if (p.Value.m1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "W1":
                        if (p.Value.w1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "R1":
                        if (p.Value.r1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "E1":
                        if (p.Value.e1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "T1":
                        if (p.Value.t1)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "P0":
                        if (p.Value.p0)
                            lista.Add(p.Value.descripcion);
                        break;
                    case "D1":
                        if (p.Value.d1)
                            lista.Add(p.Value.descripcion);
                        break;
                }



            }

            return lista;

        }

        public string GetCodigo_Motivo_Rechazo(string descripcion)
        {
            string codigo = "";
            foreach(KeyValuePair<string, EndesaEntity.cnmc.Motivo_Rechazo> p in dic_motivos_rechazo)
            {
                if (p.Value.descripcion == descripcion)
                {
                    codigo = p.Key;
                    break;
                }                    

            }

            return codigo;
        }
        private Dictionary<string, EndesaEntity.cnmc.Motivo_Rechazo> Carga_Motivos_Rechazo()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;            

            Dictionary<string, EndesaEntity.cnmc.Motivo_Rechazo> d =
                new Dictionary<string, EndesaEntity.cnmc.Motivo_Rechazo>();
            try
            {
                strSql = "SELECT codigo, descripcion, a3, b1, c1, c2, m1, w1, r1, e1, t1, p0, d1, fecha_baja"
                    + " FROM cont.cnmc_p_motivos_rechazo";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cnmc.Motivo_Rechazo c = new EndesaEntity.cnmc.Motivo_Rechazo();
                    c.codigo = r["codigo"].ToString();
                    c.descripcion = r["descripcion"].ToString();
                    c.a3 = r["a3"] != System.DBNull.Value;
                    c.b1 = r["b1"] != System.DBNull.Value;
                    c.c1 = r["c1"] != System.DBNull.Value;
                    c.c2 = r["c2"] != System.DBNull.Value;
                    c.m1 = r["m1"] != System.DBNull.Value;
                    c.w1 = r["w1"] != System.DBNull.Value;
                    c.r1 = r["r1"] != System.DBNull.Value;
                    c.e1 = r["e1"] != System.DBNull.Value;
                    c.t1 = r["t1"] != System.DBNull.Value;
                    c.p0 = r["p0"] != System.DBNull.Value;
                    c.d1 = r["d1"] != System.DBNull.Value;

                    d.Add(c.codigo, c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public Dictionary<string, string> Carga_Poblaciones(Dictionary<string, string> dic, string provincia, string municipio)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string codigo = "";
            string descripcion = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select codigo, descripcion from cnmc_p_poblacion"
                    + " where codigo like '"
                    + GetCodigo("cnmc_p_provincias", provincia)
                    + GetCodigoPoblacion(dic, municipio) + "%'"
                    + " and descripcion <> '*DISEMINADO*'"
                    + " group by descripcion";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    codigo = r["codigo"].ToString();
                    descripcion = r["descripcion"].ToString();
                    d.Add(codigo, descripcion);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private Dictionary<string, string> Carga_Tipo_Producto()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string codigo = "";
            string descripcion = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select codigo, descripcion from cont.cnmc_p_tipo_producto;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    codigo = r["codigo"].ToString();
                    descripcion = r["descripcion"].ToString().ToUpper();
                    d.Add(descripcion, codigo);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }
        private List<EndesaEntity.cnmc.PeajeGas> Carga_Tipo_Peajes()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;            

            List<EndesaEntity.cnmc.PeajeGas> l = new List<EndesaEntity.cnmc.PeajeGas>();
            try
            {
                strSql = "select tarifa, codigo, grupo_presion, consumo_anual_desde, consumo_anual_hasta"
                    + " from cont.cnmc_p_tipo_peajes";                    
                    
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cnmc.PeajeGas c = new EndesaEntity.cnmc.PeajeGas();
                    c.tarifa = r["tarifa"].ToString();
                    c.codigo = r["codigo"].ToString();
                    c.grupo_presion = Convert.ToInt32(r["grupo_presion"]);
                    c.consumo_anual_desde = Convert.ToDouble(r["consumo_anual_desde"]);
                    if (r["consumo_anual_hasta"] != System.DBNull.Value)
                        c.consumo_anual_hasta = Convert.ToDouble(r["consumo_anual_hasta"]);
                    else
                        c.consumo_anual_hasta = Double.MaxValue;
                    
                    l.Add(c);
                }
                db.CloseConnection();
                return l;
            }
            catch (Exception e)
            {
                return null;
            }

        }
        private Dictionary<string, EndesaEntity.cnmc.PeajeGas> Carga_Tipo_Peajes(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.cnmc.PeajeGas> d = new Dictionary<string, PeajeGas>();

            try
            {
                strSql = "select tarifa, codigo, grupo_presion, consumo_anual_desde, consumo_anual_hasta"
                    + " from cont.cnmc_p_tipo_peajes where"
                    + " (fecha_desde <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + fh.ToString("yyyy-MM-dd") + "')";

                
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cnmc.PeajeGas c = new EndesaEntity.cnmc.PeajeGas();
                    c.tarifa = r["tarifa"].ToString();
                    c.codigo = r["codigo"].ToString();
                    c.grupo_presion = Convert.ToInt32(r["grupo_presion"]);
                    c.consumo_anual_desde = Convert.ToDouble(r["consumo_anual_desde"]);
                    if (r["consumo_anual_hasta"] != System.DBNull.Value)
                        c.consumo_anual_hasta = Convert.ToDouble(r["consumo_anual_hasta"]);
                    else
                        c.consumo_anual_hasta = Double.MaxValue;

                    EndesaEntity.cnmc.PeajeGas o;
                    if (!d.TryGetValue(c.codigo, out o))
                        d.Add(c.codigo, c);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }

        }

        
        public string Codigo_Tipo_Producto(string producto)
        {
            string o;
            if (dic_tipo_producto.TryGetValue(producto.ToUpper(), out o))
                return o;
            else
                return "";
        }
        public string Codigo_Tipo_Peaje(int grupo_presion, double consumo_anual)       {

            // Ahora el grupo de presion
            // esta determinado por la distribuidora.

            EndesaEntity.cnmc.PeajeGas p =
                lista_tipo_peajes.Find(z =>
                (z.grupo_presion == grupo_presion) && 
                (consumo_anual >= z.consumo_anual_desde && consumo_anual <= z.consumo_anual_hasta));

            if (p != null)
                return p.codigo;
            else
                return "21";
        }
        public string Tarifa(int grupo_presion, double consumo_anual)
        {
            EndesaEntity.cnmc.PeajeGas p =
               lista_tipo_peajes.Find(z =>
               (z.grupo_presion == grupo_presion) &&
               (consumo_anual >= z.consumo_anual_desde && consumo_anual <= z.consumo_anual_hasta));

            if (p != null)
                return p.tarifa;
            else
                return "2.1";
        }
        
        public string Distribuidora(string IDU)
        {
            string o;
            if (dic_distribuidoras.TryGetValue(IDU.ToUpper(), out o))
                return o;
            else
                return "";
        }

        public string Codigo_Tipo_Peaje(string tarifa)
        {
            EndesaEntity.cnmc.PeajeGas p =
               lista_tipo_peajes.Find(z => (z.tarifa == tarifa));

            if (p != null)
                return p.codigo;
            else
                return "";
        }

        public string GetCodigoMunicipio(string codigo_postal, string municipio)
        {
            string codigo_municipio = "";
            if (dic_municipios.TryGetValue(codigo_postal + "|" + municipio.ToLower(), out codigo_municipio))
                return codigo_municipio;
            else
                return "";

        }

        public string Tarifa_CNMC_a_Tarifa_SIGAME(string tarifa_cnmc)
        {
            EndesaEntity.cnmc.PeajeGas p =
               lista_tipo_peajes.Find(z => (z.codigo == tarifa_cnmc));

            if (p != null)
                return p.tarifa;
            else
                return "";
        }

        public string GetCodigo(string tabla, string descripcion)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            switch (tabla)
            {
                case "cnmc_p_provincias":
                    dic = dic_provincias;
                    break;
                case "cnmc_p_cnae":
                    dic = dic_cnae;
                    break;
                case "cnmc_p_tipo_activacion_prevista":
                    dic = dic_tipo_activacion_prevista;
                    break;
                case "cnmc_p_tarifa_atr":
                    dic = dic_tarifa_atr;
                    break;
            }

            string o;
            if (dic.TryGetValue(descripcion, out o))
                return o;
            else
                return null;

        }
        public string GetTipo_Autoconsumo(string autoconsumo)
        {
            string o;
            if (dic_autoconsumo.TryGetValue(autoconsumo, out o))
                return o;
            else
                return null;
        }


        public string GetCodigoPoblacion(Dictionary<string, string> dic, string municipio)
        {
            string o;
            if (dic.TryGetValue(municipio.ToLower(), out o))
                return o;
            else
                return null;
        }

        public string GetTipo_Identificador(string tipo_identificador)
        {
            string o;
            if (dic_identificador.TryGetValue(tipo_identificador, out o))
                return o;
            else
                return null;
        }

        public bool Es_Tarifa_GAS_XXI(string codigo_tarifa_cnmc)
        {
            bool es_tarifa_gas_xxi = false;
            EndesaEntity.cnmc.PeajeGas o;
            if (dic_peajes_gas.TryGetValue(codigo_tarifa_cnmc, out o))
                es_tarifa_gas_xxi = o.gas_xxi;

            return es_tarifa_gas_xxi;

        }



    }
}

