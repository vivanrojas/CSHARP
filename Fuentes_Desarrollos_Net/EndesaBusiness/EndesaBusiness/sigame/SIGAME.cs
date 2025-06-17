using EndesaBusiness.servidores;
using EndesaEntity;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.sigame
{
    public class SIGAME : EndesaEntity.sigame.ContratoGas
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "SIGAME");
        public Dictionary<string, EndesaEntity.sigame.ContratoGas> dic { get; set; }
        public Dictionary<string, string> dic_nif_vigentes { get; set; }

        public EndesaEntity.sigame.Consumos consumos { get; set; }
        Dictionary<int, List<int>> dic_pm;
        Dictionary<int, List<EndesaEntity.sigame.Consumos>> dic_consumos;
        Dictionary<int, List<EndesaEntity.sigame.Uso_Gas>> dic_uso_gas;
        Dictionary<string, EndesaEntity.sigame.Factura> dic_facturas;
        Dictionary<string, List<EndesaEntity.sigame.Factura>> dic_facturas_emitidas;

        public SIGAME()
        {            
            dic = new Dictionary<string, EndesaEntity.sigame.ContratoGas>();
            dic_nif_vigentes = new Dictionary<string, string>();
            Carga();
        }

        public SIGAME(DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, EndesaEntity.sigame.ContratoGas>();
            consumos = new EndesaEntity.sigame.Consumos();
            dic_pm = new Dictionary<int, List<int>>();
            dic_consumos = new Dictionary<int, List<EndesaEntity.sigame.Consumos>>();
            dic_uso_gas = new Dictionary<int, List<EndesaEntity.sigame.Uso_Gas>>();
            
            CargaPuntos();
            CargaConsumos(fd, fh);
            Carga(fd, fh);
            
        }

        public SIGAME(DateTime fd, DateTime fh, bool cuadro_mando_facturacion)
        {
            dic = new Dictionary<string, EndesaEntity.sigame.ContratoGas>();
            CargaCM(fd, fh);
        }

        public SIGAME(DateTime fd, DateTime fh, bool facturas, string entorno)
        {
            dic_facturas = new Dictionary<string, EndesaEntity.sigame.Factura>();
            CargaFacturas(fd, fh, entorno);
        }

        public SIGAME(DateTime fd, DateTime fh, bool facturas, bool facturas_emitidas_sigame, string entorno)
        {
            dic_facturas_emitidas = new Dictionary<string, List<EndesaEntity.sigame.Factura>>();
            CargaFacturasEmitidas(fd, fh, entorno);
        }
                


        private void CargaFacturas(DateTime fd, DateTime fh, string entorno)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            int id = 0;
                       

            try
            {

                DateTime fecha_desde = fd;
                fecha_desde = fd.AddMonths(-1);

                #region Query
                strSql = "select CD_CUPS, DE_NOMBRE_CLIENTE, CD_NFACTURA_REALES_PS, FH_FACTURA, FH_INI_FACTURACION, FH_FIN_FACTURACION"
                    + " ,T_SGM_M_FACTURAS_REALES_PS.NM_CONSUMO, NM_IMPORTE_BRUTO, TX_TIPO_FACTURA_NUEVO,TX_ORIGEN"
                    + " ,case when right(TX_ORIGEN,2)= '_E' then 'ESTIMADA' else 'REAL' end as MEDIDA"
                    + " from T_SGM_G_PS"
                    + " inner join T_SGM_G_CONTRATOS_PS"
                    + " on T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                    + " inner"
                    + " join T_SGM_M_CLIENTES"
                    + " on T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " inner"
                    + " join dbo.T_SGM_M_FACTURAS_REALES_PS"
                    + " on T_SGM_G_CONTRATOS_PS.ID_CTO_PS = T_SGM_M_FACTURAS_REALES_PS.ID_CTO_PS"
                    + " inner"
                    + " join T_SGM_M_PUNTOS_MEDIDA"
                    + " on T_SGM_G_PS.id_ps = T_SGM_M_PUNTOS_MEDIDA.ID_PS"
                    + " inner"
                    + " join T_SGM_M_LECTURAS_CONSUMOS"
                    + " ON T_SGM_M_PUNTOS_MEDIDA.ID_PMEDIDA = T_SGM_M_LECTURAS_CONSUMOS.ID_PMEDIDA"  
                    + " AND T_SGM_M_LECTURAS_CONSUMOS.TX_MES_LECTURA = FORMAT(FH_INI_FACTURACION,'yyyyMM')"
                    + " where T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL <= FH_INI_FACTURACION"
                    + " and (T_SGM_G_CONTRATOS_PS.FH_FIN_REAL >= FH_FIN_FACTURACION or T_SGM_G_CONTRATOS_PS.FH_FIN_REAL is null)"
                    + " and T_SGM_M_PUNTOS_MEDIDA.FH_INICIO <= FH_INI_FACTURACION"
                    + " and (T_SGM_M_PUNTOS_MEDIDA.FH_FIN >= FH_FIN_FACTURACION or T_SGM_M_PUNTOS_MEDIDA.FH_FIN is null)"
                    // + " AND TX_MES_LECTURA = '" + fd.ToString("yyyyMM") + "'"
                    + " and year(FH_FACTURA)= " + fd.ToString("yyyy")
                    + " and month(FH_FACTURA)= " + fd.ToString("MM")
                    + " and left(CD_CUPS,2)= '" + entorno + "'"
                    + " group by CD_CUPS, DE_NOMBRE_CLIENTE, CD_NFACTURA_REALES_PS, FH_FACTURA, FH_INI_FACTURACION, FH_FIN_FACTURACION"
                    + " ,T_SGM_M_FACTURAS_REALES_PS.NM_CONSUMO, NM_IMPORTE_BRUTO, TX_TIPO_FACTURA_NUEVO, TX_ORIGEN";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.Factura c = new EndesaEntity.sigame.Factura();

                    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        c.cfactura = r["CD_NFACTURA_REALES_PS"].ToString();

                    if (r["TX_ORIGEN"] != System.DBNull.Value)
                        c.fuente = r["TX_ORIGEN"].ToString();
                    if (r["MEDIDA"] != System.DBNull.Value)
                        c.medida = r["MEDIDA"].ToString();



                    EndesaEntity.sigame.Factura o;
                    if (!dic_facturas.TryGetValue(c.cfactura, out o))
                        dic_facturas.Add(c.cfactura, c);
                    

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(id);
                ficheroLog.AddError("Carga: " + e.Message);
            }
        }

        private void CargaFacturasEmitidas(DateTime fd, DateTime fh, string entorno)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            int id = 0;


            try
            {

               

                #region Query
                strSql = "select T_SGM_G_CONTRATOS_PS.ID_PS, CD_CUPS, DE_NOMBRE_CLIENTE, CD_NFACTURA_REALES_PS, FH_FACTURA, FH_INI_FACTURACION, FH_FIN_FACTURACION"
                    + " ,T_SGM_M_FACTURAS_REALES_PS.NM_CONSUMO, NM_IMPORTE_BRUTO, TX_TIPO_FACTURA_NUEVO,TX_ORIGEN"
                    + " ,case when right(TX_ORIGEN,2)= '_E' then 'ESTIMADA' else 'REAL' end as MEDIDA,"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_ULT_ACTUALIZACION"
                    + " from T_SGM_G_PS"
                    + " inner join T_SGM_G_CONTRATOS_PS"
                    + " on T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                    + " inner"
                    + " join T_SGM_M_CLIENTES"
                    + " on T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " inner"
                    + " join dbo.T_SGM_M_FACTURAS_REALES_PS"
                    + " on T_SGM_G_CONTRATOS_PS.ID_CTO_PS = T_SGM_M_FACTURAS_REALES_PS.ID_CTO_PS"
                    + " inner"
                    + " join T_SGM_M_PUNTOS_MEDIDA"
                    + " on T_SGM_G_PS.id_ps = T_SGM_M_PUNTOS_MEDIDA.ID_PS"
                    + " inner"
                    + " join T_SGM_M_LECTURAS_CONSUMOS"
                    + " ON T_SGM_M_PUNTOS_MEDIDA.ID_PMEDIDA = T_SGM_M_LECTURAS_CONSUMOS.ID_PMEDIDA"
                    + " AND T_SGM_M_LECTURAS_CONSUMOS.TX_MES_LECTURA = FORMAT(FH_INI_FACTURACION,'yyyyMM')"
                    + " where T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL <= FH_INI_FACTURACION"
                    + " and (T_SGM_G_CONTRATOS_PS.FH_FIN_REAL >= FH_FIN_FACTURACION or T_SGM_G_CONTRATOS_PS.FH_FIN_REAL is null)"
                    + " and T_SGM_M_PUNTOS_MEDIDA.FH_INICIO <= FH_INI_FACTURACION"
                    + " and (T_SGM_M_PUNTOS_MEDIDA.FH_FIN >= FH_FIN_FACTURACION or T_SGM_M_PUNTOS_MEDIDA.FH_FIN is null)"
                    + " and FH_INI_FACTURACION >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " and FH_FIN_FACTURACION <= '" + fh.ToString("yyyy-MM-dd") + "'"
                    + " and left(CD_CUPS,2)= '" + entorno + "'"
                    + " group by T_SGM_G_CONTRATOS_PS.ID_PS, CD_CUPS, DE_NOMBRE_CLIENTE, CD_NFACTURA_REALES_PS, FH_FACTURA, FH_INI_FACTURACION, FH_FIN_FACTURACION"
                    + " ,T_SGM_M_FACTURAS_REALES_PS.NM_CONSUMO, NM_IMPORTE_BRUTO, TX_TIPO_FACTURA_NUEVO, TX_ORIGEN,"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_ULT_ACTUALIZACION"
                    + " ORDER BY CD_CUPS, FH_INI_FACTURACION DESC";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.Factura c = new EndesaEntity.sigame.Factura();

                    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        c.cfactura = r["CD_NFACTURA_REALES_PS"].ToString();

                    if (r["TX_ORIGEN"] != System.DBNull.Value)
                        c.fuente = r["TX_ORIGEN"].ToString();
                    if (r["MEDIDA"] != System.DBNull.Value)
                        c.medida = r["MEDIDA"].ToString();

                    if (r["FH_INI_FACTURACION"] != System.DBNull.Value)
                        c.fd = Convert.ToDateTime(r["FH_INI_FACTURACION"]);

                    if (r["FH_FIN_FACTURACION"] != System.DBNull.Value)
                        c.fh = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);

                    if (r["FH_FACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FH_FACTURA"]);

                    if (r["FH_ULT_ACTUALIZACION"] != System.DBNull.Value)
                        c.last_update_date = Convert.ToDateTime(r["FH_ULT_ACTUALIZACION"]);

                    List<EndesaEntity.sigame.Factura> o;
                    if (!dic_facturas_emitidas.TryGetValue(c.cfactura, out o))
                    {
                        o = new List<EndesaEntity.sigame.Factura>();
                        o.Add(c);
                        dic_facturas_emitidas.Add(c.cfactura, o);
                    }
                    else
                        o.Add(c);
                        


                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(id);
                ficheroLog.AddError("Carga: " + e.Message);
            }
        }

        private void Carga()
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            int id = 0;

            try
            {
                Console.WriteLine("Cargando datos de SIGAME");
                #region Query
                strSql = "SELECT DISTINCT"
                    + " T_SGM_G_PS.ID_PS, T_SGM_G_PS.DE_PS, T_SGM_M_CLIENTES.CD_CIF,T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
                    + " T_SGM_G_CONTRATOS_PS.ID_CTO_PS, T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL,"
                    + " T_SGM_G_CONTRATOS_PS.FH_FIN_REAL, T_SGM_G_CONTRATOS_PS.FH_FIN_PREVISTA, T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO,"
                    + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA, T_SGM_G_PS.CD_CUPS, T_SGM_M_GESTORES.DE_GESTOR,"
                    + " DE_NOMBRE_MUNICIPIO, T_SGM_P_PROVINCIAS.DE_PROVINCIA,T_SGM_G_PS.CD_PAIS,"
                    + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES, T_SGM_M_RED_DISTRIBUCION.DE_RED_DISTRIBUCION,"
                    + " (T_SGM_G_PS.NM_SUMA_CONSUMOS_MEN) / 12 AS PROMEDIO, T_SGM_G_PS.NM_ENERGIA_DIARIA_KWH,"
                    + " T_SGM_G_CONTRATOS_PS.FH_ULT_ACTUALIZACION"
                    + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
                        + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                    + " LEFT JOIN T_SGM_M_CLIENTES ON"
                        + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " LEFT JOIN T_SGM_M_GESTORES ON"
                        + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR "
                    + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
                        + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
                    + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
                        + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES "
                    + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON "
                        + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
                    + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
                        + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
                        + " T_SGM_G_PS.CD_PAIS  = T_SGM_P_PROVINCIAS.CD_PAIS"
                    + " LEFT JOIN T_SGM_P_MUNICIPIOS ON "
                        + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
                    + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.ContratoGas c = new EndesaEntity.sigame.ContratoGas();
                    if (r["ID_PS"] != System.DBNull.Value)
                    {
                        c.id_ps = Convert.ToInt32(r["ID_PS"]);
                        id = c.id_ps;                        
                    }
                        

                    if (r["ID_CTO_PS"] != System.DBNull.Value)
                        c.id_cto_ps = Convert.ToInt32(r["ID_CTO_PS"]);

                    if (r["DE_PS"] != System.DBNull.Value)
                        c.descripcion_ps = Convert.ToString(r["DE_PS"]);

                    if (r["DE_NOMBRE_CLIENTE"] != System.DBNull.Value)
                        c.nombre_cliente = Convert.ToString(r["DE_NOMBRE_CLIENTE"]);

                    //if (r["NM_ENERGIA_DIARIA_KWH"] != System.DBNull.Value)
                    //    c.qd = Convert.ToDouble(r["NM_ENERGIA_DIARIA_KWH"]);

                    c.nif = Convert.ToString(r["CD_CIF"]);

                    if (r["FH_INICIO_REAL"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["FH_INICIO_REAL"]);
                    else
                        c.fecha_inicio = DateTime.MinValue;

                    if (r["FH_FIN_REAL"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["FH_FIN_REAL"]);
                    else
                        c.fecha_fin = new DateTime(4999, 12, 31);

                    if (r["ID_ESTADO_CTO"] != System.DBNull.Value)
                        c.id_estado_contrato = Convert.ToInt32(r["ID_ESTADO_CTO"]);
                    else
                        c.id_estado_contrato = 0;

                    if (r["DE_TIPO_TARIFA"] != System.DBNull.Value)
                    {
                        c.tarifa = Convert.ToString(r["DE_TIPO_TARIFA"]);                        
                    }
                        
                        
                    if (r["CD_CUPS"] != System.DBNull.Value)
                    {

                        c.cupsree = Convert.ToString(r["CD_CUPS"]).Trim();
                        if (c.cupsree.Length > 2)
                        {
                            c.pais = c.cupsree.Substring(0, 2) == "PT" ? "Portugal" : "España";
                            c.es_cisterna = false;
                            
                            if(c.tarifa != null)
                            {
                                int n;
                                if (int.TryParse(c.tarifa.Substring(0, 1), out n))
                                    c.grupo_presion = Convert.ToInt32(c.tarifa.Substring(0, 1));
                            }
                           

                        }
                        else
                        {
                            c.es_cisterna = true;
                            c.cupsree = "Cisterna_" + c.id_ps;
                            c.pais = "España";
                        }                           
                            
                    }
                    else
                    {
                        c.cupsree = "Cisterna_" + c.id_ps;
                        c.es_cisterna = true;
                        c.pais = "España";
                    }
                        

                    if (r["DE_DISTRIBUIDORES"] != System.DBNull.Value)
                        c.distribuidora = r["DE_DISTRIBUIDORES"].ToString().Trim();


                    if ((c.id_estado_contrato == 3 || c.id_estado_contrato == 6 || c.id_estado_contrato == 7 ||
                        c.id_estado_contrato == 8 || c.id_estado_contrato == 9 || c.id_estado_contrato == 10 ||
                        c.id_estado_contrato == 15 || c.id_estado_contrato == 16) && 
                        (c.fecha_inicio <= DateTime.Now && c.fecha_fin >= DateTime.Now))
                    {
                        EndesaEntity.sigame.ContratoGas o;
                        if (!dic.TryGetValue(c.cupsree, out o))
                        {
                            dic.Add(c.cupsree, c);

                            string oo;
                            if (!dic_nif_vigentes.TryGetValue(c.nif, out oo))
                                dic_nif_vigentes.Add(c.nif, c.nif);

                        }
                            
                    }

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(id);
                ficheroLog.AddError("Carga: " + e.Message);
            }
        }

        private void CargaCM(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            int id = 0;
            SqlDataAdapter da;

            try
            {
                Console.WriteLine("Cargando datos de SIGAME");
                #region Query
                strSql = "SELECT DISTINCT"
                    + " T_SGM_G_PS.ID_PS, T_SGM_G_PS.DE_PS, T_SGM_M_CLIENTES.CD_CIF,T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
                    + " T_SGM_G_CONTRATOS_PS.ID_CTO_PS, T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL,"
                    + " T_SGM_G_CONTRATOS_PS.FH_FIN_REAL, T_SGM_G_CONTRATOS_PS.FH_FIN_PREVISTA, T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO,"
                    + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA, T_SGM_G_PS.CD_CUPS, T_SGM_M_GESTORES.DE_GESTOR,"
                    + " DE_NOMBRE_MUNICIPIO, T_SGM_P_PROVINCIAS.DE_PROVINCIA,T_SGM_G_PS.CD_PAIS,"
                    + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES, T_SGM_M_RED_DISTRIBUCION.DE_RED_DISTRIBUCION,"
                    + " (T_SGM_G_PS.NM_SUMA_CONSUMOS_MEN) / 12 AS PROMEDIO, T_SGM_G_PS.NM_ENERGIA_DIARIA_KWH,"
                    + " T_SGM_G_CONTRATOS_PS.FH_ULT_ACTUALIZACION"
                    + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
                        + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                    + " LEFT JOIN T_SGM_M_CLIENTES ON"
                        + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " LEFT JOIN T_SGM_M_GESTORES ON"
                        + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR "
                    + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
                        + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
                    + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
                        + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES "
                    + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON "
                        + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
                    + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
                        + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
                        + " T_SGM_G_PS.CD_PAIS  = T_SGM_P_PROVINCIAS.CD_PAIS"
                    + " LEFT JOIN T_SGM_P_MUNICIPIOS ON "
                        + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
                    + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.ContratoGas c = new EndesaEntity.sigame.ContratoGas();
                    if (r["ID_PS"] != System.DBNull.Value)
                    {
                        c.id_ps = Convert.ToInt32(r["ID_PS"]);
                        id = c.id_ps;
                    }


                    if (r["ID_CTO_PS"] != System.DBNull.Value)
                        c.id_cto_ps = Convert.ToInt32(r["ID_CTO_PS"]);

                    if (r["DE_PS"] != System.DBNull.Value)
                        c.descripcion_ps = Convert.ToString(r["DE_PS"]);

                    if (r["DE_NOMBRE_CLIENTE"] != System.DBNull.Value)
                        c.nombre_cliente = Convert.ToString(r["DE_NOMBRE_CLIENTE"]);

                    //if (r["NM_ENERGIA_DIARIA_KWH"] != System.DBNull.Value)
                    //    c.qd = Convert.ToDouble(r["NM_ENERGIA_DIARIA_KWH"]);

                    c.nif = Convert.ToString(r["CD_CIF"]);

                    if (r["FH_INICIO_REAL"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["FH_INICIO_REAL"]);
                    else
                        c.fecha_inicio = DateTime.MinValue;

                    if (r["FH_FIN_REAL"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["FH_FIN_REAL"]);
                    else
                        c.fecha_fin = new DateTime(4999, 12, 31);

                    if (r["ID_ESTADO_CTO"] != System.DBNull.Value)
                        c.id_estado_contrato = Convert.ToInt32(r["ID_ESTADO_CTO"]);
                    else
                        c.id_estado_contrato = 0;

                    if (r["DE_TIPO_TARIFA"] != System.DBNull.Value)
                    {
                        c.tarifa = Convert.ToString(r["DE_TIPO_TARIFA"]);
                    }


                    if (r["CD_CUPS"] != System.DBNull.Value)
                    {

                        c.cupsree = Convert.ToString(r["CD_CUPS"]).Trim();
                        if (c.cupsree.Length > 2)
                        {
                            c.pais = c.cupsree.Substring(0, 2) == "PT" ? "Portugal" : "España";
                            c.es_cisterna = false;

                            if (c.tarifa != null)
                            {
                                int n;
                                if (int.TryParse(c.tarifa.Substring(0, 1), out n))
                                    c.grupo_presion = Convert.ToInt32(c.tarifa.Substring(0, 1));
                            }


                        }
                        else
                        {
                            c.es_cisterna = true;
                            c.cupsree = "Cisterna_" + c.id_ps;
                            c.pais = "España";
                        }

                    }
                    else
                    {
                        c.cupsree = "Cisterna_" + c.id_ps;
                        c.es_cisterna = true;
                        c.pais = "España";
                    }


                    if (r["DE_DISTRIBUIDORES"] != System.DBNull.Value)
                        c.distribuidora = r["DE_DISTRIBUIDORES"].ToString().Trim();


                    if ((c.id_estado_contrato == 3 || c.id_estado_contrato == 6 || c.id_estado_contrato == 7 ||
                         c.id_estado_contrato == 8 || c.id_estado_contrato == 9 || c.id_estado_contrato == 10 ||
                         c.id_estado_contrato == 15 || c.id_estado_contrato == 16) &&
                         (c.fecha_inicio <= DateTime.Now && c.fecha_fin >= DateTime.Now))
                    {
                        EndesaEntity.sigame.ContratoGas o;
                        if (!dic.TryGetValue(c.cupsree, out o))
                            dic.Add(c.cupsree, c);
                    }

                }
                db.CloseConnection();

                #region Query
                strSql = "SELECT DISTINCT"
                    + " T_SGM_G_PS.ID_PS, T_SGM_G_PS.DE_PS, T_SGM_M_CLIENTES.CD_CIF,T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
                    + " T_SGM_G_CONTRATOS_PS.ID_CTO_PS, T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL,"
                    + " T_SGM_G_CONTRATOS_PS.FH_FIN_REAL, T_SGM_G_CONTRATOS_PS.FH_FIN_PREVISTA, T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO,"
                    + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA, T_SGM_G_PS.CD_CUPS, T_SGM_M_GESTORES.DE_GESTOR,"
                    + " DE_NOMBRE_MUNICIPIO, T_SGM_P_PROVINCIAS.DE_PROVINCIA,T_SGM_G_PS.CD_PAIS,"
                    + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES, T_SGM_M_RED_DISTRIBUCION.DE_RED_DISTRIBUCION,"
                    + " (T_SGM_G_PS.NM_SUMA_CONSUMOS_MEN) / 12 AS PROMEDIO, T_SGM_G_PS.NM_ENERGIA_DIARIA_KWH,"
                    + " T_SGM_G_CONTRATOS_PS.FH_ULT_ACTUALIZACION"
                    + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
                        + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                    + " LEFT JOIN T_SGM_M_CLIENTES ON"
                        + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " LEFT JOIN T_SGM_M_GESTORES ON"
                        + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR "
                    + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
                        + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
                    + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
                        + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES "
                    + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON "
                        + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
                    + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
                        + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
                        + " T_SGM_G_PS.CD_PAIS  = T_SGM_P_PROVINCIAS.CD_PAIS"
                    + " LEFT JOIN T_SGM_P_MUNICIPIOS ON "
                        + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
                    + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.ContratoGas c = new EndesaEntity.sigame.ContratoGas();
                    if (r["ID_PS"] != System.DBNull.Value)
                    {
                        c.id_ps = Convert.ToInt32(r["ID_PS"]);
                        id = c.id_ps;
                    }


                    if (r["ID_CTO_PS"] != System.DBNull.Value)
                        c.id_cto_ps = Convert.ToInt32(r["ID_CTO_PS"]);

                    if (r["DE_PS"] != System.DBNull.Value)
                        c.descripcion_ps = Convert.ToString(r["DE_PS"]);

                    if (r["DE_NOMBRE_CLIENTE"] != System.DBNull.Value)
                        c.nombre_cliente = Convert.ToString(r["DE_NOMBRE_CLIENTE"]);

                    //if (r["NM_ENERGIA_DIARIA_KWH"] != System.DBNull.Value)
                    //    c.qd = Convert.ToDouble(r["NM_ENERGIA_DIARIA_KWH"]);

                    c.nif = Convert.ToString(r["CD_CIF"]);

                    if (r["FH_INICIO_REAL"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["FH_INICIO_REAL"]);
                    else
                        c.fecha_inicio = DateTime.MinValue;

                    if (r["FH_FIN_REAL"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["FH_FIN_REAL"]);
                    else
                        c.fecha_fin = new DateTime(4999, 12, 31);

                    if (r["ID_ESTADO_CTO"] != System.DBNull.Value)
                        c.id_estado_contrato = Convert.ToInt32(r["ID_ESTADO_CTO"]);
                    else
                        c.id_estado_contrato = 0;

                    if (r["DE_TIPO_TARIFA"] != System.DBNull.Value)
                    {
                        c.tarifa = Convert.ToString(r["DE_TIPO_TARIFA"]);
                    }


                    if (r["CD_CUPS"] != System.DBNull.Value)
                    {

                        c.cupsree = Convert.ToString(r["CD_CUPS"]).Trim();
                        if (c.cupsree.Length > 2)
                        {
                            c.pais = c.cupsree.Substring(0, 2) == "PT" ? "Portugal" : "España";
                            c.es_cisterna = false;

                            if (c.tarifa != null)
                            {
                                int n;
                                if (int.TryParse(c.tarifa.Substring(0, 1), out n))
                                    c.grupo_presion = Convert.ToInt32(c.tarifa.Substring(0, 1));
                            }


                        }
                        else
                        {
                            c.es_cisterna = true;
                            c.cupsree = "Cisterna_" + c.id_ps;
                            c.pais = "España";
                        }

                    }
                    else
                    {
                        c.cupsree = "Cisterna_" + c.id_ps;
                        c.es_cisterna = true;
                        c.pais = "España";
                    }


                    if (r["DE_DISTRIBUIDORES"] != System.DBNull.Value)
                        c.distribuidora = r["DE_DISTRIBUIDORES"].ToString().Trim();


                    if ((c.id_estado_contrato == 10) &&
                         (c.fecha_fin > DateTime.Now.AddMonths(-3)))
                    {
                        EndesaEntity.sigame.ContratoGas o;
                        if (!dic.TryGetValue(c.cupsree, out o))
                            dic.Add(c.cupsree, c);
                    }

                }
                db.CloseConnection();





            }
            catch (Exception e)
            {
                Console.WriteLine(id);
                ficheroLog.AddError("Carga: " + e.Message);
            }
        }

        private void Carga(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            int id = 0;

            try
            {
                Console.WriteLine("Cargando datos de SIGAME");
                #region Query
                strSql = "SELECT DISTINCT"
                    + " T_SGM_G_PS.ID_PS, T_SGM_G_PS.DE_PS, T_SGM_M_CLIENTES.CD_CIF,T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
                    + " T_SGM_G_CONTRATOS_PS.CD_COD_CTO_PAPEL, T_SGM_G_PS.TX_DIRECCION, T_SGM_G_PS.TX_NUMERO_FINCA,"
                    + " T_SGM_G_PS.CD_MUNICIPIO, T_SGM_G_PS.ID_PROVINCIA, T_SGM_G_PS.TX_CP,"
                    + " T_SGM_G_CONTRATOS_PS.ID_CTO_PS, T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL,"
                    + " T_SGM_G_CONTRATOS_PS.FH_FIN_REAL, T_SGM_G_CONTRATOS_PS.FH_FIN_PREVISTA, T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO,"
                    + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA, T_SGM_G_PS.CD_CUPS, T_SGM_M_GESTORES.DE_GESTOR,"
                    + " DE_NOMBRE_MUNICIPIO, T_SGM_P_PROVINCIAS.DE_PROVINCIA,T_SGM_G_PS.CD_PAIS,"
                    + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES, T_SGM_M_RED_DISTRIBUCION.DE_RED_DISTRIBUCION,"
                    + " (T_SGM_G_PS.NM_SUMA_CONSUMOS_MEN) / 12 AS PROMEDIO, T_SGM_G_PS.NM_ENERGIA_DIARIA_KWH,"
                    + " T_SGM_G_CONTRATOS_PS.FH_ULT_ACTUALIZACION"
                    + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
                        + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                    + " LEFT JOIN T_SGM_M_CLIENTES ON"
                        + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " LEFT JOIN T_SGM_M_GESTORES ON"
                        + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR "
                    + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
                        + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
                    + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
                        + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES "
                    + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON "
                        + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
                    + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
                        + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
                        + " T_SGM_G_PS.CD_PAIS  = T_SGM_P_PROVINCIAS.CD_PAIS"
                    + " LEFT JOIN T_SGM_P_MUNICIPIOS ON "
                        + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
                    + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS"
                    + " WHERE T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL <= '" + fh.ToString("yyyy-MM-dd") + "' AND"
                    + " (T_SGM_G_CONTRATOS_PS.FH_FIN_REAL >= '" + fd.ToString("yyyy-MM-dd") + "' OR"
                    + " T_SGM_G_CONTRATOS_PS.FH_FIN_REAL IS null)"
                    + " order by T_SGM_G_PS.ID_PS, T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL DESC";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.ContratoGas c = new EndesaEntity.sigame.ContratoGas();
                    if (r["ID_PS"] != System.DBNull.Value)
                    {
                        c.id_ps = Convert.ToInt32(r["ID_PS"]);
                        id = c.id_ps;
                    }

                    if (r["CD_COD_CTO_PAPEL"] != System.DBNull.Value)                    
                        c.contrato = r["CD_COD_CTO_PAPEL"].ToString();

                    if (r["TX_DIRECCION"] != System.DBNull.Value)
                        c.direccion = r["TX_DIRECCION"].ToString();

                    if (r["TX_NUMERO_FINCA"] != System.DBNull.Value)
                        c.numero_finca = r["TX_NUMERO_FINCA"].ToString();

                    if (r["CD_MUNICIPIO"] != System.DBNull.Value)
                        c.codigo_municipio = r["CD_MUNICIPIO"].ToString();

                    if (r["ID_PROVINCIA"] != System.DBNull.Value)
                        c.id_provincia = r["ID_PROVINCIA"].ToString();

                    if (r["TX_CP"] != System.DBNull.Value)
                        c.codigo_postal = r["TX_CP"].ToString();
                    

                    if (r["ID_CTO_PS"] != System.DBNull.Value)
                        c.id_cto_ps = Convert.ToInt32(r["ID_CTO_PS"]);

                    if (r["DE_PS"] != System.DBNull.Value)
                        c.descripcion_ps = Convert.ToString(r["DE_PS"]);

                    if (r["DE_NOMBRE_CLIENTE"] != System.DBNull.Value)
                        c.nombre_cliente = Convert.ToString(r["DE_NOMBRE_CLIENTE"]);

                    //if (r["NM_ENERGIA_DIARIA_KWH"] != System.DBNull.Value)
                    //    c.qd = Convert.ToDouble(r["NM_ENERGIA_DIARIA_KWH"]);

                    c.nif = Convert.ToString(r["CD_CIF"]);

                    if (r["FH_INICIO_REAL"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["FH_INICIO_REAL"]);
                    else
                        c.fecha_inicio = DateTime.MinValue;

                    if (r["FH_FIN_REAL"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["FH_FIN_REAL"]);
                    

                    if (r["ID_ESTADO_CTO"] != System.DBNull.Value)
                        c.id_estado_contrato = Convert.ToInt32(r["ID_ESTADO_CTO"]);
                    else
                        c.id_estado_contrato = 0;

                    if (r["DE_TIPO_TARIFA"] != System.DBNull.Value)
                    {
                        c.tarifa = Convert.ToString(r["DE_TIPO_TARIFA"]);
                    }


                    if (r["CD_CUPS"] != System.DBNull.Value)
                    {

                        c.cupsree = Convert.ToString(r["CD_CUPS"]).Trim();
                        if (c.cupsree.Length > 2)
                        {
                            c.pais = c.cupsree.Substring(0, 2) == "PT" ? "Portugal" : "España";
                            c.es_cisterna = false;

                            if (c.tarifa != null)
                            {
                                int n;
                                if (int.TryParse(c.tarifa.Substring(0, 1), out n))
                                    c.grupo_presion = Convert.ToInt32(c.tarifa.Substring(0, 1));
                            }


                        }
                        else
                        {
                            c.es_cisterna = true;
                            c.cupsree = "Cisterna_" + c.id_ps;
                            c.pais = "España";
                        }

                    }
                    else
                    {
                        c.cupsree = "Cisterna_" + c.id_ps;
                        c.es_cisterna = true;
                        c.pais = "España";
                    }


                    if (r["DE_DISTRIBUIDORES"] != System.DBNull.Value)
                        c.distribuidora = r["DE_DISTRIBUIDORES"].ToString().Trim();
                                        
                    EndesaEntity.sigame.ContratoGas o;
                    if (!dic.TryGetValue(c.cupsree, out o))
                        dic.Add(c.cupsree, c);

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(id);
                ficheroLog.AddError("Carga: " + e.Message);
            }
        }

        private void CargaPuntos()
        {

            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql;

            try
            {
                strSql = "select ID_PS, ID_PMEDIDA from T_SGM_M_PUNTOS_MEDIDA GROUP BY ID_PS, ID_PMEDIDA";
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    if (r["ID_PS"] != System.DBNull.Value && r["ID_PMEDIDA"] != System.DBNull.Value)
                    {
                        List<int> o;
                        if (!dic_pm.TryGetValue(Convert.ToInt32(r["ID_PS"]), out o))
                        {
                            o = new List<int>();
                            o.Add(Convert.ToInt32(r["ID_PMEDIDA"]));
                            dic_pm.Add(Convert.ToInt32(r["ID_PS"]), o);
                        }
                        else
                            o.Add(Convert.ToInt32(r["ID_PMEDIDA"]));
                    }

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public bool ExisteCUPS(string cups)
        {
            EndesaEntity.sigame.ContratoGas o;
            return dic.TryGetValue(cups, out o);
        }

        private void CargaConsumos(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql;

            try
            {
                strSql = "select LC.ID_PMEDIDA,"
                    + " SUM(LC.NM_VOL_CORR_PTZ) as CONSUMO_TM,"
                    + " SUM(LC.NM_TELEMEDIDA_PTZ) as CONSUMO_RUTA,"
                    + " SUM(LC.NM_CONSUMO) AS KWH_RUTA,"
                    + " SUM(LC.NM_CONSUMO_TELEMEDIDO) as KWH_TM,"
                    + " LC.TX_MES_LECTURA"
                    + " from T_SGM_M_LECTURAS_CONSUMOS AS LC"
                    + " WHERE"
                    + " (LC.TX_MES_LECTURA >= " + fd.ToString("yyyyMM") + " and"
                    + " LC.TX_MES_LECTURA <= " + fh.ToString("yyyyMM") + ")"
                    + " group by LC.ID_PMEDIDA, LC.TX_MES_LECTURA";

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.Consumos c = new EndesaEntity.sigame.Consumos();
                    c.id_pmedida = Convert.ToInt32(r["ID_PMEDIDA"]);
                    if (r["CONSUMO_TM"] != System.DBNull.Value)
                        c.consumo_tm = Math.Round(Convert.ToDouble(r["CONSUMO_TM"]), 0);
                    if (r["CONSUMO_RUTA"] != System.DBNull.Value)
                        c.consumo_ruta = Math.Round(Convert.ToDouble(r["CONSUMO_RUTA"]), 0);
                    if (r["KWH_RUTA"] != System.DBNull.Value)
                        c.kwh_ruta = Math.Round(Convert.ToDouble(r["KWH_RUTA"]), 0);
                    if (r["KWH_TM"] != System.DBNull.Value)
                        c.kwh_tm = Math.Round(Convert.ToDouble(r["KWH_TM"]), 0);
                    c.mes_consumo = Convert.ToInt32(r["TX_MES_LECTURA"]);

                    List<EndesaEntity.sigame.Consumos> o;
                    if (!dic_consumos.TryGetValue(c.id_pmedida, out o))
                    {
                        o = new List<EndesaEntity.sigame.Consumos>();
                        o.Add(c);
                        dic_consumos.Add(c.id_pmedida, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void GetConsumo(int id_ps, int mes_consumo)
        {
            EndesaEntity.sigame.Consumos c = new EndesaEntity.sigame.Consumos();
            List<int> o;
            List<EndesaEntity.sigame.Consumos> oo;
            if (dic_pm.TryGetValue(id_ps, out o))
            {
                this.consumos.consumo_tm = 0;
                this.consumos.consumo_ruta = 0;
                this.consumos.kwh_ruta = 0;
                this.consumos.kwh_tm = 0;

                for (int i = 0; i < o.Count; i++)
                {
                    if (dic_consumos.TryGetValue(o[i], out oo))
                    {
                        c = oo.Find(z => z.mes_consumo == mes_consumo);
                        if (c != null)
                        {
                            this.consumos.consumo_tm += c.consumo_tm;
                            this.consumos.consumo_ruta += c.consumo_ruta;
                            this.consumos.kwh_ruta += c.kwh_ruta;
                            this.consumos.kwh_tm += c.kwh_tm;
                        }
                    }

                }

            }
        }

        public bool EsVigente(string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return true;
            else
                return false;
        }

        public string Distribuidora(string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return o.distribuidora.Trim().ToUpper();
            else
                return null;
        }

        public DateTime FechaInicio(string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return o.fecha_inicio;
            else
                return DateTime.MinValue;
        }

        public string NIF(string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return o.nif;
            else
                return null;
        }

        public string NombreCliente(string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return o.nombre_cliente;
            else
                return null;
        }

        public int Grupo_Presion(string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return o.grupo_presion;
            else
                return 2; // Grupo por defecto por ser el caso normal
        }

        public string Tarifa (string cups20)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups20, out o))
                return o.tarifa;
            else
                return "";
        }

        public string Fuente(string cfactura)
        {
            EndesaEntity.sigame.Factura o;
            if (dic_facturas.TryGetValue(cfactura, out o))
                return o.fuente;
            else
                return "";
        }

        public string Medida(string cfactura)
        {
            EndesaEntity.sigame.Factura o;
            if (dic_facturas.TryGetValue(cfactura, out o))
                return o.medida;
            else
                return "";
        }

        public void GetContrato(string cups)
        {
            EndesaEntity.sigame.ContratoGas o;
            if (dic.TryGetValue(cups, out o))
            {
                this.fecha_inicio = o.fecha_inicio;
                this.fecha_fin = o.fecha_fin;
                this.tarifa = o.tarifa;
            }
        }

        public bool EsNIF_Vigente(string nif)
        {
            string o;
            return dic_nif_vigentes.TryGetValue(nif, out o);
        }
    }


}
