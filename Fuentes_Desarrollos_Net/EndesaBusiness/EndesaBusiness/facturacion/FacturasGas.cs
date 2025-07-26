using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class FacturasGas : EndesaEntity.facturacion.Factura
    {
        public Dictionary<int, List<EndesaEntity.facturacion.Factura>> dic_facturas_CM_GAS;

        Dictionary<int, EndesaEntity.facturacion.Factura> dic_fo_s;
        Dictionary<int, EndesaEntity.facturacion.Factura> dic_fo_s_sce;
        Dictionary<int, EndesaEntity.facturacion.SIGAME_TablaBase> dic_ultimaFactura;
        
        public FacturasGas()
        {
            
        }

        public void CargaFacturasGasCuadrodeMando(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            utilidades.Fechas utilFechas = new utilidades.Fechas();
            

            try
            {
                dic_facturas_CM_GAS = new Dictionary<int, List<EndesaEntity.facturacion.Factura>>();

                strSql = "Select ID_PS, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS, FH_FACTURA,"
                    + " ID_FACTURA, FH_ULT_ACTUALIZACION, NM_IMPORTE_BRUTO"
                    + " from fo_s f where"
                    + " FH_INI_FACTURACION >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FIN_FACTURACION <= '" + fh.ToString("yyyy-MM-dd") + "'"
                    // + " AND FH_FACTURA >= '" + utilFechas.UltimoDiaHabil().ToString("yyyy-MM-01") + "' and"
                    //+ " CD_NFACTURA_REALES_PS is not null and"
                    //+ " instr(CD_NFACTURA_REALES_PS,'S') < 1"
                    + " ORDER BY FH_INI_FACTURACION DESC";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();
                    c.id_pto_suministro = Convert.ToInt32(r["ID_PS"]);
                    c.fecha_expedicion_factura = Convert.ToDateTime(r["FH_ULT_ACTUALIZACION"]);
                    c.ffactdes = Convert.ToDateTime(r["FH_INI_FACTURACION"]);
                    c.ffacthas = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);

                    if (r["FH_FACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FH_FACTURA"]);

                    c.ifactura = Convert.ToDouble(r["NM_IMPORTE_BRUTO"]);

                    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        c.cfactura = r["CD_NFACTURA_REALES_PS"].ToString();

                    List<EndesaEntity.facturacion.Factura> o;
                    if (!dic_facturas_CM_GAS.TryGetValue(c.id_pto_suministro, out o))
                    {
                        o = new List<EndesaEntity.facturacion.Factura>();
                        o.Add(c);
                        dic_facturas_CM_GAS.Add(c.id_pto_suministro, o);
                    }
                    else
                        o.Add(c);
                    
                }
                db.CloseConnection();

                strSql = "Select ID_PS, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS,"                    
                    + " FH_FACTURA, FH_ULT_ACTUALIZACION, NM_IMPORTE_BRUTO"
                    + " from fo_s_sce f where "
                    + " FH_INI_FACTURACION >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FIN_FACTURACION <= '" + fh.ToString("yyyy-MM-dd") + "'"
                    // + " AND FH_FACTURA >= '" + utilFechas.UltimoDiaHabil().ToString("yyyy-MM-01") + "' and"
                    //+ " CD_NFACTURA_REALES_PS is not null and"
                    //+ " instr(CD_NFACTURA_REALES_PS,'S') < 1"
                    + " ORDER BY FH_INI_FACTURACION DESC";


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();

                    c.id_pto_suministro = Convert.ToInt32(r["ID_PS"]);
                    c.fecha_expedicion_factura = Convert.ToDateTime(r["FH_ULT_ACTUALIZACION"]);
                    c.ffactdes = Convert.ToDateTime(r["FH_INI_FACTURACION"]);
                    c.ffacthas = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);

                    if (r["FH_FACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FH_FACTURA"]);

                    c.ifactura = Convert.ToDouble(r["NM_IMPORTE_BRUTO"]);

                    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        c.cfactura = r["CD_NFACTURA_REALES_PS"].ToString();


                    List<EndesaEntity.facturacion.Factura> o;
                    if (!dic_facturas_CM_GAS.TryGetValue(c.id_pto_suministro, out o))
                    {
                        o = new List<EndesaEntity.facturacion.Factura>();
                        dic_facturas_CM_GAS.Add(c.id_pto_suministro, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();



                //strSql = "Select ID_PS, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS,"
                //    + " FH_FACTURA, FH_ULT_ACTUALIZACION, NM_IMPORTE_BRUTO"
                //    + " from fo_s_sce where "
                //    + " FH_INI_FACTURACION < '" + fd.ToString("yyyy-MM-dd") + "' and"
                //    + " FH_FACTURA > date_format(FH_FACTURA,'%Y-%m-01')"
                //    + " ORDER BY FH_ULT_ACTUALIZACION desc";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();
                //while (r.Read())
                //{
                //    EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();

                //    c.id_pto_suministro = Convert.ToInt32(r["ID_PS"]);
                //    c.fecha_expedicion_factura = Convert.ToDateTime(r["FH_ULT_ACTUALIZACION"]);
                //    c.ffactdes = Convert.ToDateTime(r["FH_INI_FACTURACION"]);
                //    c.ffacthas = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);
                //    c.ffactura = Convert.ToDateTime(r["FH_FACTURA"]);
                //    c.ifactura = Convert.ToDouble(r["NM_IMPORTE_BRUTO"]);
                //    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                //        c.cfactura = r["CD_NFACTURA_REALES_PS"].ToString();


                //    List<EndesaEntity.facturacion.Factura> o;
                //    if (!dic_facturas_CM_GAS.TryGetValue(c.id_pto_suministro, out o))
                //    {
                //        o = new List<EndesaEntity.facturacion.Factura>();
                //        dic_facturas_CM_GAS.Add(c.id_pto_suministro, o);
                //    }
                //    else
                //        o.Add(c);
                //}
                //db.CloseConnection();

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            //                return null;
            }
        }

        public void _Get_ID_PS_PrimeraFactura(int id_ps)
        {
            this.cfactura = null;
            this.facturado = false;
            this.existe = false;
            this.ultimo_mes_facturado = 0;
            this.fecha_expedicion_factura = DateTime.MinValue;
            this.ffactura = DateTime.MinValue;

            List<EndesaEntity.facturacion.Factura> o;
            if (dic_facturas_CM_GAS.TryGetValue(id_ps, out o))
            {
                if(o.Count > 0)
                {
                    this.existe = true;
                    this.ultimo_mes_facturado = Convert.ToInt32(o[0].ffactdes.ToString("yyyyMM"));
                    this.fecha_expedicion_factura = o[0].fecha_expedicion_factura;
                    this.ifactura = o[0].ifactura;
                    if (o[0].cfactura != null &&
                        (o[0].cfactura.Contains("N") || o[0].cfactura.Contains("Y")))
                    {
                        this.cfactura = o[0].cfactura;
                        this.ffactura = o[0].ffactura;
                        this.facturado = true;
                    }
                }
            } else
                this.existe = false;
        }

        public void Get_ID_PS_PrimeraFactura(int id_ps)
        {
            this.cfactura = null;
            this.facturado = false;
            this.existe = false;
            this.ultimo_mes_facturado = 0;
            this.fecha_expedicion_factura = DateTime.MinValue;
            this.ffactura = DateTime.MinValue;
            bool firstOnly = true;

            List<EndesaEntity.facturacion.Factura> o;
            if (dic_facturas_CM_GAS.TryGetValue(id_ps, out o))
            {
                this.existe = true;
                foreach (EndesaEntity.facturacion.Factura p in o)
                {
                    if (firstOnly)
                    {
                        this.ultimo_mes_facturado = Convert.ToInt32(p.ffactdes.ToString("yyyyMM"));
                        this.fecha_expedicion_factura = p.fecha_expedicion_factura;
                        this.ifactura = p.ifactura;
                        if (p.cfactura != null &&
                        (p.cfactura.Contains("N") ||p.cfactura.Contains("Y")))
                        {
                            this.cfactura = p.cfactura;
                            this.ffactura = p.ffactura;
                            this.facturado = true;
                        }
                        firstOnly = false;
                    }
                    else
                    {
                        if(Convert.ToInt32(p.ffactdes.ToString("yyyyMM")) > this.ultimo_mes_facturado)
                        {
                            this.ultimo_mes_facturado = Convert.ToInt32(p.ffactdes.ToString("yyyyMM"));
                            this.fecha_expedicion_factura = p.fecha_expedicion_factura;
                            this.ifactura = p.ifactura;
                            if (p.cfactura != null &&
                            (p.cfactura.Contains("N") || p.cfactura.Contains("Y")))
                            {
                                this.cfactura = p.cfactura;
                                this.ffactura = p.ffactura;
                                this.facturado = true;
                            }
                        }
                    }
                    
                }              
            }
            else
                this.existe = false;
        }

        public FacturasGas(DateTime fd, DateTime fh)
        {
            dic_fo_s = CargaFacturas("fo_s", fd, fh);
            dic_fo_s_sce = CargaFacturas("fo_s_sce", fd, fh);
        }

        public void LanzaCargaUltimaFactura()
        {
            dic_ultimaFactura = CargaUltimaFactura();
        }


        private Dictionary<int, EndesaEntity.facturacion.SIGAME_TablaBase> CargaUltimaFactura()
        {

            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";

            Dictionary<int, EndesaEntity.facturacion.SIGAME_TablaBase> d =
                new Dictionary<int, EndesaEntity.facturacion.SIGAME_TablaBase>();

            try
            {
                
                strSql = "SELECT"
                    + " T_SGM_G_PS.ID_PS,"
                    + " month(A.FH_FACTURA) as PERIODO,"
                    + " A.FH_INI_FACTURACION,"
                    + " A.FH_FIN_FACTURACION,"
                    + " Int_SCE_R_10.PeriodoFacturacion as PeriodoFacturacion,"
                    + " T_SGM_G_PS.DE_PS AS RAZON_SOCIAL,"
                    + " T_SGM_M_CLIENTES.CD_CIF as CIFNIF,"
                    + " T_SGM_G_PS.CD_CUPS as CUPS22,"
                    + " A.FH_FACTURA as FH_EMISION,"
                    + " A.CD_CREFAEXT,"
                    + " A.CD_NFACTURA_REALES_PS as C_FACTURA,"
                    + " A.CD_TESTFACT as CODIFICACION_FACTURA,"
                    + " A.TX_TIPO_FACTURA_NUEVO as TIPOLOGIA_FACTURA,"
                    + " RTRIM(Int_SCE_R_10.DireccionPS) as DireccionPuntoSuministro,"
                    + " A.NM_IEH_IMPORTE as IMPORTE_IH,"
                    + " A.NM_BASE_IMPONIBLE as BASE_IMPONIBLE_IVA,"
                    + " A.NM_IMPORTE_IVA,"
                    + " A.NM_IMPORTE_BRUTO As ImporteTotalFactura"
                    + " FROM T_SGM_M_FACTURAS_REALES_PS as A INNER JOIN"
                    + " (select B.ID_CTO_PS, max(B.FH_INI_FACTURACION) as maxFecha"
                    + " from T_SGM_M_FACTURAS_REALES_PS as B WHERE"
                    + " B.FH_FACTURA > B.FH_FIN_FACTURACION"
                    + " GROUP BY B.ID_CTO_PS) as B ON"
                    + " B.ID_CTO_PS = A.ID_CTO_PS AND"
                    + " B.maxFecha = A.FH_INI_FACTURACION"
                    + " INNER JOIN Int_SCE_R_10 ON A.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura"
                    + " INNER JOIN T_SGM_G_CONTRATOS_PS ON A.ID_CTO_PS = T_SGM_G_CONTRATOS_PS.ID_CTO_PS"
                    + " INNER JOIN T_SGM_G_PS ON T_SGM_G_CONTRATOS_PS.ID_PS = T_SGM_G_PS.ID_PS"
                    + " INNER JOIN T_SGM_M_CLIENTES ON T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " Group by"
                    + " T_SGM_G_PS.ID_PS,"
                    + " month(FH_FACTURA),"
                    + " A.FH_INI_FACTURACION,"
                    + " A.FH_FIN_FACTURACION,"
                    + " Int_SCE_R_10.PeriodoFacturacion,"
                    + " T_SGM_G_PS.DE_PS , T_SGM_M_CLIENTES.CD_CIF, T_SGM_G_PS.CD_CUPS, FH_FACTURA,"
                    + " A.CD_CREFAEXT,"
                    + " A.CD_NFACTURA_REALES_PS,"
                    + " CD_TESTFACT,TX_TIPO_FACTURA_NUEVO, RTRIM(Int_SCE_R_10.DireccionPS),"
                    + " NM_IEH_IMPORTE, NM_BASE_IMPONIBLE, NM_IMPORTE_IVA,NM_IMPORTE_BRUTO"
                    + " order by T_SGM_G_PS.CD_CUPS asc, FH_INI_FACTURACION desc";

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.SIGAME_TablaBase c = new EndesaEntity.facturacion.SIGAME_TablaBase();
                    if (r["ID_PS"] != System.DBNull.Value)
                        c.id_ps = Convert.ToInt32(r["ID_PS"]);
                                        
                    if (r["PERIODO"] != System.DBNull.Value)
                        c.periodo = Convert.ToInt32(r["PERIODO"]);
                    //if (r["SECFACTU"] != System.DBNull.Value)
                    //    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    if (r["FH_INI_FACTURACION"] != System.DBNull.Value)
                        c.fh_ini_facturacion = Convert.ToDateTime(r["FH_INI_FACTURACION"]);

                    if (r["FH_FIN_FACTURACION"] != System.DBNull.Value)
                        c.fh_fin_facturacion = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);

                    if (r["PeriodoFacturacion"] != System.DBNull.Value)
                        c.periodofacturacion = r["PeriodoFacturacion"].ToString();

                    if (r["RAZON_SOCIAL"] != System.DBNull.Value)
                        c.razon_social = r["RAZON_SOCIAL"].ToString();

                    if (r["CIFNIF"] != System.DBNull.Value)
                        c.cifnif = r["CIFNIF"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["FH_EMISION"] != System.DBNull.Value)
                        c.fh_emision = Convert.ToDateTime(r["FH_EMISION"]);

                    if (r["CD_CREFAEXT"] != System.DBNull.Value)
                        c.numrefFactura = r["CD_CREFAEXT"].ToString();

                    if (r["C_FACTURA"] != System.DBNull.Value)
                        c.c_factura = r["C_FACTURA"].ToString().Trim();

                    if (r["CODIFICACION_FACTURA"] != System.DBNull.Value)
                        c.codificacion_factura = r["CODIFICACION_FACTURA"].ToString();

                    if (r["TIPOLOGIA_FACTURA"] != System.DBNull.Value)
                        c.tipologia_factura = r["TIPOLOGIA_FACTURA"].ToString();

                    if (r["DireccionPuntoSuministro"] != System.DBNull.Value)
                        c.direccion_punto_suministro = r["DireccionPuntoSuministro"].ToString();

                    EndesaEntity.facturacion.SIGAME_TablaBase o;
                    if(!d.TryGetValue(c.id_ps, out o))
                        d.Add(c.id_ps, c);
                }
                db.CloseConnection();

                return d;
            }
            catch(Exception e)
            {
                return null;
            }
        }
        
        public int UltimaFactura(int id_ps)
        {
            EndesaEntity.facturacion.SIGAME_TablaBase o;
            if (dic_ultimaFactura.TryGetValue(id_ps, out o))
                return Convert.ToInt32(o.fh_ini_facturacion.AddMonths(1).ToString("yyyyMM"));
            else
                return 0;
        }


        public void InformeIH(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            List<EndesaEntity.facturacion.SIGAME_Informe_IH> l = 
                new List<EndesaEntity.facturacion.SIGAME_Informe_IH>();
            List<EndesaEntity.facturacion.SIGAME_TablaBase> subLista =
                new List<EndesaEntity.facturacion.SIGAME_TablaBase>();


            try
            {
                //List<EndesaEntity.facturacion.SIGAME_TablaBase> lista = CargaFacturasGas(fd, fh);

                #region Carga Plantilla Base
                strSql = "select cliente, nif, cups" +
                    " from tmp_angel_20200407_gas_inventario";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.SIGAME_Informe_IH c = 
                        new EndesaEntity.facturacion.SIGAME_Informe_IH();
                    if (r["cliente"] != System.DBNull.Value)
                        c.cliente = r["cliente"].ToString();
                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();
                    if (r["cups"] != System.DBNull.Value)
                        c.cups = r["cups"].ToString().Trim();

                    //subLista = lista.FindAll(z => z.cifnif == c.nif && z.cups22 == c.cups);
                    for(int i = 0; i < subLista.Count; i++)
                    {
                        if (!c.ihtc)
                            c.ihtc = (subLista[i].IHTC_importe != 0);
                        if (!c.ihtd)
                            c.ihtd = (subLista[i].IHTD_importe != 0);
                        if (!c.ihti)
                            c.ihti = (subLista[i].IHTI_importe != 0);

                        if (c.direccion_punto_suministro == null)
                            c.direccion_punto_suministro = subLista[i].direccion_punto_suministro;
                    }


                }
                db.CloseConnection();
                #endregion
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


        private Dictionary<string, EndesaEntity.facturacion.Fo_Tabla> CargaFacturasGO(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int total_registros = 0;
            int i = 0;
            Dictionary<string, EndesaEntity.facturacion.Fo_Tabla> d = 
                new Dictionary<string, EndesaEntity.facturacion.Fo_Tabla>();
            try
            {

                //Console.WriteLine("Cargando facturas de FO ...");
                //Console.WriteLine("");

                strSql = "SELECT count(*) as total_registros"                    
                    + " FROM fo f inner join fo_empresas e on" 
                    + " e.empresa_id = f.ID_ENTORNO"
                    + " WHERE e.descripcion in ('MT-España')"
                    + " and (f.ffactura >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.ffactura <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " and f.CFACTURA <> ''"
                    + " and f.TFACTURA = 4";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                    total_registros = Convert.ToInt32(r["total_registros"]);
                db.CloseConnection();

                strSql = "SELECT"
                    + " f.CFACTURA,  f.SECFACTU, f.IFACTURA, f.VCUOVAFA, f.TESTFACT,"
                    + " f.IVA, (f.IFACTURA - f.IVA) as BASE_IMPONIBLE"
                    + " FROM fo f inner join fo_empresas e on"
                    + " e.empresa_id = f.ID_ENTORNO"
                    + " WHERE e.descripcion in ('MT-España')"
                    + " and (f.ffactura >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.ffactura <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " and f.CFACTURA <> ''"
                    + " and f.TFACTURA = 4";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    i++;
                    //Console.CursorLeft = 0;
                    //Console.Write(String.Format("{0:#,##0}",i) + " de " + String.Format("{0:#,##0}", total_registros));
                    EndesaEntity.facturacion.Fo_Tabla c = new EndesaEntity.facturacion.Fo_Tabla();
                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();
                    if (r["SECFACTU"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);
                    if (r["VCUOVAFA"] != System.DBNull.Value)
                        c.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);
                    if (r["IVA"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["IVA"]);
                    if (r["BASE_IMPONIBLE"] != System.DBNull.Value)
                        c.base_imposible = Convert.ToDouble(r["BASE_IMPONIBLE"]);
                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();
                    d.Add(c.cfactura, c);

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        //public void GeneraInformeImpuestos(DateTime fd, DateTime fh)
        //{
        //    List<EndesaEntity.facturacion.SIGAME_TablaBase> lista  = CargaFacturasGas(fd, fh);
        //    Console.WriteLine("Generando Excel");
        //    //InformeExcel(@"c:\Temp\FAC_SIG_" + fd.Year + "_" + DateTime.Now.ToString("HHmmss") + ".xlsx", lista);
        //}

        public void InformeFiscalGas(DateTime fd, DateTime fh, string nombre_archivo)
        {

            int num_Impuestos = 0;

            MySQLDB dbm;
            MySqlCommand commandm;
            MySqlDataReader rm;
            string strSql = "";

            SQLServer db;
            SqlCommand command;
            SqlDataReader r;

            List<EndesaEntity.facturacion.SIGAME_TablaBase> lista = new List<EndesaEntity.facturacion.SIGAME_TablaBase>();
            Dictionary<string,List<EndesaEntity.facturacion.SIGAME_R_20>> dic_R20;
            // Dictionary<string, EndesaEntity.facturacion.SIGAME_R_30> dic_R30; // --> Para consumos totales
            Dictionary<string, EndesaEntity.facturacion.Fo_Tabla> dic_fo; // Para almacenar facturas SCE
            //Dictionary<string, string> dic_dir_suministro;

            // Dic para la 2º hoja del informe sobre que suministros tienen que impuestos de IH
            Dictionary<int, EndesaEntity.facturacion.SIGAME_Informe_IH> dic_ih =
                new Dictionary<int, EndesaEntity.facturacion.SIGAME_Informe_IH>();

            Dictionary<int, int> dic_exentos;

            EndesaBusiness.sigame.UsoGas uso = new sigame.UsoGas(fd, fh);

            try
            {
                dic_R20 =  CargaConceptosR20(fd, fh);
                // dic_R30 = CargaConceptosR30(fd, fh);
                // dic_dir_suministro = CargaDirPS();
                dic_fo = CargaFacturasGO(fd, fh);
                dic_exentos = CargaExentos(fd, fh);



                //Console.WriteLine("");
                //Console.WriteLine("Consultando facturas SIGAME");
                //Console.WriteLine("");

                #region Query
                strSql = "SELECT T_SGM_G_PS.ID_PS,"
                    + " month(FH_FACTURA) as PERIODO,"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_INI_FACTURACION,"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_FIN_FACTURACION,"
                    + " Int_SCE_R_10.PeriodoFacturacion as PeriodoFacturacion,"
                    + " T_SGM_G_PS.DE_PS AS RAZON_SOCIAL,"                    
                    + " T_SGM_M_CLIENTES.CD_CIF as CIFNIF,"
                    + " T_SGM_G_PS.CD_CUPS as CUPS22,"
                    + " T_SGM_P_PROVINCIAS.DE_PROVINCIA,"
                    + " FH_FACTURA as FH_EMISION,"
                    + " T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT,"
                    + " CD_NFACTURA_REALES_PS as C_FACTURA,"
                    + " CD_TESTFACT as CODIFICACION_FACTURA,"
                    + " TX_TIPO_FACTURA_NUEVO as TIPOLOGIA_FACTURA,"
                    + " RTRIM(Int_SCE_R_10.DireccionPS) as DireccionPuntoSuministro,"
                    // + " --RTRIM(Int_SCE_R_20.Descripcion) as Descripcion,"
                    // + " sum(CONVERT(Decimal,case when right(Int_SCE_R_30.ConsumoTotal, 1) = '1' then"
                    // + " Left(Int_SCE_R_30.ConsumoTotal, Len(Int_SCE_R_30.ConsumoTotal) - 1) else"
                    // + " Left(Int_SCE_R_30.ConsumoTotal, Len(Int_SCE_R_30.ConsumoTotal) - 1) * -1 end)) AS ConsumoTotal,"
                    // + " --Int_SCE_R_20.CodConcepto,
                    // + " --replace(replace(Int_SCE_R_25.Concepto, 'Impto. hidrocarburos ', ''), ' Eur/GJulio', '') as TIPO_IMPOSITIVO_IH,
                    + " T_SGM_M_FACTURAS_REALES_PS.NM_CONSUMO, "
                    + " NM_IEH_IMPORTE as IMPORTE_IH,"
                    + " NM_BASE_IMPONIBLE as BASE_IMPONIBLE_IVA,"
                    + " NM_IMPORTE_IVA,"
                    + " NM_IMPORTE_BRUTO As ImporteTotalFactura"
                    + " FROM T_SGM_M_FACTURAS_REALES_PS"
                    + " INNER JOIN T_SGM_G_CONTRATOS_PS ON T_SGM_M_FACTURAS_REALES_PS.ID_CTO_PS = T_SGM_G_CONTRATOS_PS.ID_CTO_PS"
                    + " INNER JOIN T_SGM_G_PS ON T_SGM_G_CONTRATOS_PS.ID_PS = T_SGM_G_PS.ID_PS"
                    + " INNER JOIN T_SGM_M_CLIENTES ON T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " INNER JOIN Int_SCE_R_10 ON T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura"
                    + " INNER JOIN T_SGM_P_PROVINCIAS ON T_SGM_P_PROVINCIAS.ID_PROVINCIA = T_SGM_G_PS.ID_PROVINCIA AND"
                    + " T_SGM_P_PROVINCIAS.CD_PAIS = 'ESP'"
                    //+ " INNER JOIN Int_SCE_R_20 ON Int_SCE_R_10.NumRefFactura = Int_SCE_R_20.NumRefFactura)"
                    // + " INNER JOIN Int_SCE_R_30 ON(Int_SCE_R_20.NumRefFactura = Int_SCE_R_30.NumRefFactura)"
                    //+ " AND(Int_SCE_R_20.FxGenInterfaz = Int_SCE_R_30.FxGenInterfaz))"
                    // + " LEFT OUTER JOIN Int_SCE_R_25 ON Int_SCE_R_20.CodConcepto = Int_SCE_R_25.CodConcepto"
                    + " WHERE SUBSTRING(T_SGM_M_CLIENTES.CD_CIF,1,2) <> 'PT' AND "
                    + " (FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    // + " and CD_NFACTURA_REALES_PS = '00Z604N0000537'
                    // + " and Int_SCE_R_20.CodConcepto in ('IHTC', 'IHTD', 'IHTI')"
                    + " Group by T_SGM_G_PS.ID_PS,"
                    + " month(FH_FACTURA),"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_INI_FACTURACION,"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_FIN_FACTURACION," 
                    + " Int_SCE_R_10.PeriodoFacturacion," 
                    + " T_SGM_G_PS.DE_PS, T_SGM_M_CLIENTES.CD_CIF, T_SGM_G_PS.CD_CUPS, "
                    + " T_SGM_P_PROVINCIAS.DE_PROVINCIA, FH_FACTURA,"
                    + " T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT,"
                    + " CD_NFACTURA_REALES_PS," 
                    + " CD_TESTFACT,TX_TIPO_FACTURA_NUEVO, RTRIM(Int_SCE_R_10.DireccionPS),"
                    + " T_SGM_M_FACTURAS_REALES_PS.NM_CONSUMO,"
                    // + " --RTRIM(Int_SCE_R_20.Descripcion), 
                    // + " --Int_SCE_R_20.CodConcepto,
                    // + " replace(replace(Int_SCE_R_25.Concepto, 'Impto. hidrocarburos ', ''), ' Eur/GJulio', '') , 
                    + " NM_IEH_IMPORTE, NM_BASE_IMPONIBLE, NM_IMPORTE_IVA,NM_IMPORTE_BRUTO"
                    + " order by T_SGM_M_CLIENTES.CD_CIF, T_SGM_G_PS.ID_PS, FH_FACTURA, CD_NFACTURA_REALES_PS";

                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();
                while (r.Read())
                {
                 
                    num_Impuestos = 0;

                    EndesaEntity.facturacion.SIGAME_TablaBase c = new EndesaEntity.facturacion.SIGAME_TablaBase();

                    if (r["ID_PS"] != System.DBNull.Value)
                        c.id_ps = Convert.ToInt32(r["ID_PS"]);

                    c.origen = "SIGAME";
                    if (r["PERIODO"] != System.DBNull.Value)
                        c.periodo = Convert.ToInt32(r["PERIODO"]);
                    //if (r["SECFACTU"] != System.DBNull.Value)
                    //    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    if (r["FH_INI_FACTURACION"] != System.DBNull.Value)
                        c.fh_ini_facturacion = Convert.ToDateTime(r["FH_INI_FACTURACION"]);

                    if (r["FH_FIN_FACTURACION"] != System.DBNull.Value)
                        c.fh_fin_facturacion = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);

                    if (r["PeriodoFacturacion"] != System.DBNull.Value)
                        c.periodofacturacion = r["PeriodoFacturacion"].ToString();

                    c.uso_gas = uso.Get_Uso_Gas(c.id_ps, c.fh_ini_facturacion, c.fh_fin_facturacion);

                    if (r["RAZON_SOCIAL"] != System.DBNull.Value)
                        c.razon_social = r["RAZON_SOCIAL"].ToString();

                    if (r["CIFNIF"] != System.DBNull.Value)
                        c.cifnif = r["CIFNIF"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["FH_EMISION"] != System.DBNull.Value)
                        c.fh_emision = Convert.ToDateTime(r["FH_EMISION"]);

                    if (r["CD_CREFAEXT"] != System.DBNull.Value)                    
                        c.numrefFactura = r["CD_CREFAEXT"].ToString();

                    if (r["C_FACTURA"] != System.DBNull.Value)
                        c.c_factura = r["C_FACTURA"].ToString().Trim();

                    if (r["CODIFICACION_FACTURA"] != System.DBNull.Value)
                        c.codificacion_factura = r["CODIFICACION_FACTURA"].ToString();

                    if (r["TIPOLOGIA_FACTURA"] != System.DBNull.Value)
                        c.tipologia_factura = r["TIPOLOGIA_FACTURA"].ToString();

                    if (r["DireccionPuntoSuministro"] != System.DBNull.Value)
                        c.direccion_punto_suministro = r["DireccionPuntoSuministro"].ToString();

                    

                    //if (r["NM_CONSUMO"] != System.DBNull.Value)
                    //    c.consumo_total = Convert.ToInt32(r["NM_CONSUMO"]);

                    if (r["IMPORTE_IH"] != System.DBNull.Value)
                        c.importe_ih = Convert.ToDouble(r["IMPORTE_IH"]);

                    if (r["BASE_IMPONIBLE_IVA"] != System.DBNull.Value)
                        c.base_imponible_iva = Convert.ToDouble(r["BASE_IMPONIBLE_IVA"]);

                    if (r["NM_IMPORTE_IVA"] != System.DBNull.Value)
                        c.nm_importe_iva = Convert.ToDouble(r["NM_IMPORTE_IVA"]);

                    if (r["ImporteTotalFactura"] != System.DBNull.Value)
                        c.ImporteTotalFactura = Convert.ToDouble(r["ImporteTotalFactura"]);

                    // Buscamos los datos de la factura en el SCE
                    EndesaEntity.facturacion.Fo_Tabla fo;
                    if(dic_fo.TryGetValue(c.c_factura, out fo))
                    {
                        c.ImporteTotalFactura_sce = fo.ifactura;
                        c.secfactu = fo.secfactu;
                        c.codificacion_factura = fo.testfact;

                        if(c.base_imponible_iva + c.nm_importe_iva != c.ImporteTotalFactura)
                        {
                            c.base_imponible_iva = fo.base_imposible;
                            c.nm_importe_iva = fo.iva;
                        }
                    }

                    if (r["NM_CONSUMO"] != System.DBNull.Value)
                        c.consumo_total = Math.Abs(Convert.ToInt32(r["NM_CONSUMO"]))
                            * (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1);


                    //if (r["IFACTURA"] != System.DBNull.Value)
                    //    c.ImporteTotalFactura_sce = Convert.ToDouble(r["IFACTURA"]);

                    c.marca = "SIN IH";

                    if (r["CD_CREFAEXT"] != System.DBNull.Value)
                    {
                        c.numrefFactura = r["CD_CREFAEXT"].ToString();
                        //string dir_ps;
                        //if (dic_dir_suministro.TryGetValue(c.numrefFactura, out dir_ps))
                        //    c.direccion_punto_suministro = dir_ps;

                        List<EndesaEntity.facturacion.SIGAME_R_20> l_r20;
                        //if (dic_R20.TryGetValue(c.numrefFactura, out l_r20))
                        if (dic_R20.TryGetValue(c.c_factura, out l_r20))
                        {

                            num_Impuestos = 0;                       
                            for(int j = 0; j < l_r20.Count; j++)
                            {                                
                                switch (l_r20[j].tipo_impositivo)
                                {
                                    case 0.15:
                                    num_Impuestos++;
                                    c.IHTI_tipo_impositivo = 0.15;
                                    c.IHTI_descripcion = l_r20[j].descripcion;
                                    //c.IHTI_consumo = Math.Abs(l_r20[j].consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1);
                                    c.IHTI_importe = Math.Abs(l_r20[j].importe) *
                                        (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1);
                                    break;
                                    case 0.65:
                                    num_Impuestos++;
                                    c.IHTD_tipo_impositivo = 0.65;
                                    c.IHTD_descripcion = l_r20[j].descripcion;
                                    //c.IHTD_consumo = Math.Abs(l_r20[j].consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1);
                                    c.IHTD_importe = Math.Abs(l_r20[j].importe) *
                                        (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1);
                                        break;
                                    case 1.15:
                                    num_Impuestos++;
                                    c.IHTC_tipo_impositivo = 1.15;
                                        c.IHTC_descripcion = l_r20[j].descripcion;
                                        //c.IHTC_consumo = Math.Abs(l_r20[j].consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1);
                                        c.IHTC_importe = Math.Abs(l_r20[j].importe) * 
                                            (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? - 1: 1);
                                        break;
                                }

                           }
                        }
                    }
                    //EndesaEntity.facturacion.SIGAME_R_30 r30;
                    //if(dic_R30.TryGetValue(c.numrefFactura, out r30))
                    //{
                    //    c.consumo_total = Math.Abs(r30.consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1);
                    //    //c.consumo_total = r30.consumo;
                    //}

                    switch (num_Impuestos)
                    {
                        case 0:
                            c.marca = "SIN IH";
                            break;
                        case 1:
                            c.marca = "Factura Simple IH";
                            break;
                        case 2:
                            c.marca = "Factura Doble IH";
                            break;
                    }
                        

                    lista.Add(c);

                    #region Datos para el informe GAS
                    EndesaEntity.facturacion.SIGAME_Informe_IH o;
                    if(!dic_ih.TryGetValue(c.id_ps, out o))
                    {
                        o = new EndesaEntity.facturacion.SIGAME_Informe_IH();
                        o.cups = c.cups22;
                        o.id_ps = c.id_ps;
                        o.nif = c.cifnif;
                        o.cliente = c.razon_social;
                        o.direccion_punto_suministro = c.direccion_punto_suministro;
                        if (r["DE_PROVINCIA"] != System.DBNull.Value)
                            o.provincia = r["DE_PROVINCIA"].ToString();

                        o.ihtc = c.IHTC_importe != 0;
                        o.ihtd = c.IHTD_importe != 0;
                        o.ihti = c.IHTI_importe != 0;

                        int x;
                        if (dic_exentos.TryGetValue(c.id_ps, out x))
                            o.exento = true;
                            
                        dic_ih.Add(o.id_ps, o);
                    }
                    else
                    {
                        o.ihtc = c.IHTC_importe != 0;
                        o.ihtd = c.IHTD_importe != 0;
                        o.ihti = c.IHTI_importe != 0;
                        int x;
                        if (dic_exentos.TryGetValue(c.id_ps, out x))
                            o.exento = true;

                    }
                    #endregion

                }
                db.CloseConnection();

                Console.WriteLine("Consultando facturas SCE");

                strSql = " SELECT 'SCE' AS origen," 
                    + " MONTH(s.FH_FACTURA) AS periodo,"
                    + " s.ID_PS, "
                    + " s.FH_INI_FACTURACION AS fh_ini_facturacion," 
                    + " s.FH_FIN_FACTURACION AS fh_fin_facturacion,"
                    + " f.SECFACTU,"
                    + " NULL AS periodofacturacion," 
                    + " f.DAPERSOC as RAZON_SOCIAL," 
                    + " f.CNIFDNIC AS CIFNIF,"
                    + " f.CUPSREE AS cups22,"
                    + " i.Provincia,"
                    + " s.FH_FACTURA AS fh_emision, " 
                    + " s.CD_NFACTURA_REALES_PS AS c_factura,"
                    + " f.TESTFACT AS codificacion_factura," 
                    + " NULL AS tipologia_factura,"
                    + " null AS direccion_punto_suministro, "
                    + " sum(t.ICONFAC) as IMPORTE_IH," // 
                    + " s.NM_CONSUMO AS consumo_total,"                     
                    + " s.NM_BASE_IMPONIBLE as BASE_IMPONIBLE_IVA, s.NM_IMPORTE_IVA,"
                    + " f.IVA,"
                    + " s.NM_IMPORTE_BRUTO AS ImporteTotalFactura,"
                    + " f.IFACTURA"
                    + " FROM fo_s_sce s INNER JOIN"
                    + " fo f ON"
                    + " f.CFACTURA = s.CD_NFACTURA_REALES_PS"
                    + " LEFT OUTER JOIN cm_inventario_gas i on"    
                    + " s.ID_PS = i.ID_PS"
                    + " LEFT OUTER JOIN fo_tcon t on"
                    + " t.CREFEREN = f.CREFEREN and"
                    + " t.SECFACTU = f.SECFACTU and"
                    + " t.TESTFACT = f.TESTFACT and"
                    + " t.TCONFAC in (1100, 1101, 1102, 1103, 1104, 1105, 1106, 1107, 1504, 1505, 1512, 1633, 1634)"
                    + " WHERE(s.FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                    + " s.FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " AND substr(f.CNIFDNIC,1,2) <> 'PT'"
                    + " group by s.CD_NFACTURA_REALES_PS";
                dbm = new MySQLDB(MySQLDB.Esquemas.FAC);
                commandm = new MySqlCommand(strSql, dbm.con);
                rm = commandm.ExecuteReader();
                while (rm.Read())
                {
                    EndesaEntity.facturacion.SIGAME_TablaBase c = new EndesaEntity.facturacion.SIGAME_TablaBase();

                    c.origen = "SCE";
                    if (rm["ID_PS"] != System.DBNull.Value)                    
                        c.id_ps = Convert.ToInt32(rm["ID_PS"]);                    
                        
                    if (rm["PERIODO"] != System.DBNull.Value)
                        c.periodo = Convert.ToInt32(rm["PERIODO"]);
                    if (rm["SECFACTU"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(rm["SECFACTU"]);
                    if (rm["FH_INI_FACTURACION"] != System.DBNull.Value)
                        c.fh_ini_facturacion = Convert.ToDateTime(rm["FH_INI_FACTURACION"]);
                    if (rm["FH_FIN_FACTURACION"] != System.DBNull.Value)
                        c.fh_fin_facturacion = Convert.ToDateTime(rm["FH_FIN_FACTURACION"]);

                    c.uso_gas = uso.Get_Uso_Gas(c.id_ps, c.fh_ini_facturacion, c.fh_fin_facturacion);


                    //if (r["PeriodoFacturacion"] != System.DBNull.Value)
                    //    c.periodofacturacion = r["PeriodoFacturacion"].ToString();
                    if (rm["RAZON_SOCIAL"] != System.DBNull.Value)
                        c.razon_social = rm["RAZON_SOCIAL"].ToString();
                    if (rm["CIFNIF"] != System.DBNull.Value)
                        c.cifnif = rm["CIFNIF"].ToString();
                    if (rm["CUPS22"] != System.DBNull.Value)
                        c.cups22 = rm["CUPS22"].ToString();
                    if (rm["FH_EMISION"] != System.DBNull.Value)
                        c.fh_emision = Convert.ToDateTime(rm["FH_EMISION"]);
                    if (rm["C_FACTURA"] != System.DBNull.Value)
                        c.c_factura = rm["C_FACTURA"].ToString();
                    if (rm["CODIFICACION_FACTURA"] != System.DBNull.Value)
                        c.codificacion_factura = rm["CODIFICACION_FACTURA"].ToString();
                    //if (r["TIPOLOGIA_FACTURA"] != System.DBNull.Value)
                    //    c.tipologia_factura = r["TIPOLOGIA_FACTURA"].ToString();

                    if (rm["IMPORTE_IH"] != System.DBNull.Value)
                        c.importe_ih = Convert.ToDouble(rm["IMPORTE_IH"]);

                    if (rm["BASE_IMPONIBLE_IVA"] != System.DBNull.Value)
                        c.base_imponible_iva = Convert.ToDouble(rm["BASE_IMPONIBLE_IVA"]);

                    //if (r["NM_IMPORTE_IVA"] != System.DBNull.Value)
                    //    c.nm_importe_iva = Convert.ToDouble(r["NM_IMPORTE_IVA"]);

                    if (rm["IVA"] != System.DBNull.Value)
                        c.nm_importe_iva = Convert.ToDouble(rm["IVA"]);

                    if (rm["ImporteTotalFactura"] != System.DBNull.Value)
                        c.ImporteTotalFactura = Convert.ToDouble(rm["ImporteTotalFactura"]);

                    if (rm["IFACTURA"] != System.DBNull.Value)
                        c.ImporteTotalFactura_sce = Convert.ToDouble(rm["IFACTURA"]);

                    if (rm["C_FACTURA"] != System.DBNull.Value)
                    {
                        
                        //string dir_ps;
                        //if (dic_dir_suministro.TryGetValue(c.c_factura, out dir_ps))
                        //    c.direccion_punto_suministro = dir_ps;

                        List<EndesaEntity.facturacion.SIGAME_R_20> l_r20;
                        if (dic_R20.TryGetValue(c.c_factura, out l_r20))
                        {
                            num_Impuestos = 0;
                            c.marca = l_r20.Count > 1 ? "Factura Doble IH" : "";
                            for (int j = 0; j < l_r20.Count; j++)
                            {
                                switch (l_r20[j].codConcepto)
                                {
                                    case "IHTI":
                                        num_Impuestos++;
                                        c.IHTI_tipo_impositivo = l_r20[j].tipo_impositivo;
                                        c.IHTI_descripcion = l_r20[j].descripcion;
                                        //c.IHTI_consumo = Math.Abs(l_r20[j].consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1);
                                        c.IHTI_importe = Math.Abs(l_r20[j].importe) * (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1); 
                                        break;
                                    case "IHTD":
                                        num_Impuestos++;
                                        c.IHTD_tipo_impositivo = 0.65;
                                        c.IHTD_descripcion = l_r20[j].descripcion;
                                        //c.IHTD_consumo = Math.Abs(l_r20[j].consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1); 
                                        c.IHTD_importe = Math.Abs(l_r20[j].importe) * (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1);
                                        break;
                                    case "IHTC":
                                        num_Impuestos++;
                                        c.IHTC_tipo_impositivo = 1.15;
                                        c.IHTC_descripcion = l_r20[j].descripcion;
                                        //c.IHTC_consumo = Math.Abs(l_r20[j].consumo) * (c.ImporteTotalFactura_sce < 0 ? -1 : 1); 
                                        c.IHTC_importe = Math.Abs(l_r20[j].importe) * (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1);
                                        break;
                                }
                            }
                        }
                    }


                    if (rm["CONSUMO_TOTAL"] != System.DBNull.Value)
                        c.consumo_total = Math.Abs(Convert.ToInt32(rm["CONSUMO_TOTAL"])) *
                            (c.codificacion_factura == "S" || c.codificacion_factura == "A" ? -1 : 1);

                    switch (num_Impuestos)
                    {
                        case 0:
                            c.marca = "Sin Impuestos";
                            break;
                        case 1:
                            c.marca = "Factura Simple IH";
                            break;
                        case 2:
                            c.marca = "Factura Doble IH";
                            break;
                    }

                    #region Datos para el informe GAS
                    EndesaEntity.facturacion.SIGAME_Informe_IH o;
                    if (!dic_ih.TryGetValue(c.id_ps, out o))
                    {
                        o = new EndesaEntity.facturacion.SIGAME_Informe_IH();
                        o.cups = c.cups22;
                        o.id_ps = c.id_ps;
                        o.nif = c.cifnif;
                        o.cliente = c.razon_social;
                        o.direccion_punto_suministro = c.direccion_punto_suministro;
                        if (rm["Provincia"] != System.DBNull.Value)
                            o.provincia = rm["Provincia"].ToString();

                        o.ihtc = c.IHTC_importe != 0;
                        o.ihtd = c.IHTD_importe != 0;
                        o.ihti = c.IHTI_importe != 0;

                        int x;
                        if (dic_exentos.TryGetValue(c.id_ps, out x))
                            o.exento = true;

                        dic_ih.Add(o.id_ps, o);
                    }
                    else
                    {
                        o.ihtc = c.IHTC_importe != 0;
                        o.ihtd = c.IHTD_importe != 0;
                        o.ihti = c.IHTI_importe != 0;
                        int x;
                        if (!o.exento)                            
                            if (dic_exentos.TryGetValue(c.id_ps, out x))
                                o.exento = true;

                    }
                    #endregion


                    lista.Add(c);
                }
                dbm.CloseConnection();

                //return lista;


                
                //InformeExcel(@"c:\Temp\FAC_SIG_" + fd.Year + "_" + DateTime.Now.ToString("HHmmss") + ".xlsx", lista, dic_ih);
                InformeExcel(nombre_archivo, lista, dic_ih);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //return null;
            }
        }



        private Dictionary<int, EndesaEntity.facturacion.Factura> CargaFacturas(string tabla, DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<int, EndesaEntity.facturacion.Factura> d = new Dictionary<int, EndesaEntity.facturacion.Factura>();
            try
            {
                strSql = "Select ID_PS, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS, FH_FACTURA,"
                    + " ID_FACTURA, FH_ULT_ACTUALIZACION, NM_IMPORTE_BRUTO"
                    + " from " + tabla 
                    + " where "                    
                    + " FH_INI_FACTURACION >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FIN_FACTURACION <= '" + fh.ToString("yyyy-MM-dd") + "' and "
                    + " FH_FACTURA >= '" + fd.AddMonths(1).ToString("yyyy-MM-01") + "' and "
                    + " CD_NFACTURA_REALES_PS is not null and"
                    + " instr(CD_NFACTURA_REALES_PS,'S') < 1"
                    + " ORDER BY FH_ULT_ACTUALIZACION";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if(r["CD_NFACTURA_REALES_PS"].ToString().Contains("N") || r["CD_NFACTURA_REALES_PS"].ToString().Contains("Y"))
                    {
                        EndesaEntity.facturacion.Factura c = new EndesaEntity.facturacion.Factura();
                        c.id_pto_suministro = Convert.ToInt32(r["ID_PS"]);
                        c.fecha_expedicion_factura = Convert.ToDateTime(r["FH_ULT_ACTUALIZACION"]);
                        c.ifactura = Convert.ToDouble(r["NM_IMPORTE_BRUTO"]);
                        c.ffactura = Convert.ToDateTime(r["FH_FACTURA"]);

                        EndesaEntity.facturacion.Factura o;
                        if (!dic_fo_s.TryGetValue(c.id_pto_suministro, out o))
                            dic_fo_s.Add(c.id_pto_suministro, o);
                    }

                    
                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }






        public void CargaTablaBase()
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            int year = 2016;
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            string cd_crefaext = "";
            Dictionary<string, List<EndesaEntity.facturacion.SIGAME_R_20>> d_r20 = new Dictionary<string, List<EndesaEntity.facturacion.SIGAME_R_20>>();
            Dictionary<string, EndesaEntity.facturacion.SIGAME_TablaBase> d = new Dictionary<string, EndesaEntity.facturacion.SIGAME_TablaBase>();            

            try
            {

                BorrarTabla("sigame_tablabase");

                fd = new DateTime(year, 1, 1);
                fh = fd.AddYears(1).AddDays(-1);
                d_r20 = CargaConceptosR20(fd, fh);


                CargaTablaBaseFacturasSCE(fd, fh, d_r20);

                strSql = " SELECT month(FH_FACTURA) as PERIODO,"
                    + " T_SGM_M_FACTURAS_REALES_PS.FH_INI_FACTURACION, T_SGM_M_FACTURAS_REALES_PS.FH_FIN_FACTURACION, Int_SCE_R_10.PeriodoFacturacion as PeriodoFacturacion,"
                    + " T_SGM_G_PS.DE_PS AS RAZON_SOCIAL, T_SGM_M_CLIENTES.CD_CIF as CIFNIF, T_SGM_G_PS.CD_CUPS as CUPS22,"
                    + " FH_FACTURA as FH_EMISION,"
                    + " T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT,"
                    + " CD_NFACTURA_REALES_PS as C_FACTURA, CD_TESTFACT as CODIFICACION_FACTURA,TX_TIPO_FACTURA_NUEVO as TIPOLOGIA_FACTURA,"
                    + " RTRIM(Int_SCE_R_10.DireccionPS) as  DireccionPuntoSuministro, RTRIM(Int_SCE_R_20.Descripcion) as Descripcion,"
                    + " sum(CONVERT(Decimal,case when right(Int_SCE_R_30.ConsumoTotal,1)='1' then Left(Int_SCE_R_30.ConsumoTotal,Len(Int_SCE_R_30.ConsumoTotal)-1) else Left(Int_SCE_R_30.ConsumoTotal,Len(Int_SCE_R_30.ConsumoTotal)-1)*-1 end)) AS ConsumoTotal,"                    
                    + " Int_SCE_R_20.CodConcepto,replace(replace(Int_SCE_R_25.Concepto,'Impto. hidrocarburos ',''),' Eur/GJulio','') as TIPO_IMPOSITIVO_IH," 
                    + " NM_IEH_IMPORTE as IMPORTE_IH, NM_BASE_IMPONIBLE as BASE_IMPONIBLE_IVA, NM_IMPORTE_IVA,"
                    + " NM_IMPORTE_BRUTO As ImporteTotalFactura"
                    + " FROM ((((((T_SGM_M_FACTURAS_REALES_PS"
                    + " INNER JOIN T_SGM_G_CONTRATOS_PS ON T_SGM_M_FACTURAS_REALES_PS.ID_CTO_PS = T_SGM_G_CONTRATOS_PS.ID_CTO_PS)"
                    + " INNER JOIN T_SGM_G_PS ON T_SGM_G_CONTRATOS_PS.ID_PS = T_SGM_G_PS.ID_PS)"
                    + " INNER JOIN T_SGM_M_CLIENTES ON T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE)"
                    + " INNER JOIN Int_SCE_R_10 ON T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura)"
                    + " INNER JOIN Int_SCE_R_20 ON Int_SCE_R_10.NumRefFactura = Int_SCE_R_20.NumRefFactura)"
                    + " INNER JOIN Int_SCE_R_30 ON (Int_SCE_R_20.NumRefFactura = Int_SCE_R_30.NumRefFactura)"
                    + " AND (Int_SCE_R_20.FxGenInterfaz = Int_SCE_R_30.FxGenInterfaz))"
                    + " INNER JOIN Int_SCE_R_25 ON Int_SCE_R_20.CodConcepto = Int_SCE_R_25.CodConcepto"
                    + " WHERE Year(FH_FACTURA) = " + year
                    // + " and CD_NFACTURA_REALES_PS = '00Z604N0000537'"
                    // + " And Left(T_SGM_G_PS.CD_CUPS,2)='ES'"
                    // + " and Int_SCE_R_20.CodConcepto in ('IHTC','IHTD','IHTI')"
                    + " Group by"
                    + " month(FH_FACTURA) ,T_SGM_M_FACTURAS_REALES_PS.FH_INI_FACTURACION, T_SGM_M_FACTURAS_REALES_PS.FH_FIN_FACTURACION," 
                    + " Int_SCE_R_10.PeriodoFacturacion,T_SGM_G_PS.DE_PS , T_SGM_M_CLIENTES.CD_CIF, T_SGM_G_PS.CD_CUPS, FH_FACTURA," 
                    + " T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT,"
                    + " CD_NFACTURA_REALES_PS," 
                    + " CD_TESTFACT,TX_TIPO_FACTURA_NUEVO, RTRIM(Int_SCE_R_10.DireccionPS), RTRIM(Int_SCE_R_20.Descripcion)," 
                    + " Int_SCE_R_20.CodConcepto,replace(replace(Int_SCE_R_25.Concepto,'Impto. hidrocarburos ',''),' Eur/GJulio','') ," 
                    + " NM_IEH_IMPORTE, NM_BASE_IMPONIBLE, NM_IMPORTE_IVA,NM_IMPORTE_BRUTO"
                    + " order by T_SGM_G_PS.CD_CUPS asc, FH_FACTURA asc";

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["CIFNIF"].ToString().Substring(0, 2) != "PT")
                    {


                        EndesaEntity.facturacion.SIGAME_TablaBase c = new EndesaEntity.facturacion.SIGAME_TablaBase();
                        c.origen = "SIGAME";
                        c.periodo = Convert.ToInt32(r["PERIODO"]);
                        if (r["FH_INI_FACTURACION"] != System.DBNull.Value)
                            c.fh_ini_facturacion = Convert.ToDateTime(r["FH_INI_FACTURACION"]);
                        if (r["FH_FIN_FACTURACION"] != System.DBNull.Value)
                            c.fh_fin_facturacion = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);
                        if (r["PeriodoFacturacion"] != System.DBNull.Value)
                            c.periodofacturacion = r["PeriodoFacturacion"].ToString();
                        if (r["RAZON_SOCIAL"] != System.DBNull.Value)
                            c.razon_social = r["RAZON_SOCIAL"].ToString();
                        if (r["CIFNIF"] != System.DBNull.Value)
                            c.cifnif = r["CIFNIF"].ToString();
                        if (r["CUPS22"] != System.DBNull.Value)
                            c.cups22 = r["CUPS22"].ToString();
                        if (r["FH_EMISION"] != System.DBNull.Value)
                            c.fh_emision = Convert.ToDateTime(r["FH_EMISION"]);
                        if (r["C_FACTURA"] != System.DBNull.Value)
                            c.c_factura = r["C_FACTURA"].ToString();
                        if (r["CODIFICACION_FACTURA"] != System.DBNull.Value)
                            c.codificacion_factura = r["CODIFICACION_FACTURA"].ToString();
                        if (r["TIPOLOGIA_FACTURA"] != System.DBNull.Value)
                            c.tipologia_factura = r["TIPOLOGIA_FACTURA"].ToString();
                        if (r["DireccionPuntoSuministro"] != System.DBNull.Value)
                            c.direccion_punto_suministro = r["DireccionPuntoSuministro"].ToString();

                        if (r["BASE_IMPONIBLE_IVA"] != System.DBNull.Value)
                            c.base_imponible_iva = Convert.ToDouble(r["BASE_IMPONIBLE_IVA"]);

                        if (r["NM_IMPORTE_IVA"] != System.DBNull.Value)
                            c.nm_importe_iva = Convert.ToDouble(r["NM_IMPORTE_IVA"]);

                        if (r["ImporteTotalFactura"] != System.DBNull.Value)
                            c.ImporteTotalFactura = Convert.ToDouble(r["ImporteTotalFactura"]);

                        cd_crefaext = r["CD_CREFAEXT"].ToString();

                        if (r["ConsumoTotal"] != System.DBNull.Value)
                            c.consumo_total = Convert.ToInt32(r["ConsumoTotal"]);

                        EndesaEntity.facturacion.SIGAME_TablaBase o;
                        if (d.TryGetValue(c.c_factura, out o))
                        {
                         
                            for (int i = 0; i < o.l.Count; i++)
                            {
                                if (o.l[i].cod_concepto == r["CodConcepto"].ToString())
                                {
                                    if (r["TIPO_IMPOSITIVO_IH"] != System.DBNull.Value)
                                        o.l[i].tipo_impositivo_ih = r["TIPO_IMPOSITIVO_IH"].ToString();

                                    if (r["IMPORTE_IH"] != System.DBNull.Value)
                                        o.l[i].importe_ih = Convert.ToDouble(r["IMPORTE_IH"]);

                                    if (r["Descripcion"] != System.DBNull.Value)
                                        o.l[i].descripcion = r["Descripcion"].ToString();

                                    if (r["CodConcepto"] != System.DBNull.Value)
                                        o.l[i].cod_concepto = r["CodConcepto"].ToString();
                                }
                            }

                        }
                        else
                        {
                            List<EndesaEntity.facturacion.SIGAME_R_20> r_20;
                            if (d_r20.TryGetValue(cd_crefaext, out r_20))
                            {

                                for (int i = 0; i < r_20.Count; i++)
                                {
                                    EndesaEntity.facturacion.SIGAME_TablaBaseDetalle dd = new EndesaEntity.facturacion.SIGAME_TablaBaseDetalle();

                                    if (r_20[i].codConcepto == r["CodConcepto"].ToString())
                                    {
                                        if (r["TIPO_IMPOSITIVO_IH"] != System.DBNull.Value)
                                            dd.tipo_impositivo_ih = r["TIPO_IMPOSITIVO_IH"].ToString();

                                        if (r["IMPORTE_IH"] != System.DBNull.Value)
                                            dd.importe_ih = Convert.ToDouble(r["IMPORTE_IH"]);
                                    }

                                    dd.descripcion = r_20[i].descripcion;
                                    dd.cod_concepto = r_20[i].codConcepto;
                                    dd.importe_ih_parcial = r_20[i].importe;
                                    dd.consumo_parcial = r_20[i].consumo;
                                    c.l.Add(dd);

                                }

                                if(r_20.Count == 0)
                                {
                                    EndesaEntity.facturacion.SIGAME_TablaBaseDetalle dd = new EndesaEntity.facturacion.SIGAME_TablaBaseDetalle();
                                    if (r["TIPO_IMPOSITIVO_IH"] != System.DBNull.Value)
                                        dd.tipo_impositivo_ih = r["TIPO_IMPOSITIVO_IH"].ToString();

                                    if (r["IMPORTE_IH"] != System.DBNull.Value)
                                        dd.importe_ih = Convert.ToDouble(r["IMPORTE_IH"]);

                                    if (r["Descripcion"] != System.DBNull.Value)
                                        dd.descripcion = r["Descripcion"].ToString();

                                    if (r["CodConcepto"] != System.DBNull.Value)
                                        dd.cod_concepto = r["CodConcepto"].ToString();

                                    c.l.Add(dd);
                                }

                                d.Add(c.c_factura, c);

                            }

                        }
                    }

                }

                db.CloseConnection();

                GuardarEnMySQLTablaBase(d);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void CargaTablaBaseFacturasSCE(DateTime fd, DateTime fh, Dictionary<string, List<EndesaEntity.facturacion.SIGAME_R_20>> d_r20)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";            
            
            string cd_crefaext = "";
            
            Dictionary<string, EndesaEntity.facturacion.SIGAME_TablaBase> d = new Dictionary<string, EndesaEntity.facturacion.SIGAME_TablaBase>();

            try
            {
                strSql = " SELECT 'SCE' AS origen, NULL AS marca, MONTH(s.FH_FACTURA) AS periodo,"
                    + " s.FH_INI_FACTURACION AS fh_ini_facturacion, s.FH_FIN_FACTURACION AS fh_fin_facturacion,"
                    + " NULL AS periodofacturacion, f.DAPERSOC as RAZON_SOCIAL, f.CNIFDNIC AS CIFNIF,"
                    + " f.CUPSREE AS cups22,"
                    + " s.FH_FACTURA AS fh_emision, s.CD_NFACTURA_REALES_PS AS c_factura,"
                    + " f.TESTFACT AS codificacion_factura, NULL AS tipologia_factura,"
                    + " null AS direccion_punto_suministro, NULL AS descripcion,"
                    + " s.NM_CONSUMO AS consumo_total, NULL AS consumo_parcial,"
                    + " NULL AS cod_concepto, NULL AS tipo_impositivo_ih,"
                    + " NULL AS importe_ih, NULL AS importe_ih_parcial,"
                    + " s.NM_BASE_IMPONIBLE as BASE_IMPONIBLE_IVA, s.NM_IMPORTE_IVA,"
                    + " f.IFACTURA AS ImporteTotalFactura"
                    + " FROM fo_s_sce s INNER JOIN"
                    + " fo f ON"
                    + " f.CFACTURA = s.CD_NFACTURA_REALES_PS"
                    + " WHERE(s.FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                    + " s.FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd")  + "')"
                    + " AND substr(f.CNIFDNIC,1,2) <> 'PT'";


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                
                while (r.Read())
                {                                       


                        EndesaEntity.facturacion.SIGAME_TablaBase c = new EndesaEntity.facturacion.SIGAME_TablaBase();
                        c.origen = "SCE";
                        c.periodo = Convert.ToInt32(r["PERIODO"]);
                        if (r["FH_INI_FACTURACION"] != System.DBNull.Value)
                            c.fh_ini_facturacion = Convert.ToDateTime(r["FH_INI_FACTURACION"]);
                        if (r["FH_FIN_FACTURACION"] != System.DBNull.Value)
                            c.fh_fin_facturacion = Convert.ToDateTime(r["FH_FIN_FACTURACION"]);
                        if (r["PeriodoFacturacion"] != System.DBNull.Value)
                            c.periodofacturacion = r["PeriodoFacturacion"].ToString();
                        if (r["RAZON_SOCIAL"] != System.DBNull.Value)
                            c.razon_social = r["RAZON_SOCIAL"].ToString();
                        if (r["CIFNIF"] != System.DBNull.Value)
                            c.cifnif = r["CIFNIF"].ToString();
                        if (r["CUPS22"] != System.DBNull.Value)
                            c.cups22 = r["CUPS22"].ToString();
                        if (r["FH_EMISION"] != System.DBNull.Value)
                            c.fh_emision = Convert.ToDateTime(r["FH_EMISION"]);
                        if (r["C_FACTURA"] != System.DBNull.Value)
                            c.c_factura = r["C_FACTURA"].ToString();
                        if (r["CODIFICACION_FACTURA"] != System.DBNull.Value)
                            c.codificacion_factura = r["CODIFICACION_FACTURA"].ToString();
                        if (r["TIPOLOGIA_FACTURA"] != System.DBNull.Value)
                            c.tipologia_factura = r["TIPOLOGIA_FACTURA"].ToString();
                        if (r["DireccionPuntoSuministro"] != System.DBNull.Value)
                            c.direccion_punto_suministro = r["DireccionPuntoSuministro"].ToString();

                        if (r["BASE_IMPONIBLE_IVA"] != System.DBNull.Value)
                            c.base_imponible_iva = Convert.ToDouble(r["BASE_IMPONIBLE_IVA"]);

                        if (r["NM_IMPORTE_IVA"] != System.DBNull.Value)
                            c.nm_importe_iva = Convert.ToDouble(r["NM_IMPORTE_IVA"]);

                        if (r["ImporteTotalFactura"] != System.DBNull.Value)
                            c.ImporteTotalFactura = Convert.ToDouble(r["ImporteTotalFactura"]);

                        

                        if (r["ConsumoTotal"] != System.DBNull.Value)
                            c.consumo_total = Convert.ToInt32(r["ConsumoTotal"]);

                        EndesaEntity.facturacion.SIGAME_TablaBase o;
                        if (!d.TryGetValue(c.c_factura, out o))
                        {

                           
                            List<EndesaEntity.facturacion.SIGAME_R_20> r_20;
                            if (d_r20.TryGetValue(c.c_factura, out r_20))
                            {

                                for (int i = 0; i < r_20.Count; i++)
                                {
                                    EndesaEntity.facturacion.SIGAME_TablaBaseDetalle dd = new EndesaEntity.facturacion.SIGAME_TablaBaseDetalle();
                                    switch (r_20[i].codConcepto)
                                    {
                                        case "IHTI":
                                            dd.tipo_impositivo_ih = "0,15";
                                            break;
                                        case "IHTD":
                                            dd.tipo_impositivo_ih = "0,65";
                                            break;
                                        case "IHTC":
                                            dd.tipo_impositivo_ih = "1,15";
                                            break;

                                    }
                                    dd.tipo_impositivo_ih =
                                    dd.descripcion = r_20[i].descripcion;
                                    dd.cod_concepto = r_20[i].codConcepto;
                                    dd.importe_ih_parcial = r_20[i].importe;
                                    dd.consumo_parcial = r_20[i].consumo;
                                    c.l.Add(dd);

                                }

                            }

                                d.Add(c.c_factura, c);

                            
                        }
                    }

                

                db.CloseConnection();

                GuardarEnMySQLTablaBase(d);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private Dictionary<string, List<EndesaEntity.facturacion.SIGAME_R_20>> CargaConceptosR20(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;

            MySQLDB dbm;
            MySqlCommand commandm;
            MySqlDataReader rm;
            string strSql = "";
                        
            string texto = "";
            string importe = "";
            string signo_importe = "";
            int total_registros = 0;
            int i = 0;

            Dictionary<string, List<EndesaEntity.facturacion.SIGAME_R_20>> d =
                new Dictionary<string, List<EndesaEntity.facturacion.SIGAME_R_20>>();
            try
            {
                #region ParteImpuestos SIGAME
                Console.WriteLine("Cargando registros de Int_SCE_R_20.CodConcepto");

                strSql = "select count(*) as total_registros from "
                    + " T_SGM_M_FACTURAS_REALES_PS "
                    + " INNER JOIN Int_SCE_R_10 ON T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura"
                    + " INNER JOIN Int_SCE_R_20 ON Int_SCE_R_10.NumRefFactura = Int_SCE_R_20.NumRefFactura"
                    + " where "
                    + " (FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "') and"
                    + " Int_SCE_R_20.CodConcepto in('IHTC','IHTD','IHTI', 'IHCP', 'IHCG', 'IHCC')";

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    total_registros = Convert.ToInt32(r["total_registros"]);
                }
                db.CloseConnection();


                strSql = "select T_SGM_M_FACTURAS_REALES_PS.CD_NFACTURA_REALES_PS,"
                    + " Int_SCE_R_20.CodConcepto, Int_SCE_R_20.NumRefFactura, Int_SCE_R_20.Importe, Int_SCE_R_20.Descripcion"
                    + " from T_SGM_M_FACTURAS_REALES_PS"
                    + " INNER JOIN Int_SCE_R_10 ON T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura"
                    + " INNER JOIN Int_SCE_R_20 ON Int_SCE_R_10.NumRefFactura = Int_SCE_R_20.NumRefFactura"
                    + " where"
                    // + " T_SGM_M_FACTURAS_REALES_PS.CD_NFACTURA_REALES_PS = '00Z604S0001194' and"
                    + " (FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "') and"
                    + " Int_SCE_R_20.CodConcepto in('IHTC','IHTD','IHTI', 'IHCP', 'IHCG', 'IHCC')"
                    + " group by T_SGM_M_FACTURAS_REALES_PS.CD_NFACTURA_REALES_PS,"
                    + " Int_SCE_R_20.CodConcepto, Int_SCE_R_20.NumRefFactura, Int_SCE_R_20.Importe, Int_SCE_R_20.Descripcion";
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    i++;
                    //Console.CursorLeft = 0;
                    //Console.Write(i + " de " + total_registros);

                    EndesaEntity.facturacion.SIGAME_R_20 c = new EndesaEntity.facturacion.SIGAME_R_20();


                    if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        c.cfactura = r["CD_NFACTURA_REALES_PS"].ToString();
                    if (r["CodConcepto"] != System.DBNull.Value)
                        c.codConcepto = r["CodConcepto"].ToString();
                    if (r["NumRefFactura"] != System.DBNull.Value)
                        c.numRefFactura = r["NumRefFactura"].ToString();
                    if (r["Importe"] != System.DBNull.Value)
                    {
                        importe = r["Importe"].ToString();
                        signo_importe = importe.Substring(importe.Length - 1, 1);                         
                        if (signo_importe == "1")
                            c.importe = (Convert.ToDouble(importe.Substring(0, importe.Length - 1)) / 100);
                        else
                            c.importe = (Convert.ToDouble(importe.Substring(0, importe.Length - 1)) / 100) * -1;                       
                    }
                        
                    if (r["Descripcion"] != System.DBNull.Value)
                    {
                        texto = "";
                        texto = r["Descripcion"].ToString();
                        c.descripcion = r["Descripcion"].ToString().Trim();
                        texto = texto.Substring(texto.IndexOf("x") + 1);
                        texto = texto.Replace("kWh =", "");
                        texto = texto.Replace(".", "").Trim();
                        if (signo_importe == "1")
                            c.consumo = Convert.ToInt32(texto);
                        else
                            c.consumo = Convert.ToInt32(texto) * -1;

                    }

                    switch (c.codConcepto)
                    {
                        case "IHTI":
                            c.tipo_impositivo = 0.15;
                            break;
                        case "IHCP":
                            c.tipo_impositivo = 0.15;
                            break;
                        case "IHTD":
                            c.tipo_impositivo = 0.65;
                            break;
                        case "IHCG":
                            c.tipo_impositivo = 0.65;
                            break;
                        case "IHTC":
                            c.tipo_impositivo = 1.15;
                            break;
                        case "IHCC":
                            c.tipo_impositivo = 1.15;
                            break;


                    }
                    
                                                                          
                    List <EndesaEntity.facturacion.SIGAME_R_20> o;
                    //if (!d.TryGetValue(c.numRefFactura, out o))
                    if (!d.TryGetValue(c.cfactura, out o))
                    {
                        o = new List<EndesaEntity.facturacion.SIGAME_R_20>();
                        o.Add(c);
                        d.Add(c.cfactura, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();

                #endregion

                #region Parte SCE


                strSql = "select s.CD_NFACTURA_REALES_PS as c_factura, v.*"
                    + " from fact.fo_s_sce s"
                    + " INNER JOIN fact.fo f on"
                    + " f.cfactura = s.CD_NFACTURA_REALES_PS "
                    + " inner join fact.fo_tcon_v v on"
                    + " v.CREFEREN = f.CREFEREN  AND"
                    + " v.SECFACTU = f.SECFACTU AND"
                    + " v.TESTFACT = f.TESTFACT"
                    + " WHERE"
                    + " (s.FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                    + " s.FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "') AND"
                    + " v.DESCRIPCION_CORTA  in ('IHTC', 'IHTD', 'IHTI', 'IHCP', 'IHCC', 'IHCG')";


                dbm = new MySQLDB(MySQLDB.Esquemas.FAC);
                commandm = new MySqlCommand(strSql, dbm.con);
                rm = commandm.ExecuteReader();
                while (rm.Read())
                {
                    EndesaEntity.facturacion.SIGAME_R_20 c = new EndesaEntity.facturacion.SIGAME_R_20();
                    if (rm["c_factura"] != System.DBNull.Value)
                        c.cfactura = rm["c_factura"].ToString();
                    if (rm["DESCRIPCION_CORTA"] != System.DBNull.Value)
                        c.codConcepto = rm["DESCRIPCION_CORTA"].ToString();
                    if (rm["c_factura"] != System.DBNull.Value)
                        c.numRefFactura = rm["c_factura"].ToString();
                    if (rm["ICONFAC"] != System.DBNull.Value)
                        c.importe = Convert.ToDouble(rm["ICONFAC"]);
                    if (rm["DESCRIPCION"] != System.DBNull.Value)
                        c.descripcion = rm["DESCRIPCION"].ToString();

                    List<EndesaEntity.facturacion.SIGAME_R_20> o;
                    if (!d.TryGetValue(c.numRefFactura, out o))
                    {
                        o = new List<EndesaEntity.facturacion.SIGAME_R_20>();
                        o.Add(c);
                        d.Add(c.numRefFactura, o);
                    }
                    else
                    {
                        bool encontrado = false;
                        for(int h = 0; h < o.Count; h++)
                        {
                            if (!encontrado)
                                encontrado = (o[h].codConcepto == c.codConcepto);
                        }
                        if (!encontrado)
                            o.Add(c);
                    }
                        
                }
                dbm.CloseConnection();
                #endregion

                //Console.WriteLine();
                return d;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        private Dictionary<string, EndesaEntity.facturacion.SIGAME_R_30> CargaConceptosR30(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;

            string strSql = "";
            string texto = "";
            int total_registros = 0;
            int i = 0;

            Dictionary<string, EndesaEntity.facturacion.SIGAME_R_30> d =
                new Dictionary<string, EndesaEntity.facturacion.SIGAME_R_30>();
                
            try
            {
                #region ParteImpuestos SIGAME
                Console.WriteLine("Cargando registros de Int_SCE_R_30");

                strSql = "select count(*) as total_registros from "
                    + " T_SGM_M_FACTURAS_REALES_PS "
                    + " INNER JOIN Int_SCE_R_10 ON T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura"
                    + " INNER JOIN Int_SCE_R_20 ON Int_SCE_R_10.NumRefFactura = Int_SCE_R_20.NumRefFactura"
                    + " INNER JOIN Int_SCE_R_30 ON Int_SCE_R_20.NumRefFactura = Int_SCE_R_30.NumRefFactura"
                    + " AND (Int_SCE_R_20.FxGenInterfaz = Int_SCE_R_30.FxGenInterfaz)"
                    + " where "
                    + " (FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                    
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    total_registros = Convert.ToInt32(r["total_registros"]);
                }
                db.CloseConnection();


                strSql = "select Int_SCE_R_30.NumRefFactura, Int_SCE_R_30.FxGenInterfaz, Int_SCE_R_30.ConsumoTotal"
                    + " from T_SGM_M_FACTURAS_REALES_PS "
                    + " INNER JOIN Int_SCE_R_10 ON T_SGM_M_FACTURAS_REALES_PS.CD_CREFAEXT = Int_SCE_R_10.NumRefFactura"
                    + " INNER JOIN Int_SCE_R_20 ON Int_SCE_R_10.NumRefFactura = Int_SCE_R_20.NumRefFactura"
                    + " INNER JOIN Int_SCE_R_30 ON Int_SCE_R_20.NumRefFactura = Int_SCE_R_30.NumRefFactura"
                    + " AND (Int_SCE_R_20.FxGenInterfaz = Int_SCE_R_30.FxGenInterfaz)"
                    + " where "
                    + " (FH_FACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    i++;
                    Console.CursorLeft = 0;
                    Console.Write(i + " de " + total_registros);

                    EndesaEntity.facturacion.SIGAME_R_30 c = new EndesaEntity.facturacion.SIGAME_R_30();

                    
                    if (r["NumRefFactura"] != System.DBNull.Value)
                        c.numRefFactura = r["NumRefFactura"].ToString();
                    if (r["FxGenInterfaz"] != System.DBNull.Value)
                        c.fxGenInterfaz = r["FxGenInterfaz"].ToString();
                    if (r["ConsumoTotal"] != System.DBNull.Value)
                    {
                        texto = r["ConsumoTotal"].ToString();
                        if(texto.Substring(texto.Length - 1, 1) == "1")
                            c.consumo = Convert.ToInt32(texto.Substring(0, texto.Length - 1));
                        else
                            c.consumo = Convert.ToInt32(texto.Substring(0, texto.Length - 1)) * -1;
                    } 

                    EndesaEntity.facturacion.SIGAME_R_30 o;
                    if (!d.TryGetValue(c.numRefFactura, out o))                        
                        d.Add(c.numRefFactura, c);
                    
                    

                }
                db.CloseConnection();

                #endregion

                Console.WriteLine();

                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        private Dictionary<string, EndesaEntity.facturacion.SIGAME_DireccionPS> CargaDirPS()
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;

            string strSql = "";
            
            int total_registros = 0;
            int i = 0;

            Dictionary<string, EndesaEntity.facturacion.SIGAME_DireccionPS> d = 
                new Dictionary<string, EndesaEntity.facturacion.SIGAME_DireccionPS>();
            try
            {
                #region ParteImpuestos SIGAME
                Console.WriteLine("Cargando registros de Int_SCE_R_10");

                strSql = "select count(*) as total_registros from INT_SCE.Int_SCE_R_10";
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    total_registros = Convert.ToInt32(r["total_registros"]);
                }
                db.CloseConnection();


                strSql = "select NumRefFactura, DireccionPS, CUPS22  from INT_SCE.Int_SCE_R_10";
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    i++;
                    Console.CursorLeft = 0;
                    Console.Write(i + " de " + total_registros);


                    EndesaEntity.facturacion.SIGAME_DireccionPS c = 
                        new EndesaEntity.facturacion.SIGAME_DireccionPS();

                    c.numRefFactura = r["NumRefFactura"].ToString();
                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups = r["CUPS22"].ToString();

                    EndesaEntity.facturacion.SIGAME_DireccionPS o;
                    if (!d.TryGetValue(c.numRefFactura, out o))
                        d.Add(c.numRefFactura, c);
                    

                }
                db.CloseConnection();

                #endregion

                

                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        private void GuardarEnMySQLTablaBase(Dictionary<string, EndesaEntity.facturacion.SIGAME_TablaBase> d)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            bool firstOnly2 = true;
            int registros = 0;
            int totalregistros = 0;
                       
            try
            {
                               
                                               
                foreach (KeyValuePair<string,EndesaEntity.facturacion.SIGAME_TablaBase> p in d)
                {
                    registros++;
                    totalregistros++;
                    if (firstOnly)
                    {
                        sb.Append("replace into sigame_tablabase");
                        sb.Append(" (origen, marca, periodo, fh_ini_facturacion, fh_fin_facturacion,");
                        sb.Append(" periodofacturacion, razon_social, cifnif,");
                        sb.Append(" cups22, fh_emision, c_factura, codificacion_factura, tipologia_factura,");
                        sb.Append(" direccion_punto_suministro, descripcion, consumo_total, consumo_parcial,");
                        sb.Append(" cod_concepto, tipo_impositivo_ih, importe_ih, importe_ih_parcial, base_imponible_iva,");
                        sb.Append(" nm_importe_iva, ImporteTotalFactura) values ");
                        firstOnly = false;
                    }

                    firstOnly2 = true;
                    for(int i = 0; i < p.Value.l.Count; i++)
                    {

                        sb.Append("('").Append(p.Value.origen).Append("',");
                        sb.Append(p.Value.origen == "SIGAME" && p.Value.l.Count > 1 ? "'Factura Doble IH'" : "null").Append(",");
                        sb.Append(p.Value.periodo).Append(",");

                        if(p.Value.fh_ini_facturacion > DateTime.MinValue)
                            sb.Append("'").Append(p.Value.fh_ini_facturacion.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.fh_fin_facturacion > DateTime.MinValue)
                            sb.Append("'").Append(p.Value.fh_fin_facturacion.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.periodofacturacion != null)
                            sb.Append("'").Append(p.Value.periodofacturacion).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.razon_social != null)
                            sb.Append("'").Append(p.Value.razon_social).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.cifnif != null)
                            sb.Append("'").Append(p.Value.cifnif).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.cups22 != null)
                            sb.Append("'").Append(p.Value.cups22).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.fh_emision > DateTime.MinValue)
                            sb.Append("'").Append(p.Value.fh_emision.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.c_factura != null)
                            sb.Append("'").Append(p.Value.c_factura).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.codificacion_factura != null)
                            sb.Append("'").Append(p.Value.codificacion_factura).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.tipologia_factura != null)
                            sb.Append("'").Append(p.Value.tipologia_factura).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.direccion_punto_suministro != null)
                            sb.Append("'").Append(p.Value.direccion_punto_suministro).Append("',");
                        else
                            sb.Append("null,");

                        if (p.Value.l[i].descripcion != null)
                            sb.Append("'").Append(p.Value.l[i].descripcion).Append("',");
                        else
                            sb.Append("null,");

                        if (firstOnly2)
                        {
                            sb.Append(p.Value.consumo_total).Append(",");
                            sb.Append(p.Value.l[i].consumo_parcial).Append(",");

                            if (p.Value.l[i].cod_concepto != null)
                                sb.Append("'").Append(p.Value.l[i].cod_concepto).Append("',");
                            else
                                sb.Append("null,");

                            if (p.Value.l[i].tipo_impositivo_ih != null)
                                sb.Append("'").Append(p.Value.l[i].tipo_impositivo_ih).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append(p.Value.l[i].importe_ih.ToString().Replace(",", ".")).Append(",");
                            sb.Append(p.Value.l[i].importe_ih_parcial.ToString().Replace(",", ".")).Append(",");
                            sb.Append(p.Value.base_imponible_iva.ToString().Replace(",", ".")).Append(",");
                            sb.Append(p.Value.nm_importe_iva.ToString().Replace(",", ".")).Append(",");
                            sb.Append(p.Value.ImporteTotalFactura.ToString().Replace(",", ".")).Append("),");
                            firstOnly2 = false;
                        }
                        else
                        {
                            sb.Append("null,"); // sb.Append(p.Value[i].consumo_total).Append(",");
                            sb.Append(p.Value.l[i].consumo_parcial).Append(",");

                            if (p.Value.l[i].cod_concepto != null)
                                sb.Append("'").Append(p.Value.l[i].cod_concepto).Append("',");
                            else
                                sb.Append("null,");

                            if (p.Value.l[i].tipo_impositivo_ih != null)
                                sb.Append("'").Append(p.Value.l[i].tipo_impositivo_ih).Append("',");
                            else
                                sb.Append("null,");

                            sb.Append("null,"); // sb.Append(p.Value[i].importe_ih.ToString().Replace(",", ".")).Append(",");
                            sb.Append(p.Value.l[i].importe_ih_parcial.ToString().Replace(",", ".")).Append(",");
                            sb.Append("null,"); //sb.Append(p.Value[i].base_imponible_iva.ToString().Replace(",", ".")).Append(",");
                            sb.Append("null,"); //sb.Append(p.Value[i].nm_importe_iva.ToString().Replace(",", ".")).Append(",");
                            sb.Append("null),");  // sb.Append(p.Value[i].ImporteTotalFactura.ToString().Replace(",", ".")).Append("),");
                            
                        }

                        
                    }


                    if (registros == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write("Guardando " + totalregistros + " de " + d.Count);
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        registros = 0;
                    }


                }

                

                if (registros > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write("Guardando " + totalregistros + " de " + d.Count);
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    registros = 0;
                }


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void BorrarTabla(string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            Console.WriteLine("Borrando tabla " + tabla);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand("delete from " + tabla, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public void InformeExcel(string fichero, List<EndesaEntity.facturacion.SIGAME_TablaBase> l, 
            Dictionary<int, EndesaEntity.facturacion.SIGAME_Informe_IH> dd)
        {
            int f = 1;
            int c = 1;
            bool con_Detalle = true;


            try
            {

                FileInfo fileInfo = new FileInfo(fichero);

                if (fileInfo.Exists)
                    fileInfo.Delete();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("FACTURAS");

                var headerCells = workSheet.Cells[1, 1, 1, 38];
                var headerFont = headerCells.Style.Font;
                f = 1;

                headerFont.Bold = true;



                #region Cabecera_Excel

                workSheet.Cells[f, c].Value = "ORIGEN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "MARCA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                if (con_Detalle)
                {
                    workSheet.Cells[f, c].Value = "RAZON SOCIAL";
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                    workSheet.Cells[f, c].Value = "DIRECCION PUNTO SUMINISTRO";
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                }

                workSheet.Cells[f, c].Value = "ID PS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CUPS22";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                workSheet.Cells[f, c].Value = "C_FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CODIFICACION_FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TIPOLOGIA_FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                if (con_Detalle)
                {
                    workSheet.Cells[f, c].Value = "PERIODO";
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                }

                workSheet.Cells[f, c].Value = "SECFACTU";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FH_EMISION";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                workSheet.Cells[f, c].Value = "FH_INI_FACTURACION";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "FH_FIN_FACTURACION";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "PERIODO_FACTURACION";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                workSheet.Cells[f, c].Value = "CONSUMO_TOTAL";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                //workSheet.Cells[f, c].Value = "IMPORTE_IH";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                if (con_Detalle)
                {
                    workSheet.Cells[f, c].Value = "DESCRIPCION T. IMPOSITIVO 0.15";
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                }


                workSheet.Cells[f, c].Value = "T. IMPOSITIVO 0.15";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                //workSheet.Cells[f, c].Value = "CONSUMO_IHTI";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "IMPORTE T. IMPOSITIVO 0.15";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                if (con_Detalle)
                {
                    workSheet.Cells[f, c].Value = "DESCRIPCION T. IMPOSITIVO 0.65";
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                }

                workSheet.Cells[f, c].Value = "T. IMPOSITIVO 0.65";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                //workSheet.Cells[f, c].Value = "CONSUMO_IHTD";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "IMPORTE T. IMPOSITIVO 0.65";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                if (con_Detalle)
                {
                    workSheet.Cells[f, c].Value = "DESCRIPCION T. IMPOSITIVO 1.15";
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;
                }


                workSheet.Cells[f, c].Value = "T. IMPOSITIVO 1.15";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                //workSheet.Cells[f, c].Value = "CONSUMO_IHTC";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "IMPORTE T. IMPOSITIVO 1.15";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "BASE_IMPONIBLE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "IVA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "IMPORTE_TOTAL_FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "IMPORTE_TOTAL_FACTURA SCE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TOT F SIGAME - TOT F SCE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "BASE+IVA-FACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "USO GAS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;


                //workSheet.Cells[f, c].Value = "CONSUMO_TOTAL - CONSUMOS_PARCIALES";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                //workSheet.Cells[f, c].Value = "IH_TOTAL - IH_PARCIALES";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                #endregion

                List<EndesaEntity.facturacion.SIGAME_TablaBase> lista =
                    l.OrderBy(z => z.cifnif).ThenBy(z => z.id_ps).ThenBy(z => z.secfactu).ToList();

                for (int i = 0; i < lista.Count; i++)
                {
                    c = 1;
                    f++;

                    workSheet.Cells[f, c].Value = lista[i].origen; c++;
                    workSheet.Cells[f, c].Value = lista[i].marca; c++;
                    workSheet.Cells[f, c].Value = lista[i].cifnif; c++;
                    if (con_Detalle)
                    {
                        workSheet.Cells[f, c].Value = lista[i].razon_social; c++;
                        if (lista[i].direccion_punto_suministro != null)
                            workSheet.Cells[f, c].Value = lista[i].direccion_punto_suministro;
                        c++;
                    }

                    workSheet.Cells[f, c].Value = lista[i].id_ps; c++;
                    workSheet.Cells[f, c].Value = lista[i].cups22; c++;

                    workSheet.Cells[f, c].Value = lista[i].c_factura; c++;
                    workSheet.Cells[f, c].Value = lista[i].codificacion_factura; c++;
                    if (lista[i].tipologia_factura != null)
                        workSheet.Cells[f, c].Value = lista[i].tipologia_factura;
                    c++;

                    workSheet.Cells[f, c].Value = lista[i].periodo; c++;
                    workSheet.Cells[f, c].Value = lista[i].secfactu; c++;

                    workSheet.Cells[f, c].Value = lista[i].fh_emision;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    workSheet.Cells[f, c].Value = lista[i].fh_ini_facturacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = lista[i].fh_fin_facturacion;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    if (con_Detalle)
                    {
                        if (lista[i].periodofacturacion != null)
                            workSheet.Cells[f, c].Value = lista[i].periodofacturacion;
                        c++;
                    }


                    workSheet.Cells[f, c].Value = lista[i].consumo_total; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    // workSheet.Cells[f, c].Value = lista[i].importe_ih; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (lista[i].IHTI_descripcion != null)
                    {
                        if (con_Detalle)
                        {
                            workSheet.Cells[f, c].Value = lista[i].IHTI_descripcion; c++;
                        }

                        workSheet.Cells[f, c].Value = lista[i].IHTI_tipo_impositivo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        // workSheet.Cells[f, c].Value = lista[i].IHTI_consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = lista[i].IHTI_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    }
                    else
                        c = c + (con_Detalle ? 3 : 2);

                    if (lista[i].IHTD_descripcion != null)
                    {
                        if (con_Detalle)
                        {
                            workSheet.Cells[f, c].Value = lista[i].IHTD_descripcion; c++;
                        }
                        workSheet.Cells[f, c].Value = lista[i].IHTD_tipo_impositivo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        // workSheet.Cells[f, c].Value = lista[i].IHTD_consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = lista[i].IHTD_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    else
                        c = c + (con_Detalle ? 3 : 2);

                    if (lista[i].IHTC_descripcion != null)
                    {
                        if (con_Detalle)
                        {
                            workSheet.Cells[f, c].Value = lista[i].IHTC_descripcion; c++;
                        }
                        workSheet.Cells[f, c].Value = lista[i].IHTC_tipo_impositivo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                        // workSheet.Cells[f, c].Value = lista[i].IHTC_consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                        workSheet.Cells[f, c].Value = lista[i].IHTC_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    }
                    else
                        c = c + (con_Detalle ? 3 : 2);


                    workSheet.Cells[f, c].Value = lista[i].base_imponible_iva; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = lista[i].nm_importe_iva; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = lista[i].ImporteTotalFactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = lista[i].ImporteTotalFactura_sce; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = (lista[i].ImporteTotalFactura - lista[i].ImporteTotalFactura_sce);
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = (lista[i].base_imponible_iva + lista[i].nm_importe_iva) - lista[i].ImporteTotalFactura_sce;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = lista[i].uso_gas; c++;

                    // CONSUMO_TOTAL - CONSUMOS_PARCIALES

                    //workSheet.Cells[f, c].Value = lista[i].consumo_total - 
                    //    (lista[i].IHTC_consumo + lista[i].IHTD_consumo + lista[i].IHTI_consumo);
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    // IH_TOTAL - IH_PARCIALES
                    //workSheet.Cells[f, c].Value = lista[i].importe_ih -
                    //    (lista[i].IHTC_importe + lista[i].IHTD_importe + lista[i].IHTI_importe);
                    //workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";



                }

                var allCells = workSheet.Cells[1, 1, f, c];
                var cellFont = allCells.Style.Font;
                cellFont.SetFromFont(new Font("Calibri", 8));
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:AGF1"].AutoFilter = true;
                allCells.AutoFitColumns();


                workSheet = excelPackage.Workbook.Worksheets.Add("INVENTARIO");

                headerCells = workSheet.Cells[1, 1, 1, 9];
                headerFont = headerCells.Style.Font;

                f = 1;
                c = 1;

                #region Cabecera
                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "RAZON SOCIAL";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "ID PS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CUSP";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "DIRECCION PUNTO SUMINISTRO";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "PROVINCIA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "CAE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TIPO IMPOSITIVO (1,15)";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TIPO IMPOSITIVO (0,65)";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "TIPO IMPOSITIVO (0,15)";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                workSheet.Cells[f, c].Value = "EXENTO";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192)); c++;

                #endregion
                List<EndesaEntity.facturacion.SIGAME_Informe_IH> mini_list = dd.Values.ToList();
                List<EndesaEntity.facturacion.SIGAME_Informe_IH> d =
                    mini_list.OrderBy(z => z.nif).ThenBy(z => z.id_ps).ToList();


                foreach (EndesaEntity.facturacion.SIGAME_Informe_IH p in d)
                {
                    c = 1;
                    f++;
                    workSheet.Cells[f, c].Value = p.nif; c++;
                    workSheet.Cells[f, c].Value = p.cliente; c++;
                    workSheet.Cells[f, c].Value = p.id_ps; c++;
                    workSheet.Cells[f, c].Value = p.cups; c++;
                    workSheet.Cells[f, c].Value = p.direccion_punto_suministro; c++;
                    workSheet.Cells[f, c].Value = p.provincia; c++;
                    workSheet.Cells[f, c].Value = ""; c++; // CAE
                    workSheet.Cells[f, c].Value = p.ihtc ? "X" : "";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                    workSheet.Cells[f, c].Value = p.ihtd ? "X" : "";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                    workSheet.Cells[f, c].Value = p.ihti ? "X" : "";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c++;
                    workSheet.Cells[f, c].Value = p.exento ? "X" : "";
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                }

                allCells = workSheet.Cells[1, 1, f, c];
                cellFont = allCells.Style.Font;
                cellFont.SetFromFont(new Font("Calibri", 8));
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:K1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();

            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private Dictionary<int,int> CargaExentos(DateTime fd, DateTime fh)
        {
            string strSql = "";
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            Dictionary<int, int> d = new Dictionary<int, int>();
            int id_ps = 0;

            try
            {
                strSql = "SELECT DISTINCT T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE, T_SGM_M_CLIENTES.CD_CIF, T_SGM_G_PS.ID_PS,"
                    + " T_SGM_G_PS.DE_PS, T_SGM_G_PS.CD_CUPS, T_SGM_M_PUNTOS_MEDIDA.CD_UNIDAD_MEDIDA, T_SGM_P_USO_GAS.DE_USO_GAS,"
                    + " T_SGM_M_HIST_USO_GAS.NM_PORC_USO_GAS, T_SGM_M_HIST_USO_GAS.FH_INICIO, T_SGM_M_HIST_USO_GAS.FH_FIN"
                    + " FROM T_SGM_M_HIST_USO_GAS"
                    + " INNER JOIN T_SGM_G_PS ON T_SGM_M_HIST_USO_GAS.ID_PS = T_SGM_G_PS.ID_PS"
                    + " INNER JOIN T_SGM_M_CLIENTES ON T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                    + " INNER JOIN T_SGM_P_USO_GAS ON T_SGM_P_USO_GAS.CD_USO_GAS = T_SGM_M_HIST_USO_GAS.CD_USO_GAS"
                    + " LEFT JOIN T_SGM_M_PUNTOS_MEDIDA ON T_SGM_M_PUNTOS_MEDIDA.ID_PMEDIDA = T_SGM_M_HIST_USO_GAS.ID_PMEDIDA"
                    + " WHERE T_SGM_M_HIST_USO_GAS.ID_PS IN"
                    + " (SELECT DISTINCT ID_PS FROM T_SGM_G_CONTRATOS_PS WHERE FH_INICIO_REAL < '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "' AND"
                    + " (FH_FIN_REAL IS NULL OR FH_FIN_REAL >= '" + fd.ToString("yyyy-MM-dd") + "') AND ID_ESTADO_CTO NOT IN(1, 5, 7, 8))"
                    + " AND(T_SGM_M_HIST_USO_GAS.FH_FIN >= '" + fd.ToString("yyyy-MM-dd") + "' OR T_SGM_M_HIST_USO_GAS.FH_FIN IS NULL)"
                    + " AND T_SGM_M_HIST_USO_GAS.FH_INICIO <= '" + fh.ToString("yyyy-MM-dd") + "'"
                    + " AND T_SGM_P_USO_GAS.DE_USO_GAS = 'EXENTO'"
                    + " AND SUBSTRING(T_SGM_M_CLIENTES.CD_CIF, 1, 2) <> 'PT'"
                    + " ORDER BY DE_NOMBRE_CLIENTE, DE_PS";

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();
                while (r.Read())
                {
                    if(r["ID_PS"] != System.DBNull.Value)
                    {
                        id_ps = Convert.ToInt32(r["ID_PS"]);
                        int o;
                        if (!d.TryGetValue(id_ps, out o))
                            d.Add(id_ps, id_ps);
                    }

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }



    }



}
