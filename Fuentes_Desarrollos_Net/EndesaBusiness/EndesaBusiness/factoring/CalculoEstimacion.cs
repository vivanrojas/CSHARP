using EndesaBusiness.servidores;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.factoring
{
    public class CalculoEstimacion
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "mes13_prevision");

        Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic_calendario;
        public Dictionary<string, EndesaEntity.factoring.Estimacion> dic_est_ind { get; set; }
        public Dictionary<string, EndesaEntity.factoring.Estimacion> dic_est_agr { get; set; }

        List<Dictionary<string, EndesaEntity.factoring.Factura>> list_dic_ind = new List<Dictionary<string, EndesaEntity.factoring.Factura>>();
        List<Dictionary<string, EndesaEntity.factoring.Factura>> list_dic_agr = new List<Dictionary<string, EndesaEntity.factoring.Factura>>();

        contratacion.PS_AT ps; // Para puntos activos PS_AT
        contratacion.PuntosActivosBTE_COMPOR ps_bte; // Para puntos activos de BTE
        //contratacion.PuntosActivosMT_COMPOR ps_mt; // Para puntos activos de MT
        contratacion.Redshift.Inventario_PT ps_mt;
        PuntosActivosGas ps_gas; // Para puntos activos GAS
        public Lista_Negra lista_negra; // Para quitar los NIF y CUPS de la lista negra
        public Lista_Negra_Cups lista_negra_cups;

        Cal_Factorigin calendario_factoring;
        string factoring = "";

        public CalculoEstimacion(Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic)
        {
            dic_calendario = dic;

            bool firstOnly = true;            
            ps = new contratacion.PS_AT();
            ps_bte = new contratacion.PuntosActivosBTE_COMPOR();
            //ps_mt = new contratacion.PuntosActivosMT_COMPOR();
            List<string> lista_tensiones = new List<string>();
            lista_tensiones.Add("MT");
            lista_tensiones.Add("AT");
            lista_tensiones.Add("MAT");
            ps_mt = new contratacion.Redshift.Inventario_PT(null, lista_tensiones);
            ps_gas = new PuntosActivosGas();
            lista_negra = new Lista_Negra();
            lista_negra_cups = new Lista_Negra_Cups();


            foreach (KeyValuePair<string, List<EndesaEntity.factoring.CalendarioFactoring>> p in dic_calendario)
                for (int i = 0; i < p.Value.Count; i++)
                    if (firstOnly)
                    {
                        factoring = p.Value[i].factoring;
                        //[GUS 22/05/2025] modificamos la carga de los diccionarios con estimaciones existentes por la inicialización vacía de los diccionarios
                        //dic_est_ind = CargaEstimacion(p.Value[i].factoring, "INDIVIDUALES");
                        dic_est_ind = new Dictionary<string, EndesaEntity.factoring.Estimacion>();
                        //dic_est_agr = CargaEstimacion(p.Value[i].factoring, "AGRUPADAS");
                        dic_est_agr = new Dictionary<string, EndesaEntity.factoring.Estimacion>();
                        firstOnly = false;
                    }

        }



        public void CalculaEstimacion()
        {
            int num_bloques = 0;
            int contador_individuales = 0;
            string contador_individuales_texto = "";
            int contador_agrupadas = 0;
            string contador_agrupadas_texto = "";
            bool firstOnly = true;
            Dictionary<string, string> dic_referencias = new Dictionary<string, string>();
            

            try
            {


                // Siempre cargamos las facturas del bloque 0 (año anterior como base del calculo)
                foreach (KeyValuePair<string, List<EndesaEntity.factoring.CalendarioFactoring>> p in dic_calendario)
                    for (int i = 0; i < p.Value.Count; i++)
                    {
                        num_bloques++;
                        list_dic_ind.Add(CargaBloque(p.Value[i].factoring, p.Value[i].bloque, "INDIVIDUALES"));
                        list_dic_agr.Add(CargaBloque(p.Value[i].factoring, p.Value[i].bloque, "AGRUPADAS"));

                        //if (firstOnly)
                        //{
                        //    contador_individuales = UltimaReferencia(p.Value[i].factoring, "INDIVIDUALES");
                        //    contador_agrupadas = UltimaReferencia(p.Value[i].factoring, "AGRUPADAS");
                        //    firstOnly = false;
                        //}
                    }


                #region Creaccion estimacion Individuales
                // Creamos la base de estimacion individual partiendo del bloque 1
                for (int i = 1; i < num_bloques; i++)
                {
                    Console.WriteLine();
                    foreach (KeyValuePair<string, EndesaEntity.factoring.Factura> p in list_dic_ind[i])
                    {
                        //if (p.Value.cups20 == "ES0021000018587140XD")
                        //{
                        //    Console.WriteLine("Encontrado cups ES0021000018587140XD");
                        //}
                        EndesaEntity.factoring.Estimacion o;
                        if (!dic_est_ind.TryGetValue(p.Value.cups20, out o))
                        {
                            if (ExisteCUPS(p.Value.cemptitu, p.Value.cups20, ps, ps_bte, ps_mt, ps_gas) 
                                && !lista_negra.ExisteNIF(p.Value.nif))
                            {
                                contador_individuales = contador_individuales + 1;
                                contador_individuales_texto = contador_individuales.ToString().PadLeft(5, '0');
                                EndesaEntity.factoring.Estimacion c = new EndesaEntity.factoring.Estimacion();
                                c.factoring = p.Value.factoring;
                                c.tipo = p.Value.tipo;

                                if (p.Value.tiponegocio == "L")
                                    c.ln = "LUZ";
                                else
                                    c.ln = "GAS";

                                c.empresa_titular = p.Value.cemptitu;
                                c.nif = p.Value.nif;
                                c.cliente = p.Value.cliente;
                                c.ccounips = p.Value.cups13;
                                c.cupsree = p.Value.cups20;
                                c.referencia = p.Value.factoring.ToString() + "_NR" + contador_individuales_texto;

                                if (contador_individuales > 9990)
                                    Console.WriteLine(contador_individuales);

                                string refer;
                                if (!dic_referencias.TryGetValue(c.referencia, out refer))
                                    dic_referencias.Add(c.referencia, c.referencia);
                                else
                                    Console.WriteLine(c.referencia);

                                c.sec = 1;
                                c.control = 1;
                                c.ifactura[i] = p.Value.ifactura;
                                c.f[i] = p.Value.ffactura.Day;
                                //if (p.Value.flimpago > DateTime.MinValue)
                                //    c.dvto[i] = Convert.ToInt32((p.Value.flimpago - p.Value.ffactura).TotalDays);
                                c.dvto[i] = p.Value.dif;
                                c.impuestos[i] = p.Value.iva + p.Value.ise + p.Value.iimpue2 + p.Value.iimpue3 + p.Value.hidrocarburos;
                                
                                //if (p.Value.cups20 == "ES0217901000000699DN")
                                //{
                                //   Console.WriteLine("Encontrado cups");
                                //}
                                
                                dic_est_ind.Add(p.Value.cups20, c);
                                //Console.CursorLeft = 0;
                                //Console.Write(dic_est_ind.Count() + " / " + list_dic_ind[i].Count);
                            }

                        }
                        else
                        {

                            o.ifactura[i] = p.Value.ifactura;
                            o.f[i] = p.Value.ffactura.Day;
                            //if (p.Value.flimpago > DateTime.MinValue)
                            //    o.dvto[i] = Convert.ToInt32((p.Value.flimpago - p.Value.ffactura).TotalDays);
                            o.dvto[i] = p.Value.dif;
                            o.impuestos[i] = p.Value.iva + p.Value.ise + p.Value.iimpue2 + p.Value.iimpue3 + p.Value.hidrocarburos;
                            o.f[i] = p.Value.ffactura.Day;
                        }


                    }
                    // Asignamos las facturas del año anterior
                    foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> p in dic_est_ind)
                    {
                        EndesaEntity.factoring.Factura o;
                        if (list_dic_ind[0].TryGetValue(p.Value.cupsree, out o))
                        {
                            p.Value.ifactura[0] = o.ifactura;
                            p.Value.impuestos[0] = o.iva + o.ise + o.iimpue2 + o.iimpue3 + o.hidrocarburos;
                        }

                    }
                }

                GuardaEstimacion(factoring, "INDIVIDUALES", dic_est_ind);
                #endregion


                #region Creaccion estimacion Agrupadas
                // Creamos la base de estimacion agrupada partiendo del bloque 1

                for (int i = 1; i < num_bloques; i++)
                {
                    foreach (KeyValuePair<string, EndesaEntity.factoring.Factura> p in list_dic_agr[i])
                    {
                        EndesaEntity.factoring.Estimacion o;
                        if (!dic_est_agr.TryGetValue(p.Value.nif, out o))
                        {

                            if ((ps.ExisteCIF(p.Value.nif) ||                                
                                ps_mt.ExisteCIF(p.Value.nif) ||
                                ps_gas.ExisteCIF(p.Value.nif)) && !lista_negra.ExisteNIF(p.Value.nif))
                            {

                                contador_agrupadas = contador_agrupadas + 1;
                                contador_agrupadas_texto = contador_agrupadas.ToString().PadLeft(5, '0');

                                

                                EndesaEntity.factoring.Estimacion c = new EndesaEntity.factoring.Estimacion();
                                c.factoring = p.Value.factoring;
                                c.tipo = p.Value.tipo;
                                if (p.Value.tiponegocio == "L")
                                    c.ln = "LUZ";
                                else
                                    c.ln = "GAS";
                                c.empresa_titular = p.Value.cemptitu;
                                c.nif = p.Value.nif;
                                c.cliente = p.Value.cliente;
                                c.ccounips = p.Value.cups13;
                                c.cupsree = p.Value.cups20;
                                c.referencia = p.Value.factoring.ToString() + "_AG" + contador_agrupadas_texto;
                                c.sec = 1;
                                c.control = 1;
                                c.ifactura[i] = p.Value.ifactura;
                                c.f[i] = p.Value.ffactura.Day;
                                //if (p.Value.flimpago > DateTime.MinValue)
                                //    c.dvto[i] = Convert.ToInt32((p.Value.flimpago - p.Value.ffactura).TotalDays);
                                c.dvto[i] = p.Value.dif;
                                c.impuestos[i] = p.Value.iva + p.Value.ise + p.Value.iimpue2 + p.Value.iimpue3 + p.Value.hidrocarburos;
                                dic_est_agr.Add(p.Value.nif, c);
                            }

                        }
                        else
                        {
                            contador_agrupadas = contador_agrupadas + 1;
                            contador_agrupadas_texto = contador_agrupadas.ToString().PadLeft(5, '0');                            
                            o.ifactura[i] = p.Value.ifactura;
                            o.f[i] = p.Value.ffactura.Day;
                            //if (p.Value.flimpago > DateTime.MinValue)
                            //    o.dvto[i] = Convert.ToInt32((p.Value.flimpago - p.Value.ffactura).TotalDays);
                            o.dvto[i] = p.Value.dif;
                            o.impuestos[i] = p.Value.iva + p.Value.ise + p.Value.iimpue2 + p.Value.iimpue3 + p.Value.hidrocarburos;
                            o.f[i] = p.Value.ffactura.Day;
                        }

                    }

                    // Asignamos las facturas del año anterior
                    foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> p in dic_est_agr)
                    {
                        EndesaEntity.factoring.Factura o;
                        if (list_dic_agr[0].TryGetValue(p.Value.nif, out o))
                        {
                            p.Value.ifactura[0] = o.ifactura;
                            p.Value.impuestos[0] = o.iva + o.ise + o.iimpue2 + o.iimpue3 + o.hidrocarburos;
                        }

                    }
                }

                GuardaEstimacion(factoring, "AGRUPADAS", dic_est_agr);
                #endregion
            }
            catch (Exception e)
            {
                ficheroLog.addError("CalculoEstimacion.CalculaEstimacion: " + e.Message);
            }

        }

        private Dictionary<string, EndesaEntity.factoring.Factura> CargaBloque(string factoring, int bloque, string tipo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, EndesaEntity.factoring.Factura> dic = new Dictionary<string, EndesaEntity.factoring.Factura>();

            try
            {

                Console.WriteLine("Cargando facturas para Factoring " + factoring + " bloque " + bloque + " y tipo " + tipo);
                strSql = "select f.factoring, f.bloque, f.CCOUNIPS, f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,";

                if (tipo == "AGRUPADAS")
                    strSql += "SUM(f.IFACTURA) as IFACTURA,  SUM(f.IVA) as IVA, SUM(f.IIMPUES2) AS IIMPUES2,"
                        + " SUM(f.IIMPUES3) AS IIMPUES3, SUM(f.IBASEISE) as IBASEISE, SUM(f.ISE) as ISE,"
                        + " ROUND(AVG(DATEDIFF(if (s.flimpago is null or s.flimpago = '0000-00-00',if (o.FLIMPAGO is null or o.FLIMPAGO = '0000-00-00', d.FLIMPAGO, o.FLIMPAGO) , s.flimpago),f.FFACTURA)),0) AS DIF,";
                else
                    strSql += " f.IFACTURA, f.IVA, f.IIMPUES2, f.IIMPUES3, f.IBASEISE, f.ISE,"
                        + " DATEDIFF(if (s.flimpago is null or s.flimpago = '0000-00-00',if (o.FLIMPAGO is null or o.FLIMPAGO = '0000-00-00', d.FLIMPAGO, o.FLIMPAGO) , s.flimpago),f.FFACTURA) AS DIF,";

                strSql += " f.TFACTURA, f.TESTFACT, f.DAPERSOC, f.CNIFDNIC,"
                    + " f.INDEMPRE, f.CUPSREE, f.CFACTREC, f.CLINNEG, f.CSEGMERC, f.NUMLABOR, f.TIPONEGOCIO, f.COMENTARIOS, f.HIDROCARBUROS,"
                    + " if (s.flimpago is null or s.flimpago = '0000-00-00',if (o.FLIMPAGO is null or o.FLIMPAGO = '0000-00-00', d.FLIMPAGO, o.FLIMPAGO) , s.flimpago) as FLIMPAGO"
                    + " from fact.ff_facturas_all f"
                    + " left outer join 13_obligaciones_sap s on s.id_fact = f.CFACTURA"
                    + " left outer join deuda_obligaciones_original d on"
                    + " d.CEMPTITU = f.CEMPTITU and"
                    + " d.CREFEREN = f.CREFEREN and"
                    + " d.SECFACTU = f.SECFACTU"
                    + " left outer join 13_oblig_cob o on"
                    + " o.CEMPTITU = f.CEMPTITU and"
                    + " o.CREFEREN = f.CREFEREN and"
                    + " o.SECFACTU = f.SECFACTU"
                + " where"
                    + " factoring = " + factoring + " and"
                    + " bloque = " + bloque + " and"
                    + " tipo = '" + tipo + "'";

                if (tipo == "INDIVIDUALES")
                    strSql += " order by CUPSREE, FFACTDES DESC";
                else
                    strSql += " group by CNIFDNIC order by CNIFDNIC";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.factoring.Factura c = new EndesaEntity.factoring.Factura();
                    c.factoring = Convert.ToInt32(r["factoring"]);
                    c.tipo = tipo;

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.cups13 = r["CCOUNIPS"].ToString();

                    if (r["CEMPTITU"] != System.DBNull.Value)
                        c.cemptitu = Convert.ToInt32(r["CEMPTITU"]);

                    if (r["CREFEREN"] != System.DBNull.Value)
                        c.creferen = Convert.ToInt64(r["CREFEREN"]);

                    if (r["SECFACTU"] != System.DBNull.Value)
                        c.secfactu = Convert.ToInt32(r["SECFACTU"]);

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    if (r["IVA"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["IVA"]);

                    if (r["IIMPUES2"] != System.DBNull.Value)
                        c.iimpue2 = Convert.ToDouble(r["IIMPUES2"]);

                    if (r["IIMPUES3"] != System.DBNull.Value)
                        c.iimpue3 = Convert.ToDouble(r["IIMPUES3"]);

                    if (r["IBASEISE"] != System.DBNull.Value)
                        c.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                    if (r["ISE"] != System.DBNull.Value)
                        c.ise = Convert.ToDouble(r["ISE"]);

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.cliente = r["DAPERSOC"].ToString();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.nif = r["CNIFDNIC"].ToString();
                    
                    //if (c.nif == "Q2802152E")
                    //    Console.WriteLine("Encontrado NIF ADIF: Q2802152E");

                    if (r["INDEMPRE"] != System.DBNull.Value)
                        c.indempre = r["INDEMPRE"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value )
                        c.cups20 = r["CUPSREE"].ToString().Length > 20 ?
                            r["CUPSREE"].ToString().Substring(0, 20) : r["CUPSREE"].ToString();
                    else
                            c.cups20 = c.cups13;
                    
                    //if (tipo == "AGRUPADAS" && c.cups20 == "")
                    //    c.cups20 = c.cfactura.ToString().Length > 20 ? c.cfactura.ToString().Substring(0, 20) : c.cfactura.ToString();
                    

                    if (r["CFACTREC"] != System.DBNull.Value)
                        c.cfactrec = r["CFACTREC"].ToString();

                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.tiponegocio = r["TIPONEGOCIO"].ToString();


                    if (r["COMENTARIOS"] != System.DBNull.Value)
                        c.comentarios = r["COMENTARIOS"].ToString();

                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        if (r["FLIMPAGO"].ToString() != "0000-00-00")
                            c.flimpago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["HIDROCARBUROS"] != System.DBNull.Value)
                        c.hidrocarburos = Convert.ToDouble(r["HIDROCARBUROS"]);

                    if (r["DIF"] != System.DBNull.Value)
                        c.dif = Convert.ToInt32(r["DIF"]);

                    // En caso de cups20 nulo en individuales o nif nulo en individuales y agrupadas no guardamos en diccionario y descartamos (se debería revisar porque faltan estos datos en la base de datos)
                    if ( ((tipo == "INDIVIDUALES" && c.cups20 != null && c.cups20 != "") || (tipo=="AGRUPADAS")) &&  c.nif != null && c.nif != "") 
                    {
                    
                        if (tipo == "INDIVIDUALES" )
                        {
                            if (!lista_negra_cups.ExisteCUPS(c.cups20))
                            {
                                EndesaEntity.factoring.Factura o;
                                if (!dic.TryGetValue(c.cups20, out o))
                                    dic.Add(c.cups20, c);
                            }
                        }
                        else if (tipo == "AGRUPADAS")
                            dic.Add(c.nif, c);
                        
                    }
                }
                db.CloseConnection();
                return dic;
            }
            catch (Exception e)
            {
               
                ficheroLog.addError("CalculoEstimacion.CargaBloque: " + e.Message);
                return null;

            }

        }

        private Dictionary<string, EndesaEntity.factoring.Estimacion> CargaEstimacion(string factoring, string tipo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, EndesaEntity.factoring.Estimacion> dic = new Dictionary<string, EndesaEntity.factoring.Estimacion>();

            Console.WriteLine("Cargando datos estimados para Factoring " + factoring + " " + tipo);
            strSql = "select factoring, tipo, ln, empresa_titular, nif, cliente, ccounips, cupsree, referencia, sec, control,"
                + " estimacion_importe, estimacion_base, estimacion_impuestos,"
                + " diaf, diav, tam, ifactura_0, iimpuesto_0, ifactura_1, iimpuesto_1, ifactura_2, iimpuesto_2,"
                + " ifactura_3, iimpuesto_3, f_1, f_2, f_3, vto_1, vto_2, vto_3"
                + " from ff_estimacion where"
                + " factoring = '" + factoring + "' and"
                + " tipo = '" + tipo + "'"
                + " order by referencia";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.factoring.Estimacion c = new EndesaEntity.factoring.Estimacion();

                c.factoring = Convert.ToInt32(r["factoring"]);
                c.tipo = r["tipo"].ToString();

                if (r["ln"] != System.DBNull.Value)
                    c.ln = r["ln"].ToString();

                if (r["empresa_titular"] != System.DBNull.Value)
                    c.empresa_titular = Convert.ToInt32(r["empresa_titular"]);

                if (r["nif"] != System.DBNull.Value)
                    c.nif = r["nif"].ToString();

                if (r["cliente"] != System.DBNull.Value)
                    c.cliente = r["cliente"].ToString();

                if (r["ccounips"] != System.DBNull.Value)
                    c.ccounips = r["ccounips"].ToString();

                if (r["cupsree"] != System.DBNull.Value)
                    c.cupsree = r["cupsree"].ToString().Length > 20 ?
                        r["cupsree"].ToString().Substring(0, 20) : r["cupsree"].ToString();

                if (r["referencia"] != System.DBNull.Value)
                    c.referencia = r["referencia"].ToString();

                if (r["sec"] != System.DBNull.Value)
                    c.sec = Convert.ToInt32(r["sec"]);

                if (r["control"] != System.DBNull.Value)
                    c.control = Convert.ToInt32(r["control"]);

                if (r["estimacion_importe"] != System.DBNull.Value)
                    c.estimacion_importe = Convert.ToDouble(r["estimacion_importe"]);

                if (r["estimacion_base"] != System.DBNull.Value)
                    c.estimacion_base = Convert.ToDouble(r["estimacion_base"]);

                if (r["estimacion_impuestos"] != System.DBNull.Value)
                    c.estimacion_impuestos = Convert.ToDouble(r["estimacion_impuestos"]);

                if (r["tam"] != System.DBNull.Value)
                    c.tam = Convert.ToDouble(r["tam"]);

                if (r["ifactura_0"] != System.DBNull.Value)
                    c.ifactura[0] = Convert.ToDouble(r["ifactura_0"]);

                if (r["iimpuesto_0"] != System.DBNull.Value)
                    c.impuestos[0] = Convert.ToDouble(r["iimpuesto_0"]);

                if (r["ifactura_1"] != System.DBNull.Value)
                    c.ifactura[1] = Convert.ToDouble(r["ifactura_1"]);

                if (r["iimpuesto_1"] != System.DBNull.Value)
                    c.impuestos[1] = Convert.ToDouble(r["iimpuesto_1"]);

                if (r["ifactura_2"] != System.DBNull.Value)
                    c.ifactura[2] = Convert.ToDouble(r["ifactura_2"]);

                if (r["iimpuesto_2"] != System.DBNull.Value)
                    c.impuestos[2] = Convert.ToDouble(r["iimpuesto_2"]);

                if (r["ifactura_3"] != System.DBNull.Value)
                    c.ifactura[3] = Convert.ToDouble(r["ifactura_3"]);

                if (r["iimpuesto_3"] != System.DBNull.Value)
                    c.impuestos[3] = Convert.ToDouble(r["iimpuesto_3"]);

                if (r["f_1"] != System.DBNull.Value)
                    c.f[1] = Convert.ToInt32(r["f_1"]);

                if (r["f_2"] != System.DBNull.Value)
                    c.f[2] = Convert.ToInt32(r["f_2"]);

                if (r["f_3"] != System.DBNull.Value)
                    c.f[3] = Convert.ToInt32(r["f_3"]);

                if (r["vto_1"] != System.DBNull.Value)
                    c.dvto[1] = Convert.ToInt32(r["vto_1"]);

                if (r["vto_2"] != System.DBNull.Value)
                    c.dvto[2] = Convert.ToInt32(r["vto_2"]);

                if (r["vto_3"] != System.DBNull.Value)
                    c.dvto[3] = Convert.ToInt32(r["vto_3"]);

                if (tipo == "INDIVIDUALES")
                    dic.Add(c.cupsree, c);
                else
                    dic.Add(c.nif, c);

            }
            db.CloseConnection();
            return dic;
        }

        private void GuardaEstimacion(string factoring, string tipo, Dictionary<string, EndesaEntity.factoring.Estimacion> dic)
        {
            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            bool firstOnly = true;
            int id = 0;
            int total_facturas = 0;

            try
            {
                foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> p in dic)
                {
                    total_facturas++;
                    id++;
                    if (firstOnly)
                    {
                        sb.Append("replace into ff_estimacion (factoring, tipo, ln, empresa_titular, nif, cliente,");
                        sb.Append(" ccounips, cupsree, referencia, sec, control, estimacion_importe,");
                        sb.Append(" estimacion_base, estimacion_impuestos,");
                        sb.Append(" diaf, diav, tam,");
                        sb.Append(" ifactura_0, ifactura_1, ifactura_2, ifactura_3,");
                        sb.Append(" iimpuesto_0, iimpuesto_1, iimpuesto_2, iimpuesto_3,");
                        sb.Append(" f_1, f_2, f_3, vto_1, vto_2, vto_3) values");
                        firstOnly = false;
                    }

                    sb.Append(" (").Append(factoring).Append(",");
                    sb.Append("'").Append(tipo).Append("',");
                    sb.Append("'").Append(p.Value.ln).Append("',");

                    sb.Append("'").Append(p.Value.empresa_titular).Append("',");
                    sb.Append("'").Append(p.Value.nif).Append("',");
                    sb.Append("'").Append(p.Value.cliente).Append("',");

                    if (p.Value.ccounips != null)
                        sb.Append("'").Append(p.Value.ccounips).Append("',");
                    else
                        sb.Append("null,");

                    if (p.Value.cupsree != null)
                        sb.Append("'").Append(p.Value.cupsree).Append("',");
                    else
                        sb.Append("null,");

                    if (p.Value.referencia != null)
                        sb.Append("'").Append(p.Value.referencia).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append(p.Value.sec).Append(",");

                    sb.Append(p.Value.control).Append(",");

                    if (p.Value.estimacion_importe > 0)
                        sb.Append(p.Value.estimacion_importe.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.estimacion_base > 0)
                        sb.Append(p.Value.estimacion_base.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.estimacion_impuestos > 0)
                        sb.Append(p.Value.estimacion_impuestos.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");


                    sb.Append("null,null,");

                    if (p.Value.tam > 0)
                        sb.Append(p.Value.tam.ToString().Replace(",", "."));
                    else
                        sb.Append("null");

                    for (int i = 0; i < 4; i++)
                    {
                        if (p.Value.ifactura[i] > 0)
                            sb.Append(",").Append(p.Value.ifactura[i].ToString().Replace(",", "."));
                        else
                            sb.Append(",null");
                    }

                    for (int i = 0; i < 4; i++)
                    {
                        if (p.Value.impuestos[i] > 0)
                            sb.Append(",").Append(p.Value.impuestos[i].ToString().Replace(",", "."));
                        else
                            sb.Append(",null");
                    }

                    for (int i = 1; i < 4; i++)
                        sb.Append(",").Append(p.Value.f[i]);
                    for (int i = 1; i < 4; i++)
                        sb.Append(",").Append(p.Value.dvto[i]);

                    sb.Append("),");


                    if (id == 250)
                    {
                        Console.WriteLine("Guardando " + total_facturas + " de " + dic.Count());
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        id = 0;
                    }
                }

                if (id > 0)
                {
                    firstOnly = true;
                    Console.WriteLine("Guardando " + total_facturas + " de " + dic.Count());
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    id = 0;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void GuardaImporteHidrocarburo(int factoring, int bloque, string tipo)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            strSql = "update ff_facturas_all set hidrocarburos = (select sum(t.ICONFAC) from fo_tcon t where"
                + " t.CREFEREN = ff_facturas_all.CREFEREN and"
                + " t.SECFACTU = ff_facturas_all.SECFACTU and"
                + " t.TESTFACT = ff_facturas_all.TESTFACT and"
                + " t.TCONFAC in (1100, 1101, 1102, 1103, 1104, 1105, 1106, 1107, 1504, 1505, 1512, 1633, 1634))"
                + " where factoring = " + factoring
                + " and bloque = " + bloque
                + " and tipo = '" + tipo + "'";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }



        //private bool ExisteCUPS(int empresa, string cups20, contratacion.PS_AT ps,
        //    contratacion.PuntosActivosBTE_COMPOR ps_bte, contratacion.PuntosActivosMT_COMPOR ps_mt,
        //    PuntosActivosGas ps_gas)
        private bool ExisteCUPS(int empresa, string cups20, contratacion.PS_AT ps,
            contratacion.PuntosActivosBTE_COMPOR ps_bte, contratacion.Redshift.Inventario_PT ps_mt,
            PuntosActivosGas ps_gas)
        {
            bool existe = false;
            string cups = "";
            string cisterna = "";            

            if (cups20.Length >= 20)
                cups = cups20.Substring(0, 20);
            else
                cisterna = cups20;

            if (cups20.Length >= 20)
            {
                if (ps.ExisteAlta(cups) || empresa == 6)
                    existe = true;
                //else if (ps_mt.ExisteAlta(cups))
                else if (ps_mt.ExisteCUPS(cups))
                    existe = true;
                else if (ps_gas.ExisteAlta(cups.ToUpper()))
                    existe = true;
            }
            else
                existe = ps_gas.ExisteAlta(cisterna);
            //existe = true;

            return existe;
        }

        private int UltimaReferencia(string factoring, string tipo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int refe = 0;


            strSql = "SELECT MAX(referencia) as ultima FROM ff_estimacion"
                + " WHERE factoring = " + factoring + " AND"
                + " tipo = '" + tipo + "'";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["ultima"] != System.DBNull.Value)
                    if (tipo == "INDIVIDUALES")
                        refe = Convert.ToInt32(r["ultima"].ToString().Replace(factoring + "_NR", ""));
                    else
                        refe = Convert.ToInt32(r["ultima"].ToString().Replace(factoring + "_AG", ""));
            }


            db.CloseConnection();
            return refe;
        }
    }
}
