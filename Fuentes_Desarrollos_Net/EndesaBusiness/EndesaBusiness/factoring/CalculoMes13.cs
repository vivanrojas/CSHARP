using EndesaBusiness.facturacion;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.factoring
{
    public  class CalculoMes13
    {

        public double importe_factura { get; set; }
        public double importe_nif { get; set; }
        public int num_facturas_proceso { get; set; }
        public int num_facturas_total { get; set; }
        public double total_importe_facturas { get; set; }
        public double total_importe_facturas_proceso { get; set; }

        private DateTime ffd { get; set; } // Fecha Factura desde
        private DateTime ffh { get; set; } // Fecha Factura hasta
        private DateTime fcd { get; set; } // Fecha consumo desde
        private DateTime fch { get; set; } // Fecha consumo hasta



        utilidades.Param param;


        public CalculoMes13(DateTime fecha_factura_desde, DateTime fecha_factura_hasta,
            DateTime fecha_consumo_desde, DateTime fecha_consumo_hasta,
            List<string> lista_negocio, List<string> lista_empresa_titular, List<string> lista_tipo_factura,
            double imp_factura, double imp_nif,
            bool excluir_adm, bool excluir_cisternas, bool excluir_NL)
        {

            AnexionDatos(fecha_factura_desde, fecha_factura_hasta, fecha_consumo_desde, fecha_consumo_hasta, lista_tipo_factura);

            param = new utilidades.Param("13_param", MySQLDB.Esquemas.FAC);
            importe_factura = imp_factura;
            importe_nif = imp_nif;

            LanzaProceso();

            num_facturas_proceso = Calcula_num_facturas_proceso();
            total_importe_facturas_proceso = Calcula_importe_facturas_proceso();
        }

        public CalculoMes13()
        {
            param = new utilidades.Param("13_param", MySQLDB.Esquemas.FAC);
            importe_factura = Convert.ToDouble(param.GetValue("importe_facturas", DateTime.Now, DateTime.Now));
            importe_nif = Convert.ToDouble(param.GetValue("importe_nif", DateTime.Now, DateTime.Now));
            num_facturas_proceso = Calcula_num_facturas_proceso();
            total_importe_facturas_proceso = Calcula_importe_facturas_proceso();
        }

        private int Calcula_num_facturas_proceso()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int numFacturas = 0;

            try
            {
                strSql = "select count(*) num_facturas from 13_facturas;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["num_facturas"] != System.DBNull.Value)
                        numFacturas = Convert.ToInt32(r["num_facturas"]);
                }
                db.CloseConnection();
                return numFacturas;

            }
            catch (Exception e)
            {
                return 0;
            }
        }

        private double Calcula_importe_facturas_proceso()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            double total = 0;
            try
            {
                strSql = "select sum(ifactura) ifactura from 13_facturas;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["ifactura"] != System.DBNull.Value)
                        total = Convert.ToDouble(r["ifactura"]);
                }
                db.CloseConnection();
                return total;

            }
            catch (Exception e)
            {
                return 0;
            }
        }

        private void LanzaProceso()
        {

            // MarcaCisternas();
            // BorradoFacturas("A");
            // BorradoFacturas("S");
            //Mes13_AgruparFacturas_PerdiodosPartidos();            
            //Quita_Grupo_Menor();
            Mes13_Excel_Consumos();
        }



        public void MarcaCisternas()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                strSql = "update 13_facturas f"
                    + " left outer join fo_s s on"
                    + " s.CREFEREN = f.CREFEREN and"
                    + " s.SECFACTU = f.SECFACTU"
                    + " left outer join cm_inventario_gas i on"
                    + " i.ID_PS = s.ID_PS and"
                    + " (i.FINICIO <= f.FFACTDES and(i.FFIN >= f.FFACTHAS or i.FFIN is null)) and"
                    + " i.Grupo = 'Cisterna'"
                    + " set f.CUPSREE = concat('CISTERNA_', i.ID_PS)"
                    + " where"
                    + " f.TIPONEGOCIO = 'G' and"
                    + " (f.CUPSREE is null or trim(f.CUPSREE = ''));";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void AnexionDatos(DateTime f_factura_des, DateTime f_factura_has, DateTime ffactdes, DateTime ffacthas, List<string> lista_tipos_factura)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            try
            {

                facturacion.TiposFactura tf = new TiposFactura();

                // Borramos la tabla de facturas                
                strSql = "delete from 13_facturas;";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Facturas emitidas en "Cambiar por el mes que corresponda"

                strSql = "replace into 13_facturas select * from fo f where"
                    + " (f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                    + " and f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "')"
                    + " and (f.FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "'"
                    + " and f.FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "')";

                for (int i = 0; i < lista_tipos_factura.Count(); i++)
                {
                    if (firstOnly)
                    {
                        strSql += " and f.TFACTURA in (" + tf.GetIDFromDescription(lista_tipos_factura[i]);
                        firstOnly = false;
                    }
                    else
                        strSql += " ," + tf.GetIDFromDescription(lista_tipos_factura[i]);
                }

                if (lista_tipos_factura.Count() > 0)
                    strSql += ");";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                // Adjuntamos facturas de GAS
                Console.WriteLine("Adjuntamos facturas de GAS");
                strSql = "replace into 13_facturas select * from fo f"
                 + " where f.CEMPTITU <> 70 and"
                 + " f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and"
                 + " f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "' and"
                 + " f.TIPONEGOCIO = 'G';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Cambiamos el tipo de las facturas de GAS por tipo 1
                Console.WriteLine("Cambiamos el tipo de las facturas de GAS por tipo 1");
                strSql = "update 13_facturas set tfactura = 1 where TIPONEGOCIO = 'G';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de consumos del mismo mes de emisión de facturas
                Console.WriteLine("Quitamos las facturas de consumos del mismo mes de emisión de facturas");
                strSql = "delete from 13_facturas where FFACTDES >= '" + f_factura_des.ToString("yyyy-MM-dd") + "';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de EEXXI
                Console.WriteLine("Quitamos las facturas de EEXXI");
                strSql = "delete from 13_facturas where CEMPTITU = 70;";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                TrataHolanda(f_factura_des, f_factura_has);

                // Quitamos las facturas de BTN
                Console.WriteLine("Quitamos las facturas de BTN");
                strSql = "delete from 13_facturas where CEMPTITU = 20 and INDEMPRE = 'CEFACO';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de BTE
                Console.WriteLine("Quitamos las facturas de BTE");
                strSql = "delete from 13_facturas where CEMPTITU = 80 and INDEMPRE = 'CEFACO';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Borramos las facturas de consumos que no correspondan con las fechas que queramos
                Console.WriteLine("Borramos las facturas de consumos que no correspondan con las fechas que queramos");
                strSql = "delete from 13_facturas where ffacthas < '" + ffactdes.ToString("yyyy-MM-dd") + "';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Borramos las facturas individuales inferiores a 5.000 €
                // strSql = "delete from 13_facturas where IFACTURA < 5000;";

                // Para quedarnos con las últimas facturas emitidas lo vamos a hacer de la siguiente manera:
                // 1.- Guardamos codigos de refacturaciones que vamos a borrar.
                //Console.WriteLine("Para quedarnos con las últimas facturas emitidas lo vamos a hacer de la siguiente manera:");
                //Console.WriteLine("Guardamos codigos de refacturaciones que vamos a borrar.");
                //Console.WriteLine("");
                //strSql = "delete from 13_refacturaciones;";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();
                //strSql = "replace into 13_refacturaciones select f.CFACTREC from 13_facturas f where (f.CFACTREC is not null and f.CFACTREC <> '');";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "delete f"
                //    + " from 13_facturas f"
                //    + " inner join 13_refacturaciones f1 on"
                //    + " f.CFACTURA = f1.CFACTREC;";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //// Borramos las facturas Anuladas(A) y Abonos (S)
                //strSql = "delete from 13_facturas where testfact in ('A', 'S');";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();


                // FIN
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        private List<EndesaEntity.facturacion.mes13.Factura> Mes13_Total_Facturas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Console.WriteLine("Consultando facturas");
            List<EndesaEntity.facturacion.mes13.Factura> lista = new List<EndesaEntity.facturacion.mes13.Factura>();
            try
            {
                strSql = "select if (f.TIPONEGOCIO = 'L', 'LUZ', 'GAS') as TIPONEGOCIO,"
                    + " f.CEMPTITU as EMPRESATITULAR ,f.DAPERSOC as CLIENTE, f.CNIFDNIC as NIF,"
                    + " f.CCOUNIPS, f.CUPSREE,"
                    + " cea.FA, cea.FprevB, cea.Fecha_Baja,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " if (o.FLIMPAGO is null, d.FLIMPAGO,o.FLIMPAGO) as FLIMPAGO,"
                    + " f.CFACTURA, f.IFACTURA, t.TAM,"
                    + " null as VNUMPLAZ,"
                    + " s.tipoEmpresa as TIPO_EMPRESA,"
                    + " s.ambito as AMBITO, s.territorial as TERRITORIAL,"
                    + " concat(s.nombreGestor, ' ', s.Apellido1Gestor, ' ', s.Apellido2Gestor) as NOMBRE_GESTOR,"
                    + " f.TFACTURA, if(f.TIPONEGOCIO = 'L' AND f.CEMPTITU = 20 AND (f.CCOUNIPS is not null and f.CCOUNIPS <> '') AND cea.IDU is null, 'BAJA', null) as COMENTARIO"
                    // TEMPORAL
                    + " , if(f.TFACTURA <> 6,  xx.FFACTURA, ag.FFACTURA) as FFACTURA_AUX"
                    + " , if(f.TFACTURA <> 6,  xx.FFACTDES, ag.FFACTDES) as FFACTDES_AUX"
                    + " , if(f.TFACTURA <> 6,  xx.FFACTHAS, ag.FFACTHAS) as FFACTHAS_AUX"
                    + " , if(f.TFACTURA <> 6,  xx.IFACTURA, ag.IFACTURA) as IFACTURA_AUX"
                    // FIN TEMPORAL
                    + " from fact.13_facturas f"
                    + " left outer join cobr.carteraSIOC s on"
                    + " s.cif = f.CNIFDNIC"
                    + " left outer join fact.tam t on"
                    + " t.CUPS13 = f.CCOUNIPS"
                    + " left outer join fact.13_obligaciones_temp d on"
                    + " d.CEMPTITU = f.CEMPTITU and"
                    + " d.CREFEREN = lpad(f.CREFEREN, 13, 0) and"
                    + " d.SECFACTU = f.SECFACTU and"
                    + " d.FLIMPAGO <> '00000000'"
                    + " left outer join fact.13_oblig_cob o on"
                    + " o.CEMPTITU = f.CEMPTITU and"
                    + " o.CREFEREN = f.CREFEREN and"
                    + " o.SECFACTU = f.SECFACTU"
                    + " left outer join med.scea cea on"
                    + " cea.IDU = f.CCOUNIPS"
                    // SOLO PARA ESTA VEZ
                    // SACAR EL IMPORTE DE LAS FACTURAS QUE SALIERON EN 2018-01-01 CON CONSUMOS DE 2017-12-31

                    + " left outer join fact.13_facturas_aux xx on"
                    + " xx.CEMPTITU = f.CEMPTITU and"
                    + " xx.CNIFDNIC = f.CNIFDNIC and"
                    + " xx.CUPSREE = f.CUPSREE"
                    + " left outer join fact.13_facturas_aux_agrupadas ag on"
                    + " ag.CEMPTITU = f.CEMPTITU and"
                    + " ag.CNIFDNIC = f.CNIFDNIC "


                    //FIN SOLO PARA ESTA VEZ

                    + " group by f.CREFEREN, f.SECFACTU ORDER BY f.CNIFDNIC;";

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Console.Write(".");

                    EndesaEntity.facturacion.mes13.Factura c = new EndesaEntity.facturacion.mes13.Factura();
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.ln = r["TIPONEGOCIO"].ToString();

                    if (r["EMPRESATITULAR"] != System.DBNull.Value)
                        c.empresaTitular = r["EMPRESATITULAR"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["FA"] != System.DBNull.Value)
                        c.fa = Convert.ToDateTime(r["FA"]);

                    if (r["FprevB"] != System.DBNull.Value)
                        c.fprevb = Convert.ToDateTime(r["FprevB"]);

                    if (r["Fecha_Baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["Fecha_Baja"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (c.nif.Substring(0, 2) == "NL")
                    {
                        c.ffactdes = fcd;
                        c.ffacthas = fch;
                        c.ffactdes_aux = new DateTime(2017, 12, 01);
                        c.ffacthas_aux = new DateTime(2017, 12, 31);
                    }
                    else
                    {
                        if (r["FFACTDES"] != System.DBNull.Value)
                            c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                        if (r["FFACTHAS"] != System.DBNull.Value)
                            c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);
                    }


                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        c.flimpago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"].ToString());

                    if (r["TAM"] != System.DBNull.Value)
                        c.tam = Convert.ToDouble(r["TAM"].ToString());

                    if (r["VNUMPLAZ"] != System.DBNull.Value)
                        c.vnumplaz = r["VNUMPLAZ"].ToString();

                    if (r["TIPO_EMPRESA"] != System.DBNull.Value)
                        c.tipoEmpresa = r["TIPO_EMPRESA"].ToString();

                    if (r["AMBITO"] != System.DBNull.Value)
                        c.territorial = r["AMBITO"].ToString();

                    if (r["NOMBRE_GESTOR"] != System.DBNull.Value)
                        c.gestor = r["NOMBRE_GESTOR"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    // TEMPORAL

                    if (r["FFACTURA_AUX"] != System.DBNull.Value)
                        c.ffactura_aux = Convert.ToDateTime(r["FFACTURA_AUX"]);

                    if (r["FFACTDES_AUX"] != System.DBNull.Value)
                        c.ffactdes_aux = Convert.ToDateTime(r["FFACTDES_AUX"]);

                    if (r["FFACTHAS_AUX"] != System.DBNull.Value)
                        c.ffacthas_aux = Convert.ToDateTime(r["FFACTHAS_AUX"]);

                    if (r["IFACTURA_AUX"] != System.DBNull.Value)
                        c.ifactura_aux = Convert.ToDouble(r["IFACTURA_AUX"]);

                    // TEMPORAL

                    lista.Add(c);
                }
                db.CloseConnection();
                Console.WriteLine();

                return lista;
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
                return null;
            }

        }
        private List<EndesaEntity.facturacion.mes13.Factura> Mes13_Total_Facturas_SinAgrupadas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Console.WriteLine("Consultando facturas");
            List<EndesaEntity.facturacion.mes13.Factura> lista = new List<EndesaEntity.facturacion.mes13.Factura>();
            try
            {
                strSql = "select if (f.TIPONEGOCIO = 'L', 'LUZ', 'GAS') as TIPONEGOCIO,"
                    + " f.CEMPTITU as EMPRESATITULAR ,f.DAPERSOC as CLIENTE, f.CNIFDNIC as NIF,"
                    + " f.CCOUNIPS, f.CUPSREE,"
                    + " cea.FA, cea.FprevB, cea.Fecha_Baja,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " if (o.FLIMPAGO is null, d.FLIMPAGO,o.FLIMPAGO) as FLIMPAGO,"
                    + " f.CFACTURA, f.IFACTURA, t.TAM,"
                    + " null as VNUMPLAZ,"
                    + " s.tipoEmpresa as TIPO_EMPRESA,"
                    + " s.ambito as AMBITO, s.territorial as TERRITORIAL,"
                    + " concat(s.nombreGestor, ' ', s.Apellido1Gestor, ' ', s.Apellido2Gestor) as NOMBRE_GESTOR,"
                    + " f.TFACTURA, if(f.TIPONEGOCIO = 'L' AND f.CEMPTITU = 20 AND (f.CCOUNIPS is not null and f.CCOUNIPS <> '') AND cea.IDU is null, 'BAJA', null) as COMENTARIO"
                    + " from fact.13_facturas f"
                    + " left outer join cobr.carteraSIOC s on"
                    + " s.cif = f.CNIFDNIC"
                    + " left outer join fact.tam t on"
                    + " t.CUPS13 = f.CCOUNIPS"
                    + " left outer join fact.13_obligaciones_temp d on"
                    + " d.CEMPTITU = f.CEMPTITU and"
                    + " d.CREFEREN = lpad(f.CREFEREN, 13, 0) and"
                    + " d.SECFACTU = f.SECFACTU and"
                    + " d.FLIMPAGO <> '00000000'"
                    + " left outer join fact.13_oblig_cob o on"
                    + " o.CEMPTITU = f.CEMPTITU and"
                    + " o.CREFEREN = f.CREFEREN and"
                    + " o.SECFACTU = f.SECFACTU"
                    + " left outer join med.scea cea on"
                    + " cea.IDU = f.CCOUNIPS"
                    + " where f.TFACTURA <> 6"
                    + " group by f.CREFEREN, f.SECFACTU;";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Console.Write(".");

                    EndesaEntity.facturacion.mes13.Factura c = new EndesaEntity.facturacion.mes13.Factura();
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.ln = r["TIPONEGOCIO"].ToString();

                    if (r["EMPRESATITULAR"] != System.DBNull.Value)
                        c.empresaTitular = r["EMPRESATITULAR"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["FA"] != System.DBNull.Value)
                        c.fa = Convert.ToDateTime(r["FA"]);

                    if (r["FprevB"] != System.DBNull.Value)
                        c.fprevb = Convert.ToDateTime(r["FprevB"]);

                    if (r["Fecha_Baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["Fecha_Baja"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (c.nif.Substring(0, 2) == "NL")
                    {
                        c.ffactdes = fcd;
                        c.ffacthas = fch;
                        c.ffactdes_aux = new DateTime(2017, 12, 01);
                        c.ffacthas_aux = new DateTime(2017, 12, 31);
                    }
                    else
                    {
                        if (r["FFACTDES"] != System.DBNull.Value)
                            c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                        if (r["FFACTHAS"] != System.DBNull.Value)
                            c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);
                    }

                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        c.flimpago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"].ToString());

                    if (r["TAM"] != System.DBNull.Value)
                        c.tam = Convert.ToDouble(r["TAM"].ToString());

                    if (r["VNUMPLAZ"] != System.DBNull.Value)
                        c.vnumplaz = r["VNUMPLAZ"].ToString();

                    if (r["TIPO_EMPRESA"] != System.DBNull.Value)
                        c.tipoEmpresa = r["TIPO_EMPRESA"].ToString();

                    if (r["AMBITO"] != System.DBNull.Value)
                        c.territorial = r["AMBITO"].ToString();

                    if (r["NOMBRE_GESTOR"] != System.DBNull.Value)
                        c.gestor = r["NOMBRE_GESTOR"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["COMENTARIO"] != System.DBNull.Value)
                        c.comentario = r["COMENTARIO"].ToString();

                    lista.Add(c);
                }
                db.CloseConnection();
                Console.WriteLine();

                return lista;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }

        }
        private List<EndesaEntity.facturacion.mes13.Factura> Mes13_Total_Facturas_AgrupadasSinCUPS()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Console.WriteLine("Consultando facturas");
            List<EndesaEntity.facturacion.mes13.Factura> lista = new List<EndesaEntity.facturacion.mes13.Factura>();
            try
            {
                strSql = "select if (f.TIPONEGOCIO = 'L', 'LUZ', 'GAS') as TIPONEGOCIO,"
                    + " f.CEMPTITU as EMPRESATITULAR ,f.DAPERSOC as CLIENTE, f.CNIFDNIC as NIF,"
                    + " f.CCOUNIPS, f.CUPSREE,"
                    + " cea.FA, cea.FprevB, cea.Fecha_Baja,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " if (o.FLIMPAGO is null, d.FLIMPAGO,o.FLIMPAGO) as FLIMPAGO,"
                    + " f.CFACTURA, f.IFACTURA, t.TAM,"
                    + " null as VNUMPLAZ,"
                    + " s.tipoEmpresa as TIPO_EMPRESA,"
                    + " s.ambito as AMBITO, s.territorial as TERRITORIAL,"
                    + " concat(s.nombreGestor, ' ', s.Apellido1Gestor, ' ', s.Apellido2Gestor) as NOMBRE_GESTOR,"
                    + " f.TFACTURA, if(f.TIPONEGOCIO = 'L' AND f.CEMPTITU = 20 AND (f.CCOUNIPS is not null and f.CCOUNIPS <> '') AND cea.IDU is null, 'BAJA', null) as COMENTARIO"
                    + " from fact.13_facturas f"
                    + " left outer join cobr.carteraSIOC s on"
                    + " s.cif = f.CNIFDNIC"
                    + " left outer join fact.tam t on"
                    + " t.CUPS13 = f.CCOUNIPS"
                    + " left outer join fact.13_obligaciones_temp d on"
                    + " d.CEMPTITU = f.CEMPTITU and"
                    + " d.CREFEREN = lpad(f.CREFEREN, 13, 0) and"
                    + " d.SECFACTU = f.SECFACTU and"
                    + " d.FLIMPAGO <> '00000000'"
                    + " left outer join fact.13_oblig_cob o on"
                    + " o.CEMPTITU = f.CEMPTITU and"
                    + " o.CREFEREN = f.CREFEREN and"
                    + " o.SECFACTU = f.SECFACTU"
                    + " left outer join med.scea cea on"
                    + " cea.IDU = f.CCOUNIPS"
                    + " where (f.CCOUNIPS is null or trim(f.CCOUNIPS = '')) and (f.TFACTURA = 6 or f.TFACTURA = 5)"
                    + " group by f.CREFEREN, f.SECFACTU;";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Console.Write(".");

                    EndesaEntity.facturacion.mes13.Factura c = new EndesaEntity.facturacion.mes13.Factura();
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.ln = r["TIPONEGOCIO"].ToString();

                    if (r["EMPRESATITULAR"] != System.DBNull.Value)
                        c.empresaTitular = r["EMPRESATITULAR"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["FA"] != System.DBNull.Value)
                        c.fa = Convert.ToDateTime(r["FA"]);

                    if (r["FprevB"] != System.DBNull.Value)
                        c.fprevb = Convert.ToDateTime(r["FprevB"]);

                    if (r["Fecha_Baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["Fecha_Baja"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (c.nif.Substring(0, 2) == "NL")
                    {
                        c.ffactdes = fcd;
                        c.ffacthas = fch;
                        c.ffactdes_aux = new DateTime(2017, 12, 01);
                        c.ffacthas_aux = new DateTime(2017, 12, 31);
                    }
                    else
                    {
                        if (r["FFACTDES"] != System.DBNull.Value)
                            c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                        if (r["FFACTHAS"] != System.DBNull.Value)
                            c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);
                    }

                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        c.flimpago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"].ToString());

                    if (r["TAM"] != System.DBNull.Value)
                        c.tam = Convert.ToDouble(r["TAM"].ToString());

                    if (r["VNUMPLAZ"] != System.DBNull.Value)
                        c.vnumplaz = r["VNUMPLAZ"].ToString();

                    if (r["TIPO_EMPRESA"] != System.DBNull.Value)
                        c.tipoEmpresa = r["TIPO_EMPRESA"].ToString();

                    if (r["AMBITO"] != System.DBNull.Value)
                        c.territorial = r["AMBITO"].ToString();

                    if (r["NOMBRE_GESTOR"] != System.DBNull.Value)
                        c.gestor = r["NOMBRE_GESTOR"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["COMENTARIO"] != System.DBNull.Value)
                        c.comentario = r["COMENTARIO"].ToString();

                    lista.Add(c);
                }
                db.CloseConnection();
                Console.WriteLine();

                return lista;
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
                return null;
            }

        }
        private List<EndesaEntity.facturacion.mes13.Factura> Mes13_Total_Facturas_AgrupadasConCUPS()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Console.WriteLine("Consultando facturas Agrupadas con CUPS informado");
            List<EndesaEntity.facturacion.mes13.Factura> lista = new List<EndesaEntity.facturacion.mes13.Factura>();
            try
            {
                strSql = "select if (f.TIPONEGOCIO = 'L', 'LUZ', 'GAS') as TIPONEGOCIO,"
                    + " f.CEMPTITU as EMPRESATITULAR ,f.DAPERSOC as CLIENTE, f.CNIFDNIC as NIF,"
                    + " f.CCOUNIPS, f.CUPSREE,"
                    + " cea.FA, cea.FprevB, cea.Fecha_Baja,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " if (o.FLIMPAGO is null, d.FLIMPAGO,o.FLIMPAGO) as FLIMPAGO,"
                    + " f.CFACTURA, f.IFACTURA, t.TAM,"
                    + " null as VNUMPLAZ,"
                    + " s.tipoEmpresa as TIPO_EMPRESA,"
                    + " s.ambito as AMBITO, s.territorial as TERRITORIAL,"
                    + " concat(s.nombreGestor, ' ', s.Apellido1Gestor, ' ', s.Apellido2Gestor) as NOMBRE_GESTOR,"
                    + " f.TFACTURA, if(f.TIPONEGOCIO = 'L' AND f.CEMPTITU = 20 AND (f.CCOUNIPS is not null and f.CCOUNIPS <> '') AND cea.IDU is null, 'BAJA', null) as COMENTARIO"
                    + " from fact.13_facturas f"
                    + " left outer join cobr.carteraSIOC s on"
                    + " s.cif = f.CNIFDNIC"
                    + " left outer join fact.tam t on"
                    + " t.CUPS13 = f.CCOUNIPS"
                    + " left outer join fact.13_obligaciones_temp d on"
                    + " d.CEMPTITU = f.CEMPTITU and"
                    + " d.CREFEREN = lpad(f.CREFEREN, 13, 0) and"
                    + " d.SECFACTU = f.SECFACTU and"
                    + " d.FLIMPAGO <> '00000000'"
                    + " left outer join fact.13_oblig_cob o on"
                    + " o.CEMPTITU = f.CEMPTITU and"
                    + " o.CREFEREN = f.CREFEREN and"
                    + " o.SECFACTU = f.SECFACTU"
                    + " left outer join med.scea cea on"
                    + " cea.IDU = f.CCOUNIPS"
                    + " where (f.CCOUNIPS <> '') and (f.TFACTURA = 6 or f.TACTURA = 5)"
                    + " group by f.CREFEREN, f.SECFACTU;";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Console.Write(".");

                    EndesaEntity.facturacion.mes13.Factura c = new EndesaEntity.facturacion.mes13.Factura();
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.ln = r["TIPONEGOCIO"].ToString();

                    if (r["EMPRESATITULAR"] != System.DBNull.Value)
                        c.empresaTitular = r["EMPRESATITULAR"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["FA"] != System.DBNull.Value)
                        c.fa = Convert.ToDateTime(r["FA"]);

                    if (r["FprevB"] != System.DBNull.Value)
                        c.fprevb = Convert.ToDateTime(r["FprevB"]);

                    if (r["Fecha_Baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["Fecha_Baja"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        c.flimpago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"].ToString());

                    if (r["TAM"] != System.DBNull.Value)
                        c.tam = Convert.ToDouble(r["TAM"].ToString());

                    if (r["VNUMPLAZ"] != System.DBNull.Value)
                        c.vnumplaz = r["VNUMPLAZ"].ToString();

                    if (r["TIPO_EMPRESA"] != System.DBNull.Value)
                        c.tipoEmpresa = r["TIPO_EMPRESA"].ToString();

                    if (r["AMBITO"] != System.DBNull.Value)
                        c.territorial = r["AMBITO"].ToString();

                    if (r["NOMBRE_GESTOR"] != System.DBNull.Value)
                        c.gestor = r["NOMBRE_GESTOR"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["COMENTARIO"] != System.DBNull.Value)
                        c.comentario = r["COMENTARIO"].ToString();




                    lista.Add(c);
                }
                db.CloseConnection();
                Console.WriteLine();

                return lista;
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
                return null;
            }

        }
        private List<EndesaEntity.facturacion.mes13.Factura> Mes13_Total_Facturas_Errores()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Console.WriteLine("Consultando facturas con posibles errores");
            List<EndesaEntity.facturacion.mes13.Factura> lista = new List<EndesaEntity.facturacion.mes13.Factura>();
            try
            {
                strSql = "select if (f.TIPONEGOCIO = 'L', 'LUZ', 'GAS') as TIPONEGOCIO,"
                    + " f.CEMPTITU as EMPRESATITULAR ,f.DAPERSOC as CLIENTE, f.CNIFDNIC as NIF,"
                    + " f.CCOUNIPS, f.CUPSREE,"
                    + " cea.FA, cea.FprevB, cea.Fecha_Baja,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " if (o.FLIMPAGO is null, d.FLIMPAGO,o.FLIMPAGO) as FLIMPAGO,"
                    + " f.CFACTURA, f.IFACTURA, t.TAM,"
                    + " null as VNUMPLAZ,"
                    + " s.tipoEmpresa as TIPO_EMPRESA,"
                    + " s.ambito as AMBITO, s.territorial as TERRITORIAL,"
                    + " concat(s.nombreGestor, ' ', s.Apellido1Gestor, ' ', s.Apellido2Gestor) as NOMBRE_GESTOR,"
                    + " f.TFACTURA"
                    + " from fact.13_facturas f"
                    + " left outer join cobr.carteraSIOC s on"
                    + " s.cif = f.CNIFDNIC"
                    + " left outer join fact.tam t on"
                    + " t.CUPS13 = f.CCOUNIPS"
                    + " left outer join fact.13_obligaciones_temp d on"
                    + " d.CEMPTITU = f.CEMPTITU and"
                    + " d.CREFEREN = lpad(f.CREFEREN, 13, 0) and"
                    + " d.SECFACTU = f.SECFACTU and"
                    + " d.FLIMPAGO <> '00000000'"
                    + " left outer join fact.13_oblig_cob o on"
                    + " o.CEMPTITU = f.CEMPTITU and"
                    + " o.CREFEREN = f.CREFEREN and"
                    + " o.SECFACTU = f.SECFACTU"
                    + " left outer join med.scea cea on"
                    + " cea.IDU = f.CCOUNIPS"
                    + " where (f.CCOUNIPS <> '') and (f.TFACTURA = 6 or f.TACTURA = 5)"
                    + " group by f.CREFEREN, f.SECFACTU;";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Console.Write(".");

                    EndesaEntity.facturacion.mes13.Factura c = new EndesaEntity.facturacion.mes13.Factura();
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.ln = r["TIPONEGOCIO"].ToString();

                    if (r["EMPRESATITULAR"] != System.DBNull.Value)
                        c.empresaTitular = r["EMPRESATITULAR"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.ccounips = r["CCOUNIPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cupsree = r["CUPSREE"].ToString();

                    if (r["FA"] != System.DBNull.Value)
                        c.fa = Convert.ToDateTime(r["FA"]);

                    if (r["FprevB"] != System.DBNull.Value)
                        c.fprevb = Convert.ToDateTime(r["FprevB"]);

                    if (r["Fecha_Baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["Fecha_Baja"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        c.flimpago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"].ToString());

                    if (r["TAM"] != System.DBNull.Value)
                        c.tam = Convert.ToDouble(r["TAM"].ToString());

                    if (r["VNUMPLAZ"] != System.DBNull.Value)
                        c.vnumplaz = r["VNUMPLAZ"].ToString();

                    if (r["TIPO_EMPRESA"] != System.DBNull.Value)
                        c.tipoEmpresa = r["TIPO_EMPRESA"].ToString();

                    if (r["AMBITO"] != System.DBNull.Value)
                        c.territorial = r["AMBITO"].ToString();

                    if (r["NOMBRE_GESTOR"] != System.DBNull.Value)
                        c.gestor = r["NOMBRE_GESTOR"].ToString();

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["COMENTARIO"] != System.DBNull.Value)
                        c.comentario = r["COMENTARIO"].ToString();

                    lista.Add(c);
                }
                db.CloseConnection();
                Console.WriteLine();

                return lista;
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
                return null;
            }

        }

        private void Mes13_Excel_Consumos()
        {

            string rutaFichero = @"c:\Temp\consumos_201808_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

            Console.WriteLine("Generado Excel --> " + rutaFichero);
            FileInfo file = new FileInfo(rutaFichero);

            int c = 0;
            int f = 0;


            if (file.Exists)
                file.Delete();

            ExcelPackage excelPackage = new ExcelPackage(file);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            var workSheet = excelPackage.Workbook.Worksheets.Add("Total");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            #region Cebecera_Hoja_Total
            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "TIPO NEGOCIO"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "FA"; c++;
            workSheet.Cells[f, c].Value = "FprevB"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "FLIMPAGO"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TAM"; c++;
            workSheet.Cells[f, c].Value = "TIPO EMPRESA"; c++;
            workSheet.Cells[f, c].Value = "ÁMBITO"; c++;
            workSheet.Cells[f, c].Value = "TERRITORIAL"; c++;
            workSheet.Cells[f, c].Value = "GESTOR"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "COMENTARIO"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA_AUX"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES_AUX"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS_AUX"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA_AUX"; c++;

            headerFont.Bold = true;
            #endregion

            List<EndesaEntity.facturacion.mes13.Factura> l = Mes13_Total_Facturas();
            #region Pegado_Excel_Total
            for (int i = 0; i < l.Count(); i++)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = l[i].ln; c++;
                workSheet.Cells[f, c].Value = l[i].empresaTitular; c++;
                workSheet.Cells[f, c].Value = l[i].cliente; c++;
                workSheet.Cells[f, c].Value = l[i].nif; c++;
                workSheet.Cells[f, c].Value = l[i].ccounips; c++;
                workSheet.Cells[f, c].Value = l[i].cupsree; c++;

                if (l[i].fa != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fprevb != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fprevb;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fecha_baja != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactura != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactdes != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactdes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffacthas != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffacthas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].flimpago != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].flimpago;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                workSheet.Cells[f, c].Value = l[i].cfactura; c++;
                workSheet.Cells[f, c].Value = l[i].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = l[i].tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = l[i].tipoEmpresa; c++;
                workSheet.Cells[f, c].Value = l[i].ambito; c++;
                workSheet.Cells[f, c].Value = l[i].territorial; c++;
                workSheet.Cells[f, c].Value = l[i].gestor; c++;
                workSheet.Cells[f, c].Value = l[i].tfactura; c++;
                workSheet.Cells[f, c].Value = l[i].comentario; c++;

                if (l[i].ffactura_aux != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactura_aux;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactdes_aux != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactdes_aux;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffacthas_aux != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffacthas_aux;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ifactura_aux > 0)
                {
                    workSheet.Cells[f, c].Value = l[i].ifactura_aux;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                }

            }



            var allCells = workSheet.Cells[1, 1, f, 50];
            headerFont = headerCells.Style.Font;

            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:V1"].AutoFilter = true;
            allCells.AutoFitColumns();

            #endregion

            l.Clear();
            l = Mes13_Total_Facturas_SinAgrupadas();
            workSheet = excelPackage.Workbook.Worksheets.Add("CUPS_Agregados_Sin_Agrupar");

            #region Cebecera_Hoja_CUPS_Agregados_Sin_Agrupar
            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "TIPO NEGOCIO"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "FA"; c++;
            workSheet.Cells[f, c].Value = "FprevB"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "FLIMPAGO"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TAM"; c++;
            workSheet.Cells[f, c].Value = "VNUMPLAZ"; c++;
            workSheet.Cells[f, c].Value = "TIPO EMPRESA"; c++;
            workSheet.Cells[f, c].Value = "ÁMBITO"; c++;
            workSheet.Cells[f, c].Value = "TERRITORIAL"; c++;
            workSheet.Cells[f, c].Value = "GESTOR"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "COMENTARIO"; c++;
            headerFont.Bold = true;
            #endregion


            #region Pegado_Excel_CUPS_Agregados_Sin_Agrupar
            for (int i = 0; i < l.Count(); i++)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = l[i].ln; c++;
                workSheet.Cells[f, c].Value = l[i].empresaTitular; c++;
                workSheet.Cells[f, c].Value = l[i].cliente; c++;
                workSheet.Cells[f, c].Value = l[i].nif; c++;
                workSheet.Cells[f, c].Value = l[i].ccounips; c++;
                workSheet.Cells[f, c].Value = l[i].cupsree; c++;

                if (l[i].fa != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fprevb != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fprevb;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fecha_baja != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactura != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactdes != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactdes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffacthas != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffacthas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].flimpago != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].flimpago;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                workSheet.Cells[f, c].Value = l[i].cfactura; c++;
                workSheet.Cells[f, c].Value = l[i].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = l[i].tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                c++;
                workSheet.Cells[f, c].Value = l[i].tipoEmpresa; c++;
                workSheet.Cells[f, c].Value = l[i].ambito; c++;
                workSheet.Cells[f, c].Value = l[i].territorial; c++;
                workSheet.Cells[f, c].Value = l[i].gestor; c++;
                workSheet.Cells[f, c].Value = l[i].tfactura; c++;
                workSheet.Cells[f, c].Value = l[i].comentario; c++;

            }

            allCells = workSheet.Cells[1, 1, f, 50];
            allCells.AutoFitColumns();

            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            workSheet.Cells["A1:W1"].AutoFilter = true;

            #endregion

            l.Clear();
            l = Mes13_Total_Facturas_AgrupadasSinCUPS();
            workSheet = excelPackage.Workbook.Worksheets.Add("Agrupadas_Sin_CUPS");

            #region Cebecera_Hoja_CUPS_Agrupadas_Sin_CUPS
            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "TIPO NEGOCIO"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "FA"; c++;
            workSheet.Cells[f, c].Value = "FprevB"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "FLIMPAGO"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TAM"; c++;
            workSheet.Cells[f, c].Value = "VNUMPLAZ"; c++;
            workSheet.Cells[f, c].Value = "TIPO EMPRESA"; c++;
            workSheet.Cells[f, c].Value = "ÁMBITO"; c++;
            workSheet.Cells[f, c].Value = "TERRITORIAL"; c++;
            workSheet.Cells[f, c].Value = "GESTOR"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "COMENTARIO"; c++;
            headerFont.Bold = true;
            #endregion

            #region Pegado_Excel_CUPS_Agregados_Sin_CUPS
            for (int i = 0; i < l.Count(); i++)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = l[i].ln; c++;
                workSheet.Cells[f, c].Value = l[i].empresaTitular; c++;
                workSheet.Cells[f, c].Value = l[i].cliente; c++;
                workSheet.Cells[f, c].Value = l[i].nif; c++;
                workSheet.Cells[f, c].Value = l[i].ccounips; c++;
                workSheet.Cells[f, c].Value = l[i].cupsree; c++;

                if (l[i].fa != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fprevb != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fprevb;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fecha_baja != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactura != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactdes != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactdes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffacthas != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffacthas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].flimpago != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].flimpago;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                workSheet.Cells[f, c].Value = l[i].cfactura; c++;
                workSheet.Cells[f, c].Value = l[i].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = l[i].tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                c++;
                workSheet.Cells[f, c].Value = l[i].tipoEmpresa; c++;
                workSheet.Cells[f, c].Value = l[i].ambito; c++;
                workSheet.Cells[f, c].Value = l[i].territorial; c++;
                workSheet.Cells[f, c].Value = l[i].gestor; c++;
                workSheet.Cells[f, c].Value = l[i].tfactura; c++;
                workSheet.Cells[f, c].Value = l[i].comentario; c++;

            }

            allCells = workSheet.Cells[1, 1, f, 50];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            workSheet.Cells["A1:W1"].AutoFilter = true;

            #endregion



            #region Cebecera_Hoja_CUPS_Agrupadas_Con_CUPS

            l.Clear();
            l = Mes13_Total_Facturas_AgrupadasConCUPS();
            workSheet = excelPackage.Workbook.Worksheets.Add("Agrupadas_Con_CUPS");

            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "TIPO NEGOCIO"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "FA"; c++;
            workSheet.Cells[f, c].Value = "FprevB"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "FLIMPAGO"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TAM"; c++;
            workSheet.Cells[f, c].Value = "TIPO EMPRESA"; c++;
            workSheet.Cells[f, c].Value = "ÁMBITO"; c++;
            workSheet.Cells[f, c].Value = "TERRITORIAL"; c++;
            workSheet.Cells[f, c].Value = "GESTOR"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            workSheet.Cells[f, c].Value = "COMENTARIO"; c++;
            headerFont.Bold = true;
            #endregion

            #region Pegado_Excel_Agrupadas_Con_CUPS
            for (int i = 0; i < l.Count(); i++)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = l[i].ln; c++;
                workSheet.Cells[f, c].Value = l[i].empresaTitular; c++;
                workSheet.Cells[f, c].Value = l[i].cliente; c++;
                workSheet.Cells[f, c].Value = l[i].nif; c++;
                workSheet.Cells[f, c].Value = l[i].ccounips; c++;
                workSheet.Cells[f, c].Value = l[i].cupsree; c++;

                if (l[i].fa != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fprevb != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fprevb;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].fecha_baja != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactura != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffactdes != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffactdes;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].ffacthas != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].ffacthas;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (l[i].flimpago != DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = l[i].flimpago;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                workSheet.Cells[f, c].Value = l[i].cfactura; c++;
                workSheet.Cells[f, c].Value = l[i].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = l[i].tam; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = l[i].tipoEmpresa; c++;
                workSheet.Cells[f, c].Value = l[i].ambito; c++;
                workSheet.Cells[f, c].Value = l[i].territorial; c++;
                workSheet.Cells[f, c].Value = l[i].gestor; c++;
                workSheet.Cells[f, c].Value = l[i].tfactura; c++;
                workSheet.Cells[f, c].Value = l[i].comentario; c++;

            }


            workSheet.View.FreezePanes(2, 1);
            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, f, 50];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:W1"].AutoFilter = true;

            #endregion

            #region Errores
            l.Clear();
            l = Mes13_Total_Facturas_Errores();
            workSheet = excelPackage.Workbook.Worksheets.Add("Errores");

            f = 1;
            c = 1;
            workSheet.Cells[f, c].Value = "TIPO NEGOCIO"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA TITULAR"; c++;
            workSheet.Cells[f, c].Value = "CLIENTE"; c++;
            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "FA"; c++;
            workSheet.Cells[f, c].Value = "FprevB"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "FLIMPAGO"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            workSheet.Cells[f, c].Value = "TAM"; c++;
            workSheet.Cells[f, c].Value = "TIPO EMPRESA"; c++;
            workSheet.Cells[f, c].Value = "ÁMBITO"; c++;
            workSheet.Cells[f, c].Value = "TERRITORIAL"; c++;
            workSheet.Cells[f, c].Value = "GESTOR"; c++;
            workSheet.Cells[f, c].Value = "TFACTURA"; c++;
            headerFont.Bold = true;
            #endregion


            workSheet.View.FreezePanes(2, 1);
            allCells = workSheet.Cells[1, 1, f, 50];
            workSheet.Cells["A1:U1"].AutoFilter = true;
            allCells.AutoFitColumns();

            excelPackage.Save();

        }

        private void Mes13_AgruparFacturas_PerdiodosPartidos(DateTime fcd)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            MySQLDB db2;
            MySqlCommand command2;
            MySqlDataReader r2;
            Console.WriteLine("Consultando facturas de periodos fraccionados.");
            List<EndesaEntity.facturacion.mes13.FacturaMes13_Seguimiento> lista = new List<EndesaEntity.facturacion.mes13.FacturaMes13_Seguimiento>();
            int i = 0;

            //try
            //{

            //strSql = "delete from 13_facturas_2;";
            //db = new MySQLDB(MySQLDB.Esquemas.FAC);
            //command = new MySqlCommand(strSql, db.con);
            //command.ExecuteNonQuery();
            //db.CloseConnection();

            strSql = "select * from 13_facturas f where f.FFACTDES > '" + fcd.ToString("yyyy-MM-dd") + "'";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                i++;
                db2 = new MySQLDB(MySQLDB.Esquemas.FAC);
                command2 = new MySqlCommand(Busca_Periodo_Factura(Convert.ToInt64(r["CREFEREN"]), Convert.ToInt32(r["SECFACTU"])), db2.con);
                r2 = command2.ExecuteReader();
                while (r2.Read())
                {
                    GuardaFactura(r2, r);
                    BorraGrupoFacturas(Convert.ToInt64(r["CREFEREN"]), Convert.ToInt32(r["SECFACTU"]), i);
                    BorraGrupoFacturas(Convert.ToInt64(r2["CREFEREN"]), Convert.ToInt32(r2["SECFACTU"]), i);

                }
                db2.CloseConnection();
            }

            db.CloseConnection();

            strSql = "replace into 13_facturas select * from 13_facturas_2;";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            strSql = "delete from 13_facturas where IFACTURA < 5000;";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // 2.- Borramos las facturas de clientes que no superen los 20.000 €	
            Console.WriteLine("Borrando facturas de clietes que no superen los 20.000 €");
            Console.WriteLine("delete from 13_nifs");
            strSql = "delete from 13_nifs;";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "replace into 13_nifs select f.CNIFDNIC, f.TIPONEGOCIO, sum(f.IFACTURA) as total"
                + " from 13_facturas f"
                + " group by f.CNIFDNIC, f.TIPONEGOCIO having total < 20000;";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "delete f"
                + " from 13_facturas f inner join"
                + " 13_nifs r on"
                + " r.CNIFDNIC = f.CNIFDNIC;";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //}
        }

        private void GuardaFactura(MySqlDataReader r1, MySqlDataReader r2)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {


                #region Cabecera replace
                strSql = "replace into 13_facturas_2 (CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CREFFACT, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, VCUOVAFA, VENEREAC, VCUOFIFA,"
                    + " IFACTURA, IVA, IIMPUES2, IIMPUES3, IBASEISE, ISE, PRECIO_MEDIO, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC,"
                    + " TCONFAC1, ICONFAC1, TCONFAC2, ICONFAC2, TCONFAC3, ICONFAC3, TCONFAC4, ICONFAC4, TCONFAC5, ICONFAC5,"
                    + " TCONFAC6, ICONFAC6, TCONFAC7, ICONFAC7, TCONFAC8, ICONFAC8, TCONFAC9, ICONFAC9, TCONFA10, ICONFA10,"
                    + " TCONFA11, ICONFA11, TCONFA12, ICONFA12, TCONFA13, ICONFA13, TCONFA14, ICONFA14, TCONFA15, ICONFA15,"
                    + " TCONFA16, ICONFA16, TCONFA17, ICONFA17, TCONFA18, ICONFA18, TCONFA19, ICONFA19, TCONFA20, ICONFA20,"
                    + " TINDGCPY, TMODOPTA, CREFAEXT, KPERFACT, MOTIVO_REFACTURACION, SUBMOTIVO, TIPO_COMENTARIO, COMENTARIO_REFACT,"
                    + " VPOTCON1, VPOTCON2, VPOTCON3, VPOTCON4, VPOTCON5, VPOTCON6, VPOTMAX1, VPOTMAX2, VPOTMAX3, VPOTMAX4,"
                    + " VPOTMAX5, VPOTMAX6, VCONATH1, VCONATH2, VCONATH3, VCONATH4, VCONATH5, VCONATH6, VCONATHP, VCONRTH1,"
                    + " VCONRTH2, VCONRTH3, VCONRTH4, VCONRTH5, VCONRTH6, VCONRTHP, VEXCERE1, VEXCERE2, VEXCERE3, VEXCERE4,"
                    + " VEXCERE5, VEXCERE6, PRECIAC1, PRECIAC2, PRECIAC3, PRECIAC4, PRECIAC5, PRECIAC6, PRECIPO1, PRECIPO2,"
                    + " PRECIPO3, PRECIPO4, PRECIPO5, PRECIPO6, VEXCEPO1, VEXCEPO2, VEXCEPO3, VEXCEPO4, VEXCEPO5, VEXCEPO6,"
                    + " VPOTCALL, VPOTCALV, VPOTCALP, VPOTMAXL, VPOTMAXV, VPOTMAXP, VCONSACL, VCONSACV, VCONSACP, VCONSREL,"
                    + " VCONSREV, VCONSREP, VEXCEREL, VEXCEREV, VEXCEREP, PRECIAB1, PRECIAB2, PRECIAB3, PRECIPB1, PRECIPB2,"
                    + " PRECIPB3, VEXCEPOL, VEXCEPOV, VEXCEPOP, PRECIDHL, PRECIDHV, PRECIDHP, INDEMPRE, CUPSREE, CFACTREC,"
                    + " CLINNEG, CCOMAUTO, CPROVINC, CSEGMERC, NUMLABOR, TIPONEGOCIO, SUPERVALLE, ENERG_PONTA, HORAS_PONTA,"
                    + " POT_MAXIMA, COMENTARIOS) values";

                #endregion

                #region Campos
                if (r2["CCOUNIPS"] != System.DBNull.Value)
                    strSql += " ('" + r2["CCOUNIPS"].ToString() + "', ";
                else
                    strSql += " (null,";

                if (r2["CEMPTITU"] != System.DBNull.Value)
                    strSql += r2["CEMPTITU"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["CREFEREN"] != System.DBNull.Value)
                    strSql += r2["CREFEREN"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["SECFACTU"] != System.DBNull.Value)
                    strSql += r2["SECFACTU"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["CREFFACT"] != System.DBNull.Value)
                    strSql += "'" + r2["CREFFACT"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CFACTURA"] != System.DBNull.Value)
                    strSql += "'" + r2["CFACTURA"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["FFACTURA"] != System.DBNull.Value)
                    strSql += "'" + Convert.ToDateTime(r2["FFACTURA"]).ToString("yyyy-MM-dd") + "', ";
                else
                    strSql += " null,";

                if (r1["FFACTDES"] != System.DBNull.Value)
                    strSql += "'" + Convert.ToDateTime(r1["FFACTDES"]).ToString("yyyy-MM-dd") + "', ";
                else
                    strSql += " null,";

                if (r2["FFACTHAS"] != System.DBNull.Value)
                    strSql += "'" + Convert.ToDateTime(r2["FFACTHAS"]).ToString("yyyy-MM-dd") + "', ";
                else
                    strSql += " null,";

                if (r2["VCUOVAFA"] != System.DBNull.Value)
                    strSql += r2["VCUOVAFA"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["VENEREAC"] != System.DBNull.Value)
                    strSql += r2["VENEREAC"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["VCUOFIFA"] != System.DBNull.Value)
                    strSql += r2["VCUOFIFA"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["IFACTURA"] != System.DBNull.Value)
                {
                    double totalIFactura =
                        Convert.ToDouble(r1["IFACTURA"]) + Convert.ToDouble(r2["IFACTURA"]);
                    strSql += totalIFactura.ToString().Replace(",", ".") + ", ";
                }
                else
                    strSql += " null,";

                if (r2["IVA"] != System.DBNull.Value)
                    strSql += r2["IVA"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["IIMPUES2"] != System.DBNull.Value)
                    strSql += r2["IIMPUES2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["IIMPUES3"] != System.DBNull.Value)
                    strSql += r2["IIMPUES3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["IBASEISE"] != System.DBNull.Value)
                    strSql += r2["IBASEISE"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["ISE"] != System.DBNull.Value)
                    strSql += r2["ISE"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIO_MEDIO"] != System.DBNull.Value)
                    strSql += r2["PRECIO_MEDIO"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TFACTURA"] != System.DBNull.Value)
                    strSql += r2["TFACTURA"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["TESTFACT"] != System.DBNull.Value)
                    strSql += "'" + r2["TESTFACT"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["DAPERSOC"] != System.DBNull.Value)
                    strSql += "'" + r2["DAPERSOC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CNIFDNIC"] != System.DBNull.Value)
                    strSql += "'" + r2["CNIFDNIC"].ToString() + "', ";
                else
                    strSql += " null,";
                #endregion

                #region TCONFAC

                if (r2["TCONFAC1"] != System.DBNull.Value)
                    strSql += r2["TCONFAC1"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC1"] != System.DBNull.Value)
                    strSql += r2["ICONFAC1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC2"] != System.DBNull.Value)
                    strSql += r2["TCONFAC2"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC2"] != System.DBNull.Value)
                    strSql += r2["ICONFAC2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC3"] != System.DBNull.Value)
                    strSql += r2["TCONFAC3"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC3"] != System.DBNull.Value)
                    strSql += r2["ICONFAC3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC4"] != System.DBNull.Value)
                    strSql += r2["TCONFAC4"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC4"] != System.DBNull.Value)
                    strSql += r2["ICONFAC4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC5"] != System.DBNull.Value)
                    strSql += r2["TCONFAC5"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC5"] != System.DBNull.Value)
                    strSql += r2["ICONFAC5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC6"] != System.DBNull.Value)
                    strSql += r2["TCONFAC6"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC6"] != System.DBNull.Value)
                    strSql += r2["ICONFAC6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC7"] != System.DBNull.Value)
                    strSql += r2["TCONFAC7"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC7"] != System.DBNull.Value)
                    strSql += r2["ICONFAC7"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC8"] != System.DBNull.Value)
                    strSql += r2["TCONFAC8"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC8"] != System.DBNull.Value)
                    strSql += r2["ICONFAC8"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFAC9"] != System.DBNull.Value)
                    strSql += r2["TCONFAC9"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFAC9"] != System.DBNull.Value)
                    strSql += r2["ICONFAC9"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA10"] != System.DBNull.Value)
                    strSql += r2["TCONFA10"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA10"] != System.DBNull.Value)
                    strSql += r2["ICONFA10"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA11"] != System.DBNull.Value)

                    strSql += r2["TCONFA11"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA11"] != System.DBNull.Value)
                    strSql += r2["ICONFA11"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA12"] != System.DBNull.Value)

                    strSql += r2["TCONFA12"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA12"] != System.DBNull.Value)
                    strSql += r2["ICONFA12"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA13"] != System.DBNull.Value)

                    strSql += r2["TCONFA13"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA13"] != System.DBNull.Value)
                    strSql += r2["ICONFA13"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA14"] != System.DBNull.Value)

                    strSql += r2["TCONFA14"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA14"] != System.DBNull.Value)
                    strSql += r2["ICONFA14"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA15"] != System.DBNull.Value)

                    strSql += r2["TCONFA15"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA15"] != System.DBNull.Value)
                    strSql += r2["ICONFA15"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA16"] != System.DBNull.Value)

                    strSql += r2["TCONFA16"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA16"] != System.DBNull.Value)
                    strSql += r2["ICONFA16"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA17"] != System.DBNull.Value)

                    strSql += r2["TCONFA17"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA17"] != System.DBNull.Value)
                    strSql += r2["ICONFA17"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA18"] != System.DBNull.Value)

                    strSql += r2["TCONFA18"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA18"] != System.DBNull.Value)
                    strSql += r2["ICONFA18"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA19"] != System.DBNull.Value)

                    strSql += r2["TCONFA19"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA19"] != System.DBNull.Value)
                    strSql += r2["ICONFA19"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["TCONFA20"] != System.DBNull.Value)

                    strSql += r2["TCONFA20"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["ICONFA20"] != System.DBNull.Value)
                    strSql += r2["ICONFA20"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                #endregion

                #region RestoCampos
                if (r2["TINDGCPY"] != System.DBNull.Value)
                    strSql += "'" + r2["TINDGCPY"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["TMODOPTA"] != System.DBNull.Value)
                    strSql += "'" + r2["TMODOPTA"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CREFAEXT"] != System.DBNull.Value)
                    strSql += "'" + r2["CREFAEXT"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["KPERFACT"] != System.DBNull.Value)
                    strSql += "'" + r2["KPERFACT"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["MOTIVO_REFACTURACION"] != System.DBNull.Value)
                    strSql += "'" + r2["MOTIVO_REFACTURACION"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["SUBMOTIVO"] != System.DBNull.Value)
                    strSql += "'" + r2["SUBMOTIVO"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["TIPO_COMENTARIO"] != System.DBNull.Value)
                    strSql += "'" + r2["TIPO_COMENTARIO"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["COMENTARIO_REFACT"] != System.DBNull.Value)
                    strSql += "'" + r2["COMENTARIO_REFACT"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["VPOTCON1"] != System.DBNull.Value)
                    strSql += r2["VPOTCON1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCON2"] != System.DBNull.Value)
                    strSql += r2["VPOTCON2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCON3"] != System.DBNull.Value)
                    strSql += r2["VPOTCON3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCON4"] != System.DBNull.Value)
                    strSql += r2["VPOTCON4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCON5"] != System.DBNull.Value)
                    strSql += r2["VPOTCON5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCON6"] != System.DBNull.Value)
                    strSql += r2["VPOTCON6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAX1"] != System.DBNull.Value)
                    strSql += r2["VPOTMAX1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAX2"] != System.DBNull.Value)
                    strSql += r2["VPOTMAX2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAX3"] != System.DBNull.Value)
                    strSql += r2["VPOTMAX3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAX4"] != System.DBNull.Value)
                    strSql += r2["VPOTMAX4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAX5"] != System.DBNull.Value)
                    strSql += r2["VPOTMAX5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAX6"] != System.DBNull.Value)
                    strSql += r2["VPOTMAX6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATH1"] != System.DBNull.Value)
                    strSql += r2["VCONATH1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATH2"] != System.DBNull.Value)
                    strSql += r2["VCONATH2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATH3"] != System.DBNull.Value)
                    strSql += r2["VCONATH3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATH4"] != System.DBNull.Value)
                    strSql += r2["VCONATH4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATH5"] != System.DBNull.Value)
                    strSql += r2["VCONATH5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATH6"] != System.DBNull.Value)
                    strSql += r2["VCONATH6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONATHP"] != System.DBNull.Value)
                    strSql += r2["VCONATHP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTH1"] != System.DBNull.Value)
                    strSql += r2["VCONRTH1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTH2"] != System.DBNull.Value)
                    strSql += r2["VCONRTH2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTH3"] != System.DBNull.Value)
                    strSql += r2["VCONRTH3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTH4"] != System.DBNull.Value)
                    strSql += r2["VCONRTH4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTH5"] != System.DBNull.Value)
                    strSql += r2["VCONRTH5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTH6"] != System.DBNull.Value)
                    strSql += r2["VCONRTH6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONRTHP"] != System.DBNull.Value)
                    strSql += r2["VCONRTHP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCERE1"] != System.DBNull.Value)
                    strSql += r2["VEXCERE1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCERE2"] != System.DBNull.Value)
                    strSql += r2["VEXCERE2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCERE3"] != System.DBNull.Value)
                    strSql += r2["VEXCERE3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCERE4"] != System.DBNull.Value)
                    strSql += r2["VEXCERE4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCERE5"] != System.DBNull.Value)
                    strSql += r2["VEXCERE5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCERE6"] != System.DBNull.Value)
                    strSql += r2["VEXCERE6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAC1"] != System.DBNull.Value)
                    strSql += r2["PRECIAC1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAC2"] != System.DBNull.Value)
                    strSql += r2["PRECIAC2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAC3"] != System.DBNull.Value)
                    strSql += r2["PRECIAC3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAC4"] != System.DBNull.Value)
                    strSql += r2["PRECIAC4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAC5"] != System.DBNull.Value)
                    strSql += r2["PRECIAC5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAC6"] != System.DBNull.Value)
                    strSql += r2["PRECIAC6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPO1"] != System.DBNull.Value)
                    strSql += r2["PRECIPO1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPO2"] != System.DBNull.Value)
                    strSql += r2["PRECIPO2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPO3"] != System.DBNull.Value)
                    strSql += r2["PRECIPO3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPO4"] != System.DBNull.Value)
                    strSql += r2["PRECIPO4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPO5"] != System.DBNull.Value)
                    strSql += r2["PRECIPO5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPO6"] != System.DBNull.Value)
                    strSql += r2["PRECIPO6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPO1"] != System.DBNull.Value)
                    strSql += r2["VEXCEPO1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPO2"] != System.DBNull.Value)
                    strSql += r2["VEXCEPO2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPO3"] != System.DBNull.Value)
                    strSql += r2["VEXCEPO3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPO4"] != System.DBNull.Value)
                    strSql += r2["VEXCEPO4"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPO5"] != System.DBNull.Value)
                    strSql += r2["VEXCEPO5"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPO6"] != System.DBNull.Value)
                    strSql += r2["VEXCEPO6"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCALL"] != System.DBNull.Value)
                    strSql += r2["VPOTCALL"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCALV"] != System.DBNull.Value)
                    strSql += r2["VPOTCALV"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTCALP"] != System.DBNull.Value)
                    strSql += r2["VPOTCALP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAXL"] != System.DBNull.Value)
                    strSql += r2["VPOTMAXL"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAXV"] != System.DBNull.Value)
                    strSql += r2["VPOTMAXV"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VPOTMAXP"] != System.DBNull.Value)
                    strSql += r2["VPOTMAXP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONSACL"] != System.DBNull.Value)
                    strSql += r2["VCONSACL"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONSACV"] != System.DBNull.Value)
                    strSql += r2["VCONSACV"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONSACP"] != System.DBNull.Value)
                    strSql += r2["VCONSACP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONSREL"] != System.DBNull.Value)
                    strSql += r2["VCONSREL"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VCONSREV"] != System.DBNull.Value)
                    strSql += r2["VCONSREV"].ToString().Replace(",", ".") + " ,";
                else
                    strSql += " null,";

                if (r2["VCONSREP"] != System.DBNull.Value)
                    strSql += r2["VCONSREP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEREL"] != System.DBNull.Value)
                    strSql += r2["VEXCEREL"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEREV"] != System.DBNull.Value)
                    strSql += r2["VEXCEREV"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEREP"] != System.DBNull.Value)
                    strSql += r2["VEXCEREP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAB1"] != System.DBNull.Value)
                    strSql += r2["PRECIAB1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAB2"] != System.DBNull.Value)
                    strSql += r2["PRECIAB2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIAB3"] != System.DBNull.Value)
                    strSql += r2["PRECIAB3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPB1"] != System.DBNull.Value)
                    strSql += r2["PRECIPB1"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPB2"] != System.DBNull.Value)
                    strSql += r2["PRECIPB2"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIPB3"] != System.DBNull.Value)
                    strSql += r2["PRECIPB3"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPOL"] != System.DBNull.Value)
                    strSql += r2["VEXCEPOL"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPOV"] != System.DBNull.Value)
                    strSql += r2["VEXCEPOV"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["VEXCEPOP"] != System.DBNull.Value)
                    strSql += r2["VEXCEPOP"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["PRECIDHL"] != System.DBNull.Value)
                    strSql += "'" + r2["PRECIDHL"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["PRECIDHV"] != System.DBNull.Value)
                    strSql += "'" + r2["PRECIDHV"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["PRECIDHP"] != System.DBNull.Value)
                    strSql += "'" + r2["PRECIDHP"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["INDEMPRE"] != System.DBNull.Value)
                    strSql += "'" + r2["INDEMPRE"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CUPSREE"] != System.DBNull.Value)
                    strSql += "'" + r2["CUPSREE"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CFACTREC"] != System.DBNull.Value)
                    strSql += "'" + r2["CFACTREC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CLINNEG"] != System.DBNull.Value)
                    strSql += r2["CLINNEG"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["CCOMAUTO"] != System.DBNull.Value)
                    strSql += "'" + r2["CCOMAUTO"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CPROVINC"] != System.DBNull.Value)
                    strSql += "'" + r2["CPROVINC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CSEGMERC"] != System.DBNull.Value)
                    strSql += "'" + r2["CSEGMERC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["NUMLABOR"] != System.DBNull.Value)
                    strSql += r2["NUMLABOR"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["TIPONEGOCIO"] != System.DBNull.Value)
                    strSql += "'" + r2["TIPONEGOCIO"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["SUPERVALLE"] != System.DBNull.Value)
                    strSql += r2["SUPERVALLE"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["ENERG_PONTA"] != System.DBNull.Value)
                    strSql += r2["ENERG_PONTA"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["HORAS_PONTA"] != System.DBNull.Value)
                    strSql += r2["HORAS_PONTA"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                if (r2["POT_MAXIMA"] != System.DBNull.Value)
                    strSql += r2["POT_MAXIMA"].ToString().Replace(",", ".") + ", ";
                else
                    strSql += " null,";

                strSql += "'" + "FACTURA SUMADA POR PERIODOS PARTIDOS" + "');";

                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private string Busca_Periodo_Factura(long creferen, int secfactu)
        {
            string strSql = "";

            strSql = "select * from 13_facturas f where"
                + " f.creferen = " + creferen + " and"
                + " f.secfactu <> " + secfactu + ";";

            return strSql;
        }
        private void BorraGrupoFacturas(long creferen, int secfactu, int i)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                Console.WriteLine(i + " Borrando factura con CREFEREN: " + creferen + " y SECFACTU: " + secfactu);
                strSql = "delete from 13_facturas_aux where"
                    + " creferen = " + creferen + " and"
                    + " secfactu = " + secfactu + ";";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private void TrataHolanda(DateTime f_factura_des, DateTime f_factura_has)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                Console.WriteLine("Tratando facturas de Holanda.");

                strSql = "replace into 13_facturas select * from fo f where"
                    + " f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd")
                    + "' and f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd")
                    + "' and f.CNIFDNIC like 'NL%'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("     Actualizando CUPSREE a través del comentario.");
                strSql = "update 13_facturas set CUPSREE = concat(substr(COMENTARIO_REFACT, 3, 2), substr(COMENTARIO_REFACT, 8, 16))"
                    + " where CNIFDNIC like 'NL%';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //strSql = "update 13_facturas b"
                //+ " inner join relacion_cups r on"
                //+ " substr(r.CUPS20, 1, 20) = substr(b.CUPSREE, 1, 18)"
                //+ " set b.CCOUNIPS = r.CUPS_CORTO, b.CUPSREE = r.CUPS20;";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void Quita_Grupo_Menor()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string nif = "";

            double sumatorio_1_2 = 0;
            double sumatorio_6 = 0;
            try
            {
                strSql = "select f.CNIFDNIC from 13_facturas f group by f.CNIFDNIC";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    nif = r["CNIFDNIC"].ToString();

                    sumatorio_1_2 = Sumatario(nif, false);
                    sumatorio_6 = Sumatario(nif, true);

                    Console.WriteLine("NIF:  " + nif + " sumatorio_1_2: " + sumatorio_1_2 + " sumatorio_6: " + sumatorio_6);

                    if (sumatorio_1_2 > sumatorio_6 && sumatorio_6 > 0)
                        BorraGrupoFacturas(nif, true);

                    if (sumatorio_6 > sumatorio_1_2 && sumatorio_1_2 > 0)
                        BorraGrupoFacturas(nif, false);

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void BorraGrupoFacturas(string nif, bool agrupadas)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "delete from 13_facturas where CNIFDNIC = '" + nif + "' and";
            if (agrupadas)
                strSql += " TFACTURA = 6;";
            else
                strSql += " TFACTURA in (1,2);";

            Console.WriteLine("Borramos --> " + strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private Double Sumatario(string nif, bool agrupadas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            double sumatorio = 0;
            try
            {
                strSql = "select sum(f.IFACTURA) as totalFactura from 13_facturas f where"
                    + " f.CNIFDNIC = '" + nif + "' and";
                if (agrupadas)
                    strSql += " (f.TFACTURA = 6 or f.TACTURA = 5);";
                else
                    strSql += " f.TFACTURA in (1,2);";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["totalFactura"] != System.DBNull.Value)
                        sumatorio = Convert.ToDouble(r["totalFactura"]);
                }
                db.CloseConnection();

                return sumatorio;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return sumatorio;
            }
        }

        public void BorradoFacturas(string testfact)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            int cemptitu = 0;
            long creferen = 0;
            int secfactu = 0;

            try
            {
                Console.WriteLine("Borrando facturas de tipo " + testfact);
                strSql = "select f.CEMPTITU, f.CREFEREN, f.SECFACTU from fact.13_facturas f where"
                    + " f.TESTFACT = '" + testfact + "';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    cemptitu = Convert.ToInt32(r["CEMPTITU"]);
                    creferen = Convert.ToInt64(r["CREFEREN"]);
                    secfactu = Convert.ToInt32(r["SECFACTU"]);
                    BorraFactura(cemptitu, creferen, secfactu);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void BorraFactura(int cemptitu, long creferen, int secfactu)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "delete from 13_facturas where"
                + " CEMPTITU = " + cemptitu + " and"
                + " CREFEREN = " + creferen + " and"
                + " SECFACTU = " + secfactu + ";";

            Console.WriteLine("BorraFactura: " + strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
    }
}
