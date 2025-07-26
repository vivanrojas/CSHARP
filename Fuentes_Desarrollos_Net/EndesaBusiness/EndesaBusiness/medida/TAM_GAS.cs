using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class TAM_GAS
    {
        public string cups20 { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_tam_gas_endesabusiness");


        // facturacion.Impuestos imp;

        Dictionary<string, List<EndesaEntity.facturacion.TAM_factura>> dic = 
            new Dictionary<string, List<EndesaEntity.facturacion.TAM_factura>>();
        List<EndesaEntity.medida.TipoTam> l_tam = new List<EndesaEntity.medida.TipoTam>(); // Lista TAM

        EndesaBusiness.sigame.SIGAME sigame = new sigame.SIGAME();

        //Dictionary<Int32, EndesaEntity.facturacion.TAM_factura> dlf = new Dictionary<Int32, EndesaEntity.facturacion.TAM_factura>(); // Diccionario de facturas
        //Dictionary<string, EndesaEntity.medida.PuntoSuministro> dlc = new Dictionary<string, EndesaEntity.medida.PuntoSuministro>(); // Diccionario de cups

        //Dictionary<string, List<Int32>> dcf = new Dictionary<string, List<Int32>>();

        //List<EndesaEntity.facturacion.TAM_factura> lf = new List<EndesaEntity.facturacion.TAM_factura>(); // lista facturas

        //List<EndesaEntity.medida.PuntoSuministro> lc = new List<EndesaEntity.medida.PuntoSuministro>(); // lista cups

        //List<EndesaEntity.medida.PuntoSuministro> lcPS = new List<EndesaEntity.medida.PuntoSuministro>(); // lista cups PS_AT
        //List<EndesaEntity.medida.PuntoSuministro> lcmt = new List<EndesaEntity.medida.PuntoSuministro>(); // lista cups MT
        //List<string> ln = new List<string>(); // lista nifs
        //List<EndesaEntity.medida.PuntoSuministro> lista_cups_calculo_especial = new List<EndesaEntity.medida.PuntoSuministro>();


        //Dictionary<string, double> cups_tam = new Dictionary<string, double>(); // Diccionario para consultas de tam en otros procesos

        //cups.PuntosSuministro p = new cups.PuntosSuministro();

        public TAM_GAS()
        {
            fd = new DateTime();
            fh = new DateTime();

            cups20 = null;

            fh = utilidades.Fichero.UltimoDiaHabil();
            fd = fh.AddYears(-2);

            //imp = new facturacion.Impuestos();


        }

        public TAM_GAS(string _cups20)
        {
            fd = new DateTime();
            fh = new DateTime();

            cups20 = _cups20;

            fh = utilidades.Fichero.UltimoDiaHabil();
            fd = fh.AddYears(-2);

            //imp = new facturacion.Impuestos();
        }


        public void Crea_TAM()
        {
            try
            {


                Console.WriteLine(" ***** CALCULANDO TAM ***** ");
                ficheroLog.Add(" ***** CALCULANDO TAM ***** ");                
                                
                GuardaFacturas(fd, fh, cups20);

                // Calculamos el TAM
                Calcula_TAM_CUPS();

                // Calculamos Ranking
                //CreaRanking();

                // Guardamos el TAM en BBDD
                Guarda_TAM_CUPS();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError(e.Message);
            }
        }

        //private void CreaRanking()
        //{
        //    List<EndesaEntity.medida.TipoTam> t1 = new List<EndesaEntity.medida.TipoTam>();
        //    List<EndesaEntity.medida.TipoTam> t2 = new List<EndesaEntity.medida.TipoTam>();
        //    int ranking = 0;

        //    try
        //    {

        //        // Nos quedamos con los puntos PS_AT
        //        t1 = l_tam.FindAll(z => z.esPS_AT);
        //        // Ordemos por TAM
        //        t2 = t1.OrderByDescending(z => z.tam).ToList();

        //        foreach (EndesaEntity.medida.TipoTam t in t2)
        //        {
        //            ranking++;
        //            l_tam[l_tam.FindIndex(z => z.cups20 == t.cups20)].rankingTAM = ranking;
        //        }
        //        // Ordenamos por factura Maxima
        //        t2.Clear();
        //        t2 = t1.OrderByDescending(z => z.importeMaximo).ToList();
        //        ranking = 0;
        //        foreach (EndesaEntity.medida.TipoTam t in t2)
        //        {
        //            ranking++;
        //            l_tam[l_tam.FindIndex(z => z.cups20 == t.cups20)].rankingImporteMaximo = ranking;
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        ficheroLog.AddError(e.Message);
        //    }

        //}
                
        private void GuardaFacturas(DateTime fd, DateTime fh, string cups20)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            Int32 i = 0;
            

            try
            {

                Console.WriteLine("Buscando Facturas para analizar ...");

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select f.CEMPTITU, f.CNIFDNIC, DATE_FORMAT(f.FFACTDES,'%Y%m') as mes_hasta,"
                    + " SUM(f.IFACTURA) as IFACTURA, f.FFACTURA, DATEDIFF(f.FFACTURA, f.FFACTHAS) as diaFact,"
                    + " f.DAPERSOC, f.CUPSREE, f.FFACTDES, f.FFACTHAS,"
                    + " f.CREFEREN, f.SECFACTU, f.TESTFACT"
                    + " from fact.fo f"
                    + " where";

                if (cups20 != null)
                {
                    strSql = strSql + " f.CUPSREE = '" + cups20 + "' and";
                }
                strSql = strSql + " ((f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "')"
                + " and f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')"                 
                + " and f.TIPONEGOCIO = 'G'"
                + " AND f.FFACTURA > f.FFACTHAS"
                + " GROUP BY f.CUPSREE, YEAR(f.FFACTDES), MONTH(f.FFACTDES)"
                + " order by f.CCOUNIPS, YEAR(f.FFACTDES) desc, MONTH(f.FFACTDES) desc;";


                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                Console.WriteLine("Recopilando facturas ...");

                while (reader.Read())
                {
                    i++;
                    EndesaEntity.facturacion.TAM_factura t = new EndesaEntity.facturacion.TAM_factura();

                    t.cemptitu = Convert.ToInt32(reader["CEMPTITU"]);                            
                    t.cnifdnic = reader["CNIFDNIC"].ToString();
                    t.mes_hasta = Convert.ToInt32(reader["mes_hasta"]);
                    t.ifactura = Convert.ToDouble(reader["IFACTURA"]);
                    t.diaFact = Convert.ToInt32(reader["diaFact"]);
                    t.dapersoc = reader["DAPERSOC"].ToString();
                    t.creferen = Convert.ToInt64(reader["CREFEREN"]);
                    t.secfactu = Convert.ToInt32(reader["SECFACTU"]);
                    t.testfact = reader["TESTFACT"].ToString();

                    

                    if (reader["FFACTURA"] != System.DBNull.Value)
                        t.ffactura = Convert.ToDateTime(reader["FFACTURA"]);

                    t.ffactdes = Convert.ToDateTime(reader["FFACTDES"]);
                    t.ffacthas = Convert.ToDateTime(reader["FFACTHAS"]);

                    // lf.Add(t);
                    // var values = dcf.Where(z => z.Key == t.ccounips).Select(z => z.Value);
                    List<EndesaEntity.facturacion.TAM_factura> o;

                    if (reader["CUPSREE"] != System.DBNull.Value)
                    {
                        t.cupree = reader["CUPSREE"].ToString();
                        if (sigame.EsVigente(t.cupree))
                        {
                            if (!dic.TryGetValue(t.cupree, out o))
                            {
                                o = new List<EndesaEntity.facturacion.TAM_factura>();
                                o.Add(t);
                                dic.Add(t.cupree, o);
                            }
                            else
                                o.Add(t);
                        }

                    }
                        
                }

                Console.WriteLine("Total Facturas a analizar: " + i);
                ficheroLog.Add("Total facturas a analizar: " + i);

                db.CloseConnection();

                // BuscaComplementosFacturasEspeciales();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError(e.Message);
            }
        }
        
        

        
        
        private void Calcula_TAM_CUPS()
        {

            EndesaEntity.facturacion.TAM_factura factura = new EndesaEntity.facturacion.TAM_factura();

            int numeroFacturas = 0;
            int numCups = 0;
            int i = 0;
            int diaFactMax = 0;
            double diaFactMed = 0;
            double importeTotal = 0;
            double factMax = 0;
            int numFacturasTratadas = 0;
            string cups20 = "";
            int totalDias = 0;

            try
            {
                ficheroLog.Add("Calculando TAM para " + dic.Count() + " cups ");

                foreach (KeyValuePair<string, List<EndesaEntity.facturacion.TAM_factura>> p in dic)
                {
                    numCups++;
                    i++;
                    importeTotal = 0;
                    totalDias = 0;
                    factMax = 0;
                    diaFactMed = 0;
                    diaFactMax = 0;
                    bool hayFacturas = false;



                    DateTime fechaHastaMaxima = new DateTime();
                    cups20 = p.Key;

                    numFacturasTratadas = 0;

                    fechaHastaMaxima = new DateTime(2010, 1, 1);                    

                    numeroFacturas = p.Value.Count();

                    hayFacturas = false;

                    for (int j = 0; j < p.Value.Count(); j++)
                    {
                        hayFacturas = true;
                        factura = p.Value[j];                       
                        
                        if (fechaHastaMaxima < factura.ffacthas)
                            fechaHastaMaxima = factura.ffacthas;

                        numFacturasTratadas++;

                        if (numFacturasTratadas > numeroFacturas)
                        {
                            break;
                        }

                            
                        importeTotal = importeTotal + factura.ifactura;
                        totalDias = totalDias + ((factura.ffacthas.Day - factura.ffactdes.Day) + 1);

                        // Guardamos el importe maximo
                        if (factMax < factura.ifactura)
                        {
                            factMax = factura.ifactura;
                        }
                        // Guardamos el dia maximo de emision de la factura
                        if (diaFactMax < factura.diaFact)
                        {
                            diaFactMax = factura.diaFact;
                        }

                        diaFactMed = diaFactMed + factura.diaFact;

                        

                    } 


                    if (hayFacturas)
                    {
                        if (importeTotal != 0) 
                        {
                            importeTotal = Math.Round((importeTotal / numFacturasTratadas), 2);
                        }

                        if (diaFactMed != 0 && numeroFacturas != 0)
                        {
                            diaFactMed = Math.Round((diaFactMed / numeroFacturas), 0);
                        }

                        EndesaEntity.medida.TipoTam tam = new EndesaEntity.medida.TipoTam();
                        tam.cemptitu = factura.cemptitu;
                        tam.cnifdnic = factura.cnifdnic;
                        tam.dapersoc = factura.dapersoc;                        
                        tam.cups20 = (factura.cupree != null ? factura.cupree : "");
                        tam.tam = importeTotal;
                        tam.numfacturasCalculo = numeroFacturas;
                        tam.importeMaximo = factMax;
                        tam.fechaHastaMaxima = fechaHastaMaxima;
                        tam.diasFacturacionMedia = Convert.ToInt32(diaFactMed);
                        tam.diasFacturacionMaxima = Convert.ToInt32(diaFactMax);
                                                

                        l_tam.Add(tam);

                        Console.Write("*");
                        i = 0;

                    } // if (hayFacturas)

                } // foreach (KeyValuePair<string, cups.PuntoSuministro> p in dlc)

                Console.WriteLine("");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + " " + cups20);
                ficheroLog.AddError("Calcula_TAM_CUPS: " + e.Message);
            }
        }

        private void Guarda_TAM_CUPS()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            StringBuilder sb = new StringBuilder();
            Boolean firstOnly = true;

            int numReg = 0;
            int datosTratados = 0;

            try
            {

                if(l_tam.Count > 0)
                {
                    strSql = "replace into tam_g_hist select * from tam_g";
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    strSql = "delete from tam_g";
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                }


                foreach (EndesaEntity.medida.TipoTam tam in l_tam)
                {
                    numReg++;
                    datosTratados++;

                    if (firstOnly)
                    {
                        sb.Append("replace into tam_g (CEMPTITU, CNIFDNIC, DAPERSOC, CUPS20, TAM,");
                        sb.Append("NumeroFacturasCalculo, ImporteMaximo, DiasFacturacionMedia, DiasFacturacionMaxima,");
                        sb.Append("FechaHastaMaxima, RankingTAM, RankingImporteMaximo) values");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(tam.cemptitu).Append("',");
                    sb.Append("'").Append(tam.cnifdnic).Append("',");
                    sb.Append("'").Append(tam.dapersoc).Append("',");                    
                    sb.Append("'").Append(tam.cups20).Append("',");
                    sb.Append(tam.tam.ToString().Replace(",", ".")).Append(","); // TAM
                    sb.Append(tam.numfacturasCalculo).Append(",");
                    sb.Append(tam.importeMaximo.ToString().Replace(",", ".")).Append(",");
                    sb.Append(tam.diasFacturacionMedia).Append(",");
                    sb.Append(tam.diasFacturacionMaxima).Append(",");
                    sb.Append("'").Append(tam.fechaHastaMaxima.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append(tam.rankingTAM).Append(",");
                    sb.Append(tam.rankingImporteMaximo).Append("),");

                    if (numReg == 500)
                    {
                        ficheroLog.Add("Guardando datos. " + (datosTratados + 1) + " datos guardados ...");
                        Console.WriteLine("Guardando datos. " + (datosTratados + 1) + " datos guardados ...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }
                }

                if (numReg > 0)
                {
                    Console.WriteLine("Guardando datos. " + (datosTratados + 1) + " datos guardados ...");
                    ficheroLog.Add("Guardando datos. " + (datosTratados + 1) + " datos guardados ...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    numReg = 0;
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

        }

        //private decimal TamEntregasACuenta(EndesaEntity.facturacion.TAM_factura f)
        //{
        //    EndesaEntity.facturacion.TAM_factura_detalle ff = new EndesaEntity.facturacion.TAM_factura_detalle();

        //    // decimal acumulado = 0;
        //    // decimal impuestoElectrico;
        //    // decimal reduccionISE;

        //    // 1.- Solo tratamos las facturas con el complemento 276 "DESCUENTO FACTURACION A CUENTA"
        //    if (f.fd.Exists(z => z.tconfac == 276))
        //    {
        //        //impuestoElectrico = decimal.Multiply(imp.GetValue("IMPElectrico", f.ffactura, f.ffactura), 
        //        //    imp.GetValue("IMPElectrico2", f.ffactura, f.ffactura));
        //        //impuestoElectrico = decimal.Divide(impuestoElectrico, 100);
        //        //ff = f.fd.Find(z => z.tconfac == 601);

        //        //acumulado = acumulado + decimal.Multiply(ff.iconfac, impuestoElectrico);
        //        ////acumulado = 
        //    }
        //    return 0;
        //}

   

        private int NumFacturas(List<EndesaEntity.facturacion.TAM_factura> lista_facturas)
        {
            int num = 0;

            num = lista_facturas.Count();

            if (num > 12)
            {
                return 12;
            }
            else
            {
                return num;
            }

        }
        
        private string GetParameter(string codigo, DateTime f)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            string vcodigo;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "Select value from fact.tam_param where"
                    + " code = '" + codigo + "' and"
                    + " (from_date <= '" + f.ToString("yyyy-MM-dd") + "' and"
                    + " (to_date >= '" + f.ToString("yyyy-MM-dd") + "' or to_date is null))";

                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    vcodigo = reader["value"].ToString();
                }
                else
                {
                    Console.WriteLine("El valor " + codigo + " no está parametrizado en fact_p_exchange_param.");
                    vcodigo = "";
                }
                db.CloseConnection();
                return vcodigo;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return "";

        }

       

        

        

        //public void CargaTAM()
        //{
        //    MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql = "";

        //    try
        //    {
        //        strSql = "select t.CUPS13, t.TAM from tam t where t.CUPS13 is not null and trim(t.CUPS13) <> ''"
        //            + " group by t.CUPS13;";
        //        db = new MySQLDB(MySQLDB.Esquemas.FAC);
        //        command = new MySqlCommand(strSql, db.con);
        //        r = command.ExecuteReader();
        //        while (r.Read())
        //        {
        //            cups_tam.Add(r["CUPS13"].ToString(), Convert.ToDouble(r["TAM"]));
        //        }
        //        db.CloseConnection();

        //    }
        //    catch (Exception e)
        //    {
        //        ficheroLog.AddError("CargaTAM: " + e.Message);
        //    }
        //}

        //public double GetTAM(string cups13)
        //{
        //    return cups_tam.FirstOrDefault(z => z.Key == cups13).Value;
        //}
    }
}
