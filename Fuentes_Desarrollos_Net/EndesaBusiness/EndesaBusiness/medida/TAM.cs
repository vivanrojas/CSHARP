using EndesaBusiness.eer;
using EndesaBusiness.servidores;
using Microsoft.Exchange.WebServices.Data;
using MySql.Data.MySqlClient;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System;
using System.Collections.Generic;
using System.IdentityModel.Protocols.WSTrust;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class TAM
    {
        public string cups13 { get; set; }
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }


        Dictionary<string, double> cups_tam = new Dictionary<string, double>(); // Diccionario para consultas de tam en otros procesos
        Dictionary<string, double> nif_cups_tam = new Dictionary<string, double>(); // Diccionario para consultas de tam en otros procesos

        cups.PuntosSuministro p = new cups.PuntosSuministro();

        logs.Log ficheroLog;

        public TAM()
        {

            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_tam_facturas");

        }




        public void Crea_TAM_VIEJO()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            DateTime fh = DateTime.Now;
            fh = fh.AddMonths(-1);
            fh = new DateTime(fh.Year, fh.Month, DateTime.DaysInMonth(fh.Year, fh.Month));

            Dictionary<string, List<EndesaEntity.facturacion.TAM_factura>> d =
                new Dictionary<string, List<EndesaEntity.facturacion.TAM_factura>>();

            List<EndesaEntity.facturacion.TAM_factura> lista_tam = new List<EndesaEntity.facturacion.TAM_factura>();

            double importe = 0;
            DateTime min_ffactdes = new DateTime();
            DateTime max_ffacthas = new DateTime();

            StringBuilder sb = new StringBuilder();

            bool firstOnly = true;

            int numReg = 0;
            int datosTratados = 0;

            try
            {
                strSql = "SELECT t.CUPSREE, t.FFACTDES, t.FFACTHAS, t.IFACTURA" 
                    + " FROM fo_tam_consumos_meses_2 t"
                    + " WHERE"
                    // + " t.CUPSREE = 'ES0031300316474001BJ0F' and"
                    + " t.FFACTDES > t.FFACTDES - INTERVAL 12 MONTH"
                    + " ORDER BY t.CUPSREE, t.FFACTDES DESC";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.TAM_factura c = new EndesaEntity.facturacion.TAM_factura();
                    c.cupree = r["CUPSREE"].ToString();
                    c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);
                    c.ifactura = Convert.ToDouble(r["IFACTURA"]);


                    List<EndesaEntity.facturacion.TAM_factura> lista;
                    if (!d.TryGetValue(c.cupree, out lista))
                    {
                        lista = new List<EndesaEntity.facturacion.TAM_factura>();
                        lista.Add(c);
                        d.Add(c.cupree, lista);
                    }
                    else
                        lista.Add(c);
                                       

                }
                r.Close();
                db.CloseConnection();



                foreach(KeyValuePair<string, List<EndesaEntity.facturacion.TAM_factura>> p in d)
                {
                    importe = 0;
                    min_ffactdes = new DateTime(4999, 12, 31);
                    max_ffacthas = DateTime.MinValue;

                    foreach(EndesaEntity.facturacion.TAM_factura pp in p.Value)
                    {
                        importe = importe + pp.ifactura;

                        if(pp.ffactdes < min_ffactdes)
                            min_ffactdes = pp.ffactdes;

                        if(pp.ffacthas > max_ffacthas)
                            max_ffacthas = pp.ffacthas;
                    }

                    EndesaEntity.facturacion.TAM_factura c = new EndesaEntity.facturacion.TAM_factura();
                    c.cupree = p.Key;
                    c.ffactdes = min_ffactdes;
                    c.ffacthas = max_ffacthas;
                    c.ifactura = Math.Round(importe / p.Value.Count,2);

                    lista_tam.Add(c);
                }

                // Guardamos los datos en BBDD

                foreach(EndesaEntity.facturacion.TAM_factura p in lista_tam)
                {
                    numReg++;
                    datosTratados++;

                    if (firstOnly)
                    {
                        sb.Append("replace into fo_tam (cups20, ffactdes, ffacthas, tam");                        
                        sb.Append(") values");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.cupree).Append("',");
                    sb.Append("'").Append(p.ffactdes.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(p.ffacthas.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append(p.ifactura.ToString().Replace(",", "."));
                    sb.Append("),");


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
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            
            
        }

       

        public void CargaTAM()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select cups20, tam from fo_tam ";
                                 
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.TipoTam c = new EndesaEntity.medida.TipoTam();
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.cups20 = r["cups20"].ToString();
                    c.tam = Convert.ToDouble(r["tam"]);

                    double o;
                    if (!cups_tam.TryGetValue(c.cups20, out o))
                        cups_tam.Add(c.cups20, c.tam);
                }
                r.Close();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void CargaTAM_NIF_CUPS()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select CNIFDNIC, cups20, tam from tam "
                    + " order by F_ULT_MOD desc";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.TipoTam c = new EndesaEntity.medida.TipoTam();
                    c.cnifdnic = r["CNIFDNIC"].ToString();
                    c.cups20 = r["cups20"].ToString();
                    c.tam = Convert.ToDouble(r["tam"]);

                    double o;
                    if (!cups_tam.TryGetValue(c.cnifdnic + "_" + c.cups20, out o))
                        cups_tam.Add(c.cnifdnic + "_" + c.cups20, c.tam);
                }
                r.Close();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public double GetTAM(string cups20)
        {
            double tam;
            double importe_tam = 0;
            if (cups_tam.TryGetValue(cups20, out tam))
                importe_tam = tam;
            else
                importe_tam = -1111111;

            return importe_tam;
        }

        public double GetTAM(string nif, string cups20)
        {
            double tam;
            double importe_tam = 0;
            if (cups_tam.TryGetValue(nif + "_" + cups20, out tam))
                importe_tam = tam;            

            return importe_tam;
        }


    }
}
