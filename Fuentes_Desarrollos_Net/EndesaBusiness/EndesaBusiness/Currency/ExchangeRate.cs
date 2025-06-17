using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.Currency
{
    public class ExchangeRate
    {
        utilidades.Param p = new utilidades.Param("fact_p_exchange_param", MySQLDB.Esquemas.FAC);
        public bool necesitaCredenciales { get; set; }
        public ExchangeRate()
        {
            string[] listaDominios;
            listaDominios = p.GetValue("domains", DateTime.Now, DateTime.Now).Split(',');

            for (int i = 0; i < listaDominios.Count(); i++)
            {
                if (System.Environment.UserDomainName == listaDominios[i])
                {
                    necesitaCredenciales = true;
                    break;
                }
            }
        }

        public void DescargaEuroDolar(string user, string password, string domain)
        {
            bool necesitaCredenciales = false;            
            StringBuilder sb = new StringBuilder();
            string fileDestination = "";
            string fileSource = "";

            //bool necesitaCredenciales = false;
           
            WebClient client = new WebClient();

            try
            {               

                if (necesitaCredenciales)
                {
                    WebRequest.DefaultWebProxy.Credentials = new NetworkCredential(user, password, domain);
                    client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.33 Safari/537.36");


                    IWebProxy proxy = new WebProxy(p.GetValue("proxy", DateTime.Now, DateTime.Now),
                        Convert.ToInt32(p.GetValue("proxy_port", DateTime.Now,DateTime.Now)));

                    proxy.Credentials = new NetworkCredential(user, password, domain);
                    client.Proxy = proxy;
                    client.Credentials = new NetworkCredential(user, password, domain);
                }

                fileSource = p.GetValue("url_euro_dolar_exchange", DateTime.Now, DateTime.Now);
                fileDestination = p.GetValue("file_euro_dolar_destination", DateTime.Now,DateTime.Now);
                client.DownloadFile(fileSource, fileDestination);
                this.CargaArchivoEuroDolar(fileDestination);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + System.Environment.NewLine +
                 "Error en la descarga del fichero de relacción EUR-USD." + System.Environment.NewLine +
                 "No se ha podido accedera a la dirección: " + p.GetValue("url_euro_dolar_exchange", DateTime.Now, DateTime.Now)
                 , "Error en la descarga del archivo",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }



        }

        private void CargaArchivoEuroDolar(string archivo)
        {
            long i = 0;
            string line;
            StringBuilder sb = new StringBuilder();
            Boolean firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            try
            {
                System.IO.StreamReader file = new System.IO.StreamReader(archivo);
                while ((line = file.ReadLine()) != null)
                {
                    i++;
                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO fact_p_exchange_rates (");
                        sb.Append("time_period, currency, value) values ");
                        firstOnly = false;
                    }

                    if (line.Contains("<Obs TIME_PERIOD="))
                    {
                        sb.Append("('").Append(line.Substring(18, 10)).Append("', ");
                        sb.Append("'USD', ");
                        sb.Append(line.Substring(41, 6)).Append("),");
                    }

                    if (i == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                file.Close();

                MessageBox.Show("La actualización de los tipos de cambios se ha realizado correctamente.",
               "Actualización de tipos de cambio EUR / USD",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la importación del fichero de relacción EUR-USD.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }


        }
    }
}
