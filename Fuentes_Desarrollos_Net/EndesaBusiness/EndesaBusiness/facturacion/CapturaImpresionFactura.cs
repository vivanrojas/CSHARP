using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class CapturaImpresionFactura
    {
        InventarioFacturacionPortugal inventario;
        logs.Log ficheroLog;
        utilidades.Param p;
        public CapturaImpresionFactura()
        {
            DateTime ultimoDiaHabil = new DateTime();
            DateTime primerDia = new DateTime();
            string[] listaArchivos;
            string extractor;

            

            try
            {
                ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_BotesImpresion");
                p = new utilidades.Param("ag_pt_param", MySQLDB.Esquemas.FAC);
                inventario = new InventarioFacturacionPortugal(null, DateTime.Now, DateTime.Now);

                ultimoDiaHabil = utilidades.Fichero.UltimoDiaHabil();
                primerDia = new DateTime(ultimoDiaHabil.Year, ultimoDiaHabil.Month, 1);
                extractor = p.GetValue("Extractor_Bote", DateTime.Now, DateTime.Now);


                for (DateTime d = ultimoDiaHabil; d > primerDia; d = d.AddDays(-1))
                {
                    ficheroLog.Add("Ejecutando extractor: " + extractor + " --> " + d.ToString("yyMMdd"));
                    // utilidades.Fichero.EjecutaComando(extractor, d.ToString("yyMMdd"));
                }

                string prefijoArchivo = p.GetValue("prefijo_bote", DateTime.Now, DateTime.Now) + "*.txt";
                string rutaOrigen = p.GetValue("inbox", DateTime.Now, DateTime.Now);

                listaArchivos = Directory.GetFiles(rutaOrigen, prefijoArchivo);
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    ficheroLog.Add("Procesando " + listaArchivos[i]);
                    CargaArchivo(listaArchivos[i]);
                    // file = new FileInfo(listaArchivos[i]);
                    // inicio = DateTime.Now;
                    // file = new FileInfo(listaArchivos[i]);
                    // Console.WriteLine("Procensado " + listaArchivos[i]);
                    // CargaInffactPorLinea(listaArchivos[i], 20, "CEFACO", 3);
                    // CargaInffactPorLinea(listaArchivos[i], 20, "OPERACIONES", 3);
                    // fin = DateTime.Now;
                    // global.SaveProcess("Inffact", "Procesado correctamente", inicio, inicio, fin);
                    // file.Delete();
                    // Seguimiento factura BTN

                }

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            
        }

        private void CargaArchivo(string fileName)
        {
            string line = "";

            bool firstOnly = true;

            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            int i = 0;
            int j = 0;

            System.IO.StreamReader archivo;
            FileInfo file;

            string fecha = "";

            try
            {
                file = new FileInfo(fileName);
                Console.WriteLine("Reading " + fileName);
                archivo = new System.IO.StreamReader(fileName, System.Text.Encoding.GetEncoding(1252));
                while ((line = archivo.ReadLine()) != null)
                {
                    j++;
                    Console.CursorLeft = 0;
                    Console.Write("Leyendo " + j + " lineas...");

                    if (inventario.Existe(line.Substring(10406, 20)))
                    {
                        i++;
                        if (firstOnly)
                        {
                            sb.Append("replace into ag_pt_imp (cpe, cfactura, ");
                            sb.Append("ffactura, ffactdes, ffacthas, ");
                            sb.Append("ifactura, p1, p2, p3, p4, ");
                            sb.Append("tr1, tr2, tr3, cosphi, tasa_aud_t, archivo) values ");
                            firstOnly = false;
                        }

                        sb.Append("('").Append(line.Substring(10406, 20)).Append("',"); // CPE
                        sb.Append("'").Append(line.Substring(33, 14).Trim()).Append("',"); // FACTURA

                        fecha = line.Substring(101, 10).Trim();
                        fecha = fecha.Substring(6, 4) + "-"
                            + fecha.Substring(3, 2) + "-"
                            + fecha.Substring(0, 2);
                                               

                        sb.Append("'").Append(fecha).Append("',"); // FFACTURA

                        fecha = line.Substring(161, 10).Trim();
                        fecha = fecha.Substring(6, 4) + "-"
                            + fecha.Substring(3, 2) + "-"
                            + fecha.Substring(0, 2);

                        sb.Append("'").Append(fecha).Append("',"); // FFACTDES

                        fecha = line.Substring(171, 10).Trim();
                        fecha = fecha.Substring(6, 4) + "-"
                            + fecha.Substring(3, 2) + "-"
                            + fecha.Substring(0, 2);

                        sb.Append("'").Append(fecha).Append("',"); // FFACTHAS

                        sb.Append(FuncionesTexto.CN(line.Substring(111, 14))).Append(","); // IFACTURA
                        sb.Append(FuncionesTexto.CS(line.Substring(737, 120).Trim())).Append(","); // P1
                        sb.Append(FuncionesTexto.CS(line.Substring(857, 120).Trim())).Append(","); // P2
                        sb.Append(FuncionesTexto.CS(line.Substring(977, 120).Trim())).Append(","); // P3
                        sb.Append(FuncionesTexto.CS(line.Substring(1097, 55).Trim())).Append(","); // P4

                        if (FuncionesTexto.CS(line.Substring(1817, 120)).Contains("TR1"))
                            sb.Append(FuncionesTexto.CS(line.Substring(1817, 120))).Append(","); // TR1
                        else
                            sb.Append("null,");

                        if (FuncionesTexto.CS(line.Substring(1817, 120)).Contains("TR2"))
                            sb.Append(FuncionesTexto.CS(line.Substring(1817, 120))).Append(","); // TR2
                        else
                            sb.Append("null,");

                        if (FuncionesTexto.CS(line.Substring(1817, 120)).Contains("TR3"))
                            sb.Append(FuncionesTexto.CS(line.Substring(2057, 120))).Append(","); // TR3
                        else
                            sb.Append("null,");                        

                        if (FuncionesTexto.CS(line.Substring(2177, 80)).Contains("COS PHI"))
                            sb.Append(FuncionesTexto.CS(line.Substring(2177, 80))).Append(","); // COSPHI
                        else
                            sb.Append("null,");

                        // sb.Append(FuncionesTexto.CS(line.Substring(2337, 60))).Append(","); // TASA_AUD

                        sb.Append("null,");

                        sb.Append("'").Append(file.Name).Append("'),");

                        if (i == 1)
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

                }
                archivo.Close();
                //file.Delete();

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);                
                
            }
        }
    }
}
