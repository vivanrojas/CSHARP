using EndesaBusiness.servidores;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.ocasional
{
    public class ArregloTCON
    {
        public ArregloTCON()
        {

        }

        public void Carga()
        {
            StringBuilder sb = new StringBuilder();
            Boolean firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int i = 0;
            int x = 0;
            int z = 0;
            string[] cc;
            string fecha = "";
            int j = 0;
            int tcon = 0;
            double icon = 0;
            int total_lineas = 0;

            string fichero = @"C:\Endesa\peticiones_en_curso\20200521_FAC_GO_INFFACT_IV\factu_tconfact14_2020.txt";
            string line;
            System.IO.StreamReader file = new System.IO.StreamReader(fichero, System.Text.Encoding.GetEncoding(1252));
            while ((line = file.ReadLine()) != null)
            {
                total_lineas++;
            }
            file.Close();

            Console.WriteLine("");
            file = new System.IO.StreamReader(fichero, System.Text.Encoding.GetEncoding(1252));
            while ((line = file.ReadLine()) != null)
            {

                z = 1;
                cc = line.Split(';');
                j++;
                i++;

                Console.CursorLeft = 0;
                Console.Write("Leyendo "
                       + string.Format("{0:#,##0}", j) + " de "
                       + string.Format("{0:#,##0}", total_lineas));

                if (firstOnly)
                {
                    sb.Append("replace into fo_arreglo_tcon ");
                    sb.Append("(CEMPTITU,CREFEREN,SECFACTU,CREFFACT,CLINNEG,CFACTURA,FFACTURA,FFACTDES,FFACTHAS,");
                    sb.Append("VCUOVAFA,VENEREAC,VCUOFIFA,IFACTURA,TFACTURA,TESTFACT,TCONFAC1,ICONFAC1,TCONFAC2,ICONFAC2,TCONFAC3,");
                    sb.Append("ICONFAC3,TCONFAC4,ICONFAC4,TCONFAC5,ICONFAC5,TCONFAC6,ICONFAC6,TCONFAC7,ICONFAC7,TCONFAC8,ICONFAC8,");
                    sb.Append("TCONFAC9,ICONFAC9,TCONFA10,ICONFA10,TCONFA11,ICONFA11,TCONFA12,ICONFA12,TCONFA13,ICONFA13,TCONFA14,");
                    sb.Append("ICONFA14,TCONFA15,ICONFA15,TCONFA16,ICONFA16,TCONFA17,ICONFA17,TCONFA18,ICONFA18,TCONFA19,ICONFA19,");
                    sb.Append("TCONFA20,ICONFA20,CSEGMERC) values ");
                    firstOnly = false;
                }
                
                sb.Append("('").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;   // CEMPTITU
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // CREFEREN
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // SECFACTU
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // CREFFACT
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // CLINNEG
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // CFACTURA
                fecha = utilidades.FuncionesTexto.CD(cc[z]); z++;
                sb.Append(fecha).Append(", "); // FFACTURA
                fecha = utilidades.FuncionesTexto.CD(cc[z]); z++;
                sb.Append(fecha).Append(", "); // FFACTDES
                fecha = utilidades.FuncionesTexto.CD(cc[z]); z++;
                sb.Append(fecha).Append(", "); // FFACTHAS
                sb.Append(utilidades.FuncionesTexto.CN(cc[z])).Append(", "); z++; // VCUOVAFA
                sb.Append(utilidades.FuncionesTexto.CN(cc[z])).Append(", "); z++; // VENEREAC
                sb.Append(utilidades.FuncionesTexto.CN(cc[z])).Append(", "); z++; // VCUOFIFA
                sb.Append(utilidades.FuncionesTexto.CN(cc[z])).Append(", "); z++; // IFACTURA
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // TFACTURA
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("', "); z++;    // TESTFACT

                for (int w = 0; w < 20; w++)
                {
                    tcon = Convert.ToInt32(cc[z]); z++;
                    icon = Convert.ToDouble(cc[z]); z++;

                    if (icon > 0)
                        icon = icon / 1000000000;

                    sb.Append(tcon).Append(", ");   // TCONFAC
                    sb.Append(icon.ToString().Replace(",",".")).Append(", ");  // ICONFAC
                }

                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(cc[z])).Append("'),");  // CSEGMERC

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
            file.Close();
        }


    }
}
