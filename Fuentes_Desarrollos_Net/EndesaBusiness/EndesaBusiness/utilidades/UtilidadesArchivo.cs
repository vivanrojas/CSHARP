using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class UtilidadesArchivo
    {

        public UtilidadesArchivo()
        {

        }


        public int LineasArchivo(FileInfo fichero, int lineasCabecera)
        {
            
            string line = "";
            int numLine = 0;
            int totalLineas = 0;

            System.IO.StreamReader archivo = new System.IO.StreamReader(fichero.FullName);
            while ((line = archivo.ReadLine()) != null)
            {
                numLine++;
                if (numLine > lineasCabecera)
                    if (line.Trim() != "")
                        totalLineas++;
               
            }
            archivo.Close();
            return totalLineas;
        }

    }
}
