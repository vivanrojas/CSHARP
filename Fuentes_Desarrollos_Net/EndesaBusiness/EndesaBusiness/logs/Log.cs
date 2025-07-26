using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.logs
{
    public class Log
    {
        #region Private Attributes
        private string folder = "";
        private string path = "";
        private string prefijoLog = "";
        protected string BRINCO = "\n";
        protected List<string> lErrores = new List<string>();

        #endregion
        #region Propierties
        public string FullPath
        {
            get { return path + "/" + folder; }
        }
        #endregion

        public Log(string path, string folder_, string prefijoLog_)
        {
            folder = folder_;
            this.path = path;
            this.prefijoLog = prefijoLog_;
        }


        public bool Add(string sLog)
        {
            try
            {
                clearErrores();

                //verificamos si existe directorio
                if (!CreateDirectory()) return false;

                string nombreArchivo = GetNameFile();//obtenemos nombre de archivo del día
                string cadena = ""; //obtenemos el contenido del archivo


                cadena += DateTime.Now + " - " + System.Environment.UserName + " - " + sLog + Environment.NewLine;

                //creamos el archivo y guardamos
                StreamWriter sw = new StreamWriter(FullPath + "/" + nombreArchivo, true);
                sw.Write(cadena);
                sw.Close();

                return true;

            }
            catch (DirectoryNotFoundException ex)
            {
                addError(ex.Message + " Directorio invalido: " + FullPath);
                return false;
            }
            catch (FileNotFoundException ex)
            {
                addError(ex.Message + " Ruta de archivo invalida");
                return false;
            }
            catch (Exception ex)
            {

                addError("Error: " + ex.Message);
                return false;
            }
        }

        public bool AddError(string sLog)
        {
            try
            {
                clearErrores();

                //verificamos si existe directorio
                if (!CreateDirectory()) return false;

                string nombreArchivo = GetNameFileError();//obtenemos nombre de archivo del día
                string cadena = ""; //obtenemos el contenido del archivo


                cadena += DateTime.Now + " - " + System.Environment.UserName + " - " + sLog + Environment.NewLine;

                //creamos el archivo y guardamos
                StreamWriter sw = new StreamWriter(FullPath + "/" + nombreArchivo, true);
                sw.Write(cadena);
                sw.Close();

                return true;

            }
            catch (DirectoryNotFoundException ex)
            {
                addError(ex.Message + " Directorio invalido: " + FullPath);
                return false;
            }
            catch (FileNotFoundException ex)
            {
                addError(ex.Message + " Ruta de archivo invalida");
                return false;
            }
            catch (Exception ex)
            {

                addError("Error: " + ex.Message);
                return false;
            }
        }

        #region Helpers

        private string GetNameFile()
        {
            string nombre = "";
            nombre = this.prefijoLog + "_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            //nombre = "log_" + DateTime.Now.Day + "_" + DateTime.Now.Month + "_" + DateTime.Now.Year + ".txt";

            return nombre;
        }

        private string GetNameFileError()
        {
            string nombre = "";
            nombre = this.prefijoLog + "_ERROR_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            return nombre;
        }

        private bool CreateDirectory()
        {
            try
            {
                //si no existe el directorio lo creamos
                if (!Directory.Exists(FullPath))
                    Directory.CreateDirectory(FullPath);

                return true;

            }
            catch (DirectoryNotFoundException ex)
            {
                addError(ex.Message + " Directorio invalido: " + FullPath);
                return false;
            }
        }


        public void addError(string error)
        {
            lErrores.Add(error);
        }

        public string getErrores()
        {
            string error = "";
            foreach (string err in lErrores)
            {
                error += err + BRINCO;
            }
            return error;
        }

        protected void clearErrores()
        {
            lErrores = new List<string>();
        }
        #endregion
    }
}
