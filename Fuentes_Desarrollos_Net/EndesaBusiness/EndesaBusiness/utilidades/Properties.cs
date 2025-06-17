using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.utilidades
{
    public class Properties
    {
        private Dictionary<string, string> dic;
        public Properties()
        {
            dic = CargaParametros();
        }

        private Dictionary<string, string> CargaParametros()
        {
            string line;
            int pos;
            string key;
            string value;

            Dictionary<string, string> d = new Dictionary<string, string>();

            FileInfo file = new FileInfo(System.Environment.CurrentDirectory + @"\bin\" + "properties");
            if (!file.Exists)
                MessageBox.Show("No se encuentra el archivo " + System.Environment.CurrentDirectory + @"\bin\" + "properties",
                    "MySQLDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                System.IO.StreamReader archivo = new System.IO.StreamReader(File.OpenRead(file.FullName));
                while ((line = archivo.ReadLine()) != null)
                {
                    if (line.Trim() != "")
                    {
                        pos = line.IndexOf('=');
                        key = line.Substring(0, pos);
                        value = line.Substring(pos + 1);
                        string a;
                        if (!d.TryGetValue(key, out a))
                            d.Add(key, value);
                    }

                }
                archivo.Close();
            }

            return d;
        }

        public string GetValue(string key)
        {
            return dic.First(z => z.Key == key).Value;
        }
    }
}
