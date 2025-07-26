using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Procesos
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);


            if (args.Length > 0)
            {
                //Application.Run(new Main(args[0]));
                Application.Run(new Main(args[0],true));
                //FrmTimer f = new FrmTimer(args[0]);
                //f.Show();
            }
                
            else
            {
                //Application.Run(new Main("Diario Medida"));
                //Application.Run(new Main(null));
                Application.Run(new FrmNuevaCola());
                //FrmNuevaCola f = new FrmNuevaCola();
                //f.Show();
            }
                
                




        }
    }
}
