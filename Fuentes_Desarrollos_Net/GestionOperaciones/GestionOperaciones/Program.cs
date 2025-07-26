using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();

            if (!(g.UsuarioValido()))
            {
                MessageBox.Show("El usuario " + Environment.UserDomainName.ToUpper() + "\\" + Environment.UserName.ToUpper() +
                " no está registrado para utilizar la aplicación.",
                "Usuario no permitido",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FormStart());
            }

            
        }
    }
}
