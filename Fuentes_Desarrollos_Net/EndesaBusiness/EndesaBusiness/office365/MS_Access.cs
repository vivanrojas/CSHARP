using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Access = Microsoft.Office.Interop.Access;

namespace EndesaBusiness.office365
{
    public class MS_Access
    {
        public MS_Access()
        {

        }

        public EndesaEntity.cola.ProcesoCola EjecutaMacro(EndesaEntity.cola.ProcesoCola c)
        {
            Access.Application oAccess = new Access.Application();

            try
            {
                c.fecha_ultimo_lanzamiento = DateTime.Now;
                object oMissing = System.Reflection.Missing.Value;
                oAccess.Visible = true;
                oAccess.OpenCurrentDatabase(c.ruta + @"\" + c.bbdd, false, "");
                oAccess.DoCmd.RunMacro(c.nombre_proceso);

                // Run the macros.
                //RunMacroAccess(oAccess, new Object[] { macro });
                // RunMacroAccess(oAccess, new Object[]{"DoKbTestWithParameter",
                //                "Hello from C# Client."});

                // Quit Access and clean up.
                oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveNone);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                oAccess = null;
                c.fecha_ultima_ejec_ok = DateTime.Now;
                return c;
            }
            catch (Exception e)
            {
                oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveNone);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                c.mensaje_error = e.Message;
                return c;
            }


        }
    }
}
