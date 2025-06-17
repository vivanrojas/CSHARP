using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Access = Microsoft.Office.Interop.Access;

namespace Procesos.business
{
    class MSAccess
    {

        public MSAccess()
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
                c.mensaje_error = "";
                return c;
            }catch(Exception e)
            {
                oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveNone);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                c.mensaje_error = e.Message;                
                return c;
            }
            

        }

        private void RunMacroAccess(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
            System.Reflection.BindingFlags.Default |
            System.Reflection.BindingFlags.InvokeMethod,
            null, oApp, oRunArgs);
        }

        public EndesaEntity.cola.ProcesoCola ProcesoBatch(EndesaEntity.cola.ProcesoCola c)
        {
            try
            {
                c.fecha_ultimo_lanzamiento = DateTime.Now;

                if (c.parametro == null || c.parametro == "")
                    utilidades.Fichero.EjecutaComando(c.ruta + @"\\" + c.nombre_proceso);
                else                
                    utilidades.Fichero.EjecutaComando(c.ruta + @"\\" + c.nombre_proceso, Parametro(c.parametro));
                
                c.fecha_ultima_ejec_ok = DateTime.Now;
                return c;
            }
            catch(Exception e)
            {
                c.mensaje_error = e.Message;
                return c;
            }
           
        }

        private string Parametro(string parametro)
        {
            business.Fechas f = new Fechas();
            string salida = "";
            switch (parametro)
            {
                case "ULTIMO_DIA_HABIL_YYMMDD":
                    salida = f.UltimoDiaHabil().ToString("yyMMdd");
                    break;
                case "ULTIMO_DIA_HABIL_MMDD":
                    salida = f.UltimoDiaHabil().ToString("MMdd");
                    break;
                default:
                    salida = salida;
                    break;
            }

            return salida;
        }

    }
}
