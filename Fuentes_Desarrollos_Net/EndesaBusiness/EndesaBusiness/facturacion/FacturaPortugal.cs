using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class FacturaPortugal
    {
        InventarioFacturacionPortugal inventario;
        Spot spot;
        Clicks clicks;
        public FacturaPortugal(InventarioFacturacionPortugal _inventario, Spot _spot, Clicks _clicks)
        {
            inventario = _inventario;
            spot = _spot;
            clicks = _clicks;
        }

        public FacturaPortugal(DateTime periodoFacturacionDesde, DateTime periodoFacturacionHasta, string cpe)
        {
            
            int total_puntos_a_facturar = 0;
            List<string> lista_cups13_a_facturar = new List<string>();


            FacturaPortugalExcel excel = new FacturaPortugalExcel();

            try
            {
                

                foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in inventario.dic_inventario)
                {
                    if (p.Value.cpe == cpe && !p.Value.actualizado)
                    {
                        if (p.Value.ltp != "FACTURADO")
                        {
                            if (Convert.ToInt32(p.Value.ltp.Substring(0, 2)) == 10)
                            {
                                total_puntos_a_facturar++;
                                lista_cups13_a_facturar.Add(p.Value.cups13);
                            }
                        }

                    }

                }



                if (!spot.hayDatos)
                {
                    MessageBox.Show("El número de precios Spot cargados son " + String.Format("{0:#,##0}", spot.lista.Count())
                        + " cuando deberían ser " + String.Format("{0:#,##0}", spot.totalPeriodosCuartoHorarios) + "."
                        + System.Environment.NewLine
                        + "No se puede realizar la facturación.",
                     "Facturador Portugal",
                     MessageBoxButtons.OK,
                   MessageBoxIcon.Information);

                }
                else if (total_puntos_a_facturar > 0)
                {
                    DialogResult result3 = MessageBox.Show("¿Desea lanzar la facturación para " + total_puntos_a_facturar + " puntos",
                    "Facturador de Portugal",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (result3 == DialogResult.Yes)
                    {
                        //dirPS = new GO.facturacion.portugal.DireccionPS(inventario.dic_inventario.Select(z => z.Key).ToList());

                        // Cargamos medida
                        EndesaBusiness.medida.CCRD medida_cc =
                            new EndesaBusiness.medida.CCRD(lista_cups13_a_facturar, periodoFacturacionDesde, periodoFacturacionHasta);



                        foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in inventario.dic_inventario)
                        {

                            if (medida_cc.CurvaCompleta(p.Value.cups13))
                            {
                                excel.GeneraExcel(p.Value.carpeta_cliente, p.Value.ruta_plantilla, periodoFacturacionDesde, periodoFacturacionHasta, spot,
                                    medida_cc.GetCurvaVertical(p.Value.cups13), clicks.GetClicks(p.Value.cpe));
                                inventario.cpe = p.Value.cpe;
                                inventario.estado = "facturador generado";
                                inventario.actualizado = true;
                                inventario.Update();
                            }

                            else
                            {
                                inventario.cpe = p.Value.cpe;
                                inventario.estado = "CC Incompleta";
                                inventario.Update();
                            }

                        }


                    }

                    MessageBox.Show("Proceso finalizado correctamente.",
                      "Facturadores Portugal",
                      MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("No hay puntos pendientes de facturar.",
                      "Facturadores Portugal",
                      MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message,
                      "Facturadores Portugal",
                      MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
