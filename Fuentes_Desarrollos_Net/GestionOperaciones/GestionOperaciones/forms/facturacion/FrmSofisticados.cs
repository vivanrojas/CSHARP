using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmSofisticados : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmSofisticados()
        {

            usage.Start("Facturación", "FrmSofisticados" ,"N/A");
            InitializeComponent();
        }

        private void BtnCurvas_Click(object sender, EventArgs e)
        {
            DateTime fd = Convert.ToDateTime(txt_fecha_consumo_desde.Value);
            DateTime fh = Convert.ToDateTime(txt_fecha_consumo_hasta.Value);
            string cliente = txtcliente.Text;
            Dictionary<string, EndesaEntity.medida.DiccionarioCurva> dic_curvas_resultados = new Dictionary<string, EndesaEntity.medida.DiccionarioCurva>();
            string cups22 = "";

            bool existeStarBeat = false;


            bool existeFTPDistribuidora = false;

            bool encontradaCurva = false;

            int activa = 0;


            int reactiva = 0;

            EndesaEntity.medida.Curva curvaTemp = new EndesaEntity.medida.Curva();


            if (txtcliente.Text == null || txtcliente.Text == "")
                errorProvider.SetError(txtcliente, "El filtro cliente debe estar informado.");
            else
            {

                DialogResult result_1 = MessageBox.Show("¿Actualizar datos?",
                    "Curvas horarias para el periodo "
                    + fd.ToString("dd/MM/yyyy") + " al "
                    + fh.ToString("dd/MM/yyyy"),
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result_1 == DialogResult.Yes)
                {
                    EndesaBusiness.facturacion.SofisticadosOrigenCurvas oc = new EndesaBusiness.facturacion.SofisticadosOrigenCurvas();
                    EndesaBusiness.facturacion.SofisticadosInventario inventario = new EndesaBusiness.facturacion.SofisticadosInventario(fd, fh, cliente);
                    EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes pm =
                        new EndesaBusiness.medida.PuntosMedidaPrincipalesVigentes(inventario.lista_pm);

                    EndesaBusiness.medida.Kee_Medida curvas_kee_StarBeat =
                        new EndesaBusiness.medida.Kee_Medida(pm.lista_cups22, "StarBeat", fd, fh);

                    EndesaBusiness.medida.Kee_Medida curvas_kee_FTPDistribuidora =
                        new EndesaBusiness.medida.Kee_Medida(pm.lista_cups22, "FTP Distribuidora", fd, fh);

                    EndesaBusiness.facturacion.SofisticadosProgramasConsumo pc = new EndesaBusiness.facturacion.SofisticadosProgramasConsumo(cliente, fd, fh);

                    // Unimos curvas
                    foreach (KeyValuePair<string, EndesaEntity.medida.PuntoSuministro> p in pm.dic)
                    {

                        for (DateTime f = fd; f < fh.AddDays(1); f = f.AddHours(1))
                        {
                            EndesaEntity.medida.Curva curvaFinal = new EndesaEntity.medida.Curva();
                            encontradaCurva = false;
                            //EndesaEntity.medida.Curva curvaFinal = new EndesaEntity.medida.Curva();
                            //existeStarBeat = false;
                            //existeFTPDistribuidora = false;
                            // Recorremos los puntos de medida
                            for (int x = 0; x < p.Value.cups22.Count; x++)
                            {
                                encontradaCurva = false;
                                cups22 = p.Value.cups22[x];
                                if (ExisteFuente(curvas_kee_StarBeat.dic, cups22, f))
                                {
                                    encontradaCurva = true;
                                    curvaTemp = GetHora(curvas_kee_StarBeat.dic, cups22, f);
                                    curvaFinal.a += curvaTemp.a;
                                    curvaFinal.r += curvaTemp.r;
                                    curvaFinal.origen = curvaFinal.origen + "S";

                                }
                                else if (ExisteFuente(curvas_kee_FTPDistribuidora.dic, cups22, f))
                                {
                                    encontradaCurva = true;
                                    curvaTemp = GetHora(curvas_kee_FTPDistribuidora.dic, cups22, f);
                                    curvaFinal.a += curvaTemp.a;
                                    curvaFinal.r += curvaTemp.r;
                                    curvaFinal.origen = curvaFinal.origen + "F";
                                }
                                else
                                    break;
                            }

                            EndesaEntity.medida.DiccionarioCurva o;
                            if (!dic_curvas_resultados.TryGetValue(p.Value.cups20, out o))
                            {
                                if (!encontradaCurva)
                                {
                                    o = new EndesaEntity.medida.DiccionarioCurva();
                                    o.dic.Add(f, GetHora(pc.dic, p.Value.cups20, f));
                                }

                                else
                                {
                                    o = new EndesaEntity.medida.DiccionarioCurva();
                                    o.dic.Add(f, curvaFinal);
                                }

                                dic_curvas_resultados.Add(p.Value.cups20, o);
                            }
                            else
                            {
                                if (!encontradaCurva)
                                    o.dic.Add(f, GetHora(pc.dic, p.Value.cups20, f));
                                else
                                    o.dic.Add(f, curvaFinal);
                            }
                        }

                    }

                    pc.GuardaCurva(dic_curvas_resultados);
                    MessageBox.Show("Proceso finalizado."
                        + System.Environment.NewLine
                        + System.Environment.NewLine
                        + "Puede consultar los resultados en la tabla fact.ag_ch.",
                    "Curvas horarias para clientes Sofisticados",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
        }

        private bool ExisteFuente(Dictionary<string, EndesaEntity.medida.DiccionarioCurva> ch, string cups22, DateTime f)
        {
            EndesaEntity.medida.DiccionarioCurva curva;
            if (ch.TryGetValue(cups22, out curva))
            {
                EndesaEntity.medida.Curva o;
                return curva.dic.TryGetValue(f, out o);
            }
            return false;
        }

        private EndesaEntity.medida.Curva GetHora(Dictionary<string, EndesaEntity.medida.DiccionarioCurva> ch, string cups22, DateTime f)
        {
            EndesaEntity.medida.DiccionarioCurva curva;
            if (ch.TryGetValue(cups22, out curva))
            {
                EndesaEntity.medida.Curva o;
                if (curva.dic.TryGetValue(f, out o))
                    return o;
            }
            return null;
        }



        

        private void FrmSofisticados_Load(object sender, EventArgs e)
        {
            DateTime mesAnterior = DateTime.Now;
            int anio = mesAnterior.Year;
            int mes = mesAnterior.Month;
            int dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fh = new DateTime(anio, mes, dias_del_mes);
            DateTime fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fecha_consumo_desde.Value = fd;
            txt_fecha_consumo_hasta.Value = fh;
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void FrmSofisticados_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmSofisticados" ,"N/A");
        }
    }
}
