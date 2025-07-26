using EndesaBusiness.contratacion.gestionATRGas;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion.gas
{
    public partial class FrmContratos : Form
    {
        bool firstOnly_aviso = true;
        List<string> c;
        EndesaBusiness.sigame.SIGAME sigame;
        EndesaBusiness.contratacion.gestionATRGas.ContratosGas contratos;
        EndesaBusiness.contratacion.gestionATRGas.Solicitudes sol;
        EndesaBusiness.sigame.Addendas addendas;

        List<EndesaEntity.Table_atrgas_contratos> lista;
        List<EndesaEntity.Informe_contratos_gas_vencimiento> lista_vencimiento_GO;
        List<EndesaEntity.Informe_contratos_gas_vencimiento> lista_vencimiento_SIGAME;
        List<EndesaEntity.Informe_contratos_gas_vencimiento> lista_vencimiento_total;
        EndesaBusiness.utilidades.Param param = new EndesaBusiness.utilidades.Param("atrgas_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmContratos()
        {
            usage.Start("Contratación", "FrmContratos", "Contratos Gas");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmContratos_Load(object sender, EventArgs e)
        {

            actualizaAddendasToolStripMenuItem.Enabled = System.Environment.UserName.ToUpper() == "ES01855";
            sigame = new EndesaBusiness.sigame.SIGAME();
            contratos = new EndesaBusiness.contratacion.gestionATRGas.ContratosGas();
            addendas = new EndesaBusiness.sigame.Addendas();
            actualizarTodasAddendas();
            LoadData();
            


            //dgv.AutoGenerateColumns = false;
            //dgv.DataSource = contratos.dic.Values.ToList();
            //lbl_total_contratos.Text = string.Format("Total Registros: {0:#,##0}", contratos.dic.Count());
            //dgvd.AutoGenerateColumns = false;
            //dgvd.DataSource = contratos.cgd.dic.FirstOrDefault(z => z.Key == contratos.dic.First().Key).Value.OrderByDescending(z => z.fecha_inicio);

            // var lista = contratos.dic.SelectMany(x => x.Value.cups20).ToList();            

        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewCell nif = (DataGridViewCell)
                    dgv.Rows[e.RowIndex].Cells[0];

                DataGridViewCell cups = (DataGridViewCell)
                   dgv.Rows[e.RowIndex].Cells[3];

                dgvd.AutoGenerateColumns = false;
                List<EndesaEntity.Table_atrgas_contratos_detalle> lista;
                if (contratos.cgd.dic.TryGetValue(nif.Value.ToString(), out lista))
                    dgvd.DataSource = lista.Where(z => z.cups20 == cups.Value.ToString()).ToList();
                else
                    dgvd.DataSource = null;
                // dgvd.DataSource = contratos.cgd.dic.FirstOrDefault(z => z.Key == cups20.Value.ToString()).Value.OrderByDescending(z => z.fecha_inicio).ToList();                

            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void txt_cups20_TextChanged(object sender, EventArgs e)
        {
            
        }

       

        private void Edit()
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;

            FrmContratos_Edit f = new FrmContratos_Edit();
            f.Text = "Editar Contrato";

            f.con = contratos;
            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                f.txt_cnifdnic.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_dapersoc.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.cmb_distribuidora.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_cups20.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_comentarios_descuadres.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.txt_comentarios_contratacion.Text = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                f.cmb_tramitacion.Text = row.Cells[c].Value.ToString(); c++;


            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void txt_dapersoc_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void LoadData()
        {
            sol = new EndesaBusiness.contratacion.gestionATRGas.Solicitudes(sigame);
            txt_dias_vencimiento.Text = param.GetValue("dias_caducidad", DateTime.Now, DateTime.Now);

            lista = new List<EndesaEntity.Table_atrgas_contratos>();
            
            contratos = new EndesaBusiness.contratacion.gestionATRGas.ContratosGas();                        

            foreach (KeyValuePair<string, List<EndesaEntity.Table_atrgas_contratos>> p in contratos.dic)
            {
                for(int i = 0; i < p.Value.Count(); i++)
                    lista.Add(p.Value[i]);
            }

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;            
            lbl_total_contratos.Text = string.Format("Total Registros: {0:#,##0}", lista.Count());


            dgvv.AutoGenerateColumns = false;
            lista_vencimiento_GO = contratos.cgd.Lista_Proximo_Vencimiento(Convert.ToInt32(param.GetValue("dias_caducidad", DateTime.Now, DateTime.Now)));
            lista_vencimiento_total = lista_vencimiento_GO;
            lista_vencimiento_SIGAME = contratos.cgd.Lista_Proximo_Vencimiento_SIGAME(Convert.ToInt32(param.GetValue("dias_caducidad", DateTime.Now, DateTime.Now)));
            
            foreach(EndesaEntity.Informe_contratos_gas_vencimiento p in lista_vencimiento_SIGAME)
            {
                if(!Existe_Contrato_en_GO(p))
                {
                    lista_vencimiento_total.Add(p);
                }
            }

            dgvv.DataSource = lista_vencimiento_total;
            lbl_total_vencimientos.Text = string.Format("Total Registros: {0:#,##0}", lista_vencimiento_total.Count());
            LoadDataDetail();

            if(firstOnly_aviso && sol.lista.Count > 0)
            {
                MessageBox.Show("Tiene solicitudes pendientes de validar.",
                "Solicitudes pendientes",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                firstOnly_aviso = false;
            }


            dgvSol.AutoGenerateColumns = false;
            dgvSol.DataSource = sol.lista;

        }

        private bool Existe_Contrato_en_GO(EndesaEntity.Informe_contratos_gas_vencimiento contrato)
        {
            bool existe = false;
            foreach (EndesaEntity.Informe_contratos_gas_vencimiento p in lista_vencimiento_GO)
            {
                if(p.nif == contrato.nif && p.cups20 == contrato.cups20 &&
                    p.fecha_inicio == contrato.fecha_inicio && p.fecha_fin == contrato.fecha_fin &&
                    p.tipo == contrato.tipo)
                {
                    existe = true;
                    break;
                }
            }

            return existe;
        }

        private void LoadDataDetail()
        {
            dgvd.AutoGenerateColumns = false;
            dgvd.DataSource = contratos.cgd.dic.FirstOrDefault(z => z.Key == contratos.dic.First().Key).Value;
        }


        private void Filters()
        {
            dgv.AutoGenerateColumns = false;
            

            // lista = contratos.dic.Values.ToList();
            
            if (txt_cups20.Text != "")
                lista = lista.Where(z => z.cups20 != null && z.cups20.Contains(txt_cups20.Text.ToUpper())).ToList();

            if (txt_dapersoc.Text != "")
                lista = lista.Where(z => z.dapersoc != null && z.dapersoc.Contains(txt_dapersoc.Text.ToUpper())).ToList();

            if (txt_cnifdnic.Text != "")
                lista = lista.Where(z => z.cnifdnic != null && z.cnifdnic.Contains(txt_cnifdnic.Text.ToUpper())).ToList();

            if (txt_distribuidora.Text != "")
                lista = lista.Where(z => z.distribuidora != null && z.distribuidora.Contains(txt_distribuidora.Text.ToUpper())).ToList();
            
            dgv.DataSource = lista;
            lbl_total_contratos.Text = string.Format("Total contratos: {0:#,##0}", lista.Count);

            if (lista.Count() > 0)
            {
                dgvd.AutoGenerateColumns = false;
                List<EndesaEntity.Table_atrgas_contratos_detalle> o;
                if (contratos.cgd.dic.TryGetValue(lista[0].cnifdnic, out o))
                    dgvd.DataSource = o.Where(z => z.cups20 == lista[0].cups20).ToList();
                else
                    dgvd.DataSource = null;
            }
            else
                dgvd.DataSource = null;
           

            
        }

        private void txt_cnifdnic_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txt_distribuidora_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            FrmContratos_Edit f = new FrmContratos_Edit();
            f.Text = "Nuevo contrato";
            f.con = contratos;            
            f.ShowDialog();
            LoadData();
            Filters();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            Edit();
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;
            
            row = dgv.Rows[fila];

            if (row.Cells[c].Value != null)
                contratos.cnifdnic = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                contratos.dapersoc = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                contratos.distribuidora = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                contratos.cups20 = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                contratos.comentarios_descuadres = row.Cells[c].Value.ToString(); c++;

            if (row.Cells[c].Value != null)
                contratos.comentarios_contratacion = row.Cells[c].Value.ToString(); c++;

            contratos.Del();

            LoadData();
            Filters();
        }

        private void btnEdit_d_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgvd.CurrentRow.Index;
            row = dgvd.Rows[fila];

            EditDetalle(row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString(), row.Cells[2].Value.ToString(),
                Convert.ToDateTime(row.Cells[3].Value), row.Cells[4].Value != null ? Convert.ToDateTime(row.Cells[4].Value) : DateTime.MinValue, 
                Convert.ToDouble(row.Cells[5].Value), row.Cells[8].Value != null ? row.Cells[8].Value.ToString(): null,
                row.Cells[9].Value != null ? row.Cells[9].Value.ToString() : null);
        }

        private void EditDetalle(string nif, string cups, string tipo, DateTime fd, DateTime fh, double qd, string tarifa, string comentario)
        {
          

            FrmContratosDetalle_Edit f = new FrmContratosDetalle_Edit();
            f.nuevo_registro = false;
            f.Text = "Editar Producto";
            f.con = contratos.cgd;

            f.con.last_data.nif = nif;
            f.con.last_data.cups20 = cups;
            f.con.last_data.fecha_inicio = fd;
            if (fh > DateTime.MinValue)
                f.con.last_data.fecha_fin = fh;
            f.con.last_data.qd = qd;
            f.con.last_data.tipo = tipo;
            if (tarifa != null)
                f.con.last_data.tarifa = tarifa;

            if (comentario != null)
                f.con.last_data.comentario = comentario;

            f.nif = nif;
            f.txt_cups20.Text = cups;
            f.txt_fecha_inicio.Text = fd.ToShortDateString();

            if(comentario != null)
                f.txt_comentario.Text = comentario;

            if (fh > DateTime.MinValue)
            {
                f.txt_fecha_fin.Enabled = true;
                f.txt_fecha_fin.Text = fh.ToShortDateString();
            }
            else
            {
                f.txt_fecha_fin.Text = "";
                
            }

            f.txt_qd.Text = qd.ToString("#.##");
            if (tarifa != null)
                f.cmb_tarifa.Text = tarifa.ToString();
            f.cmb_tipo.Text = tipo;
            
            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void btnAdd_d_Click(object sender, EventArgs e)
        {

            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;          
            row = dgv.Rows[fila];

            FrmContratosDetalle_Edit f = new FrmContratosDetalle_Edit();
            f.Text = "Nuevo producto " + row.Cells[1].Value.ToString() + " - " + row.Cells[3].Value.ToString();
            f.con = contratos.cgd;
            f.nif = row.Cells[0].Value.ToString();
            f.con.nuevo_registro = true;
            f.txt_cups20.Text = row.Cells[3].Value.ToString();
            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void dgvd_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex >= 0)
            {
                DataGridViewCell nif = (DataGridViewCell)
                    dgvd.Rows[e.RowIndex].Cells[0];

                DataGridViewCell cups = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[1];

                DataGridViewCell tipo = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[2];

                DataGridViewCell fd = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[3];

                DataGridViewCell fh = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[4];

                DataGridViewCell qd = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[5];

                DataGridViewCell tarifa = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[6];

                DataGridViewCell comentario = (DataGridViewCell)
                   dgvd.Rows[e.RowIndex].Cells[7];


                EditDetalle(nif.Value.ToString(),cups.Value.ToString(),tipo.Value.ToString(),
                    Convert.ToDateTime(fd.Value), fh.Value != null ? Convert.ToDateTime(fh.Value) : DateTime.MinValue,
                    Convert.ToDouble(qd.Value), tarifa.Value != null ? tarifa.Value.ToString() : null,
                    comentario.Value.ToString());
            }

            
        }

        private void contratosSolapadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GestionOperaciones.forms.contratacion.gas.FrmFechasInforme f = new FrmFechasInforme();
            f.ShowDialog();

            if (!f.cancelado)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    ExportExcel(save.FileName, f.txt_fd.Value, f.txt_fh.Value);
                    Cursor.Current = Cursors.Default;

                    DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result2 == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                }
            }
        }


        private void ExportExcel(string rutaFichero, DateTime fd, DateTime fh)
        {
            int fila = 0;
            int columna = 0;

            InicializaColumnas();
            bool firstOnly = true;

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Solapados");

            var headerCells = workSheet.Cells[1, 1, 1, 10];
            var headerFont = headerCells.Style.Font;
                      

            fila++;
            columna = 0;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "ESTADO"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CUPS"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CIF"; columna++;            
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Qd"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Hora Inicio"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Qi"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "Tarifa"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "TIPO"; 
            PintaRecuadro(excelPackage, fila, columna);
                       

            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe inf = 
                new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe(fd, fh, null);

            foreach (KeyValuePair<string, List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle>> p in inf.dic)
            {
                if (p.Value.Count > 1)
                {

                    if (MostrarSolapados(p.Value))
                    {



                        firstOnly = true;
                        for (int i = 0; i < p.Value.Count; i++)
                        {

                            fila++;
                            columna = 0;
                            if (firstOnly)
                            {
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].estado; columna++;
                                workSheet.Cells[c[columna] + fila].Value = p.Key; columna++;
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].customer_name; columna++;
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].vatnum; columna++;
                                firstOnly = false;
                            }
                            else
                                columna = 4;

                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].fecha_inicio;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;

                            if (p.Value[i].fecha_fin > DateTime.MinValue)
                            {
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].fecha_fin;
                                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                            columna++;


                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].qd;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;

                            if (p.Value[i].hora_inicio > DateTime.MinValue)
                            {
                                workSheet.Cells[c[columna] + fila].Value = p.Value[i].hora_inicio;
                                workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern;
                            }
                            columna++;

                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].qi;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;

                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].tarifa; columna++;
                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].tipo.ToUpper(); columna++;
                            //workSheet.Cells[c[columna] + fila].Value = p.Value[i].last_update_date;
                            //workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;
                        }

                        fila++;
                    }
                }
            }

            var allCells = workSheet.Cells[1, 1, fila, 20];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:K1"].AutoFilter = true;
            workSheet.View.FreezePanes(2, 1);

            headerFont.Bold = true;




            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, fila, 9];
            allCells.AutoFitColumns();

            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }


        private bool MostrarSolapados(List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle> lista)
        {

            Dictionary<string, string> dic_tipos_productos_adicionales = new Dictionary<string, string>();
            Dictionary<string, string> dic_tipos_productos = new Dictionary<string, string>();
            string o;

            // Nos indica si hay que mostrar los solapados
            // Para ello es suficiente mostrar aquellos CUPS que tengan en un periodo de tiempo
            // determinado más de un tipo de producto contratado.
            foreach(EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle p in lista)
            {
                if(p.linea_solicitud > 0 && p.id_solicitud > 0)
                {

                    if(!dic_tipos_productos_adicionales.TryGetValue(p.tipo, out o))
                        dic_tipos_productos_adicionales.Add(p.tipo,p.tipo);
                }
                else
                {
                    if (!dic_tipos_productos.TryGetValue(p.tipo, out o))
                        dic_tipos_productos.Add(p.tipo, p.tipo);
                }
            }

            if (dic_tipos_productos_adicionales.Count > 0 && dic_tipos_productos.Count > 0)
                return true;
            else
                return false;

        }

        private void ExportExcel_old(string rutaFichero, DateTime fd, DateTime fh)
        {
            int fila = 0;
            int columna = 0;

            InicializaColumnas();

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Solapados");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            fila = 1;            
            workSheet.Cells[c[4] + fila].Value = "PRODUCTO BASE"; 
            workSheet.Cells[c[4] + fila + ":" + c[8] + fila].Merge = true;
            workSheet.Cells[c[4] + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[c[9] + fila].Value = "PRODUCTO SOLAPADO"; 
            workSheet.Cells[c[9] + fila + ":" + c[13] + fila].Merge = true;
            workSheet.Cells[c[9] + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            headerCells = workSheet.Cells["A2:O2"];
            headerFont = headerCells.Style.Font;

            fila++;
            columna = 0;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "ESTADO"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CUPS"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "CIF"; columna++;
            PintaRecuadro(excelPackage, fila, columna);

            for (int i = 0; i <= 1; i++)
            {
                PintaRecuadro(excelPackage, fila, columna);
                workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
                PintaRecuadro(excelPackage, fila, columna);
                workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
                PintaRecuadro(excelPackage, fila, columna);
                workSheet.Cells[c[columna] + fila].Value = "Qd"; columna++;                
                PintaRecuadro(excelPackage, fila, columna);
                workSheet.Cells[c[columna] + fila].Value = "Tarifa"; columna++;
                PintaRecuadro(excelPackage, fila, columna);
                workSheet.Cells[c[columna] + fila].Value = "TIPO"; columna++;                
            }

            PintaRecuadro(excelPackage, fila, columna);
            workSheet.Cells[c[columna] + fila].Value = "ÚLTIMA MODIFICACIÓN"; columna++;

            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe inf = new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle_Informe(fd, fh, txt_cups20.Text);

            foreach (KeyValuePair<string, List<EndesaBusiness.contratacion.gestionATRGas.ContratoGasDetalle>> p in inf.dic)
            {
                if (p.Value.Count > 1)
                {
                    for(int i = 1; i < p.Value.Count; i++)
                    {
                        // PRODUCTO BASE
                        fila++;
                        columna = 0;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].estado; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Key; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].customer_name; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].vatnum; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].fecha_inicio;
                        workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;

                        if (p.Value[0].fecha_fin > DateTime.MinValue)
                        {
                            workSheet.Cells[c[columna] + fila].Value = p.Value[0].fecha_fin;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }

                        columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].qd;
                        workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].tarifa; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[0].tipo.ToUpper(); columna++;

                        // PRODUCTO SOLAPADO                               
                        workSheet.Cells[c[columna] + fila].Value = p.Value[i].fecha_inicio;
                        workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;
                        if (p.Value[i].fecha_fin > DateTime.MinValue)
                        {
                            workSheet.Cells[c[columna] + fila].Value = p.Value[i].fecha_fin;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }

                        columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[i].qd;
                        workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[i].tarifa; columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[i].tipo.ToUpper(); columna++;
                        workSheet.Cells[c[columna] + fila].Value = p.Value[i].last_update_date;
                        workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;
                    }
                }
            }

            var allCells = workSheet.Cells[1, 1, fila, 20];
            allCells.AutoFitColumns();
            workSheet.Cells["A2:O2"].AutoFilter = true;
            workSheet.View.FreezePanes(3, 1);

            headerFont.Bold = true;
                       

            headerFont.Bold = true;
            allCells = workSheet.Cells[1, 1, fila, 10];
            allCells.AutoFitColumns();
            
            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }

        private void PintaRecuadro(ExcelPackage excelPackage, int fila, int columna)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[c[columna] + fila].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[c[columna] + fila].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[c[columna] + fila].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[c[columna] + fila].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }



        private void InicializaColumnas()
        {

            c = new List<string>();
            for (char i = 'A'; i < 'Z'; i++)
                c.Add(i.ToString());
        }

        private void btnDel_d_Click(object sender, EventArgs e)
        {
            
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;
            row = dgv.Rows[fila];
            contratos.cgd.nif = row.Cells[0].Value.ToString();
            contratos.cgd.cups20 = row.Cells[3].Value.ToString();


            row = new DataGridViewRow();
            fila = dgvd.CurrentRow.Index;
            c = 0;
            row = dgvd.Rows[fila];

           

            if (row.Cells[3].Value != null)
                contratos.cgd.fecha_inicio = Convert.ToDateTime(row.Cells[3].Value); c++;

            if (row.Cells[4].Value != null)
                contratos.cgd.fecha_fin = Convert.ToDateTime(row.Cells[4].Value); c++;

            if (row.Cells[2].Value != null)
                contratos.cgd.tipo = row.Cells[2].Value.ToString(); c++;

            if (row.Cells[5].Value != null)
                contratos.cgd.qd = Convert.ToDouble(row.Cells[5].Value);

            contratos.cgd.Del();

            LoadData();
            Filters();
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            LoadData();            
            Filters();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            DialogResult result_1 = MessageBox.Show("¿Desea exportar los contratos a Excel?", "Expotación de contratos a Excel",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)
            {

                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    ExportExcelContratos(save.FileName);
                    Cursor.Current = Cursors.Default;

                    DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result2 == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                }
            }
        }

        private void ExportExcelContratos(string rutaFichero)
        {
            int fila = 0;
            int columna = 0;

            InicializaColumnas();
            EndesaEntity.Table_atrgas_contratos_detalle dd = new EndesaEntity.Table_atrgas_contratos_detalle();

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Contratos");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            fila = 1;
            workSheet.Cells[c[columna] + fila].Value = "NIF";  columna++;                      
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "DISTRIBUIDORA"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "CUPS"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "COMENTARIOS DESCUADRES"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "COMENTARIOS CONTRATACIÓN"; columna++;            
            workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Qd (kWh/día)"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "TARIFA"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "TIPO"; columna++;

            for (int i = 0; i < lista.Count(); i++)
            {
                List<EndesaEntity.Table_atrgas_contratos_detalle> o;
                if (contratos.cgd.dic.TryGetValue(lista[i].cnifdnic, out o))
                {
                    for (int j = 0; j < o.Count(); j++)
                    {
                        fila++;
                        columna = 0;
                        workSheet.Cells[c[columna] + fila].Value = lista[i].cnifdnic; columna++;
                        workSheet.Cells[c[columna] + fila].Value = lista[i].dapersoc; columna++;
                        workSheet.Cells[c[columna] + fila].Value = lista[i].distribuidora; columna++;
                        workSheet.Cells[c[columna] + fila].Value = lista[i].cups20; columna++;
                        workSheet.Cells[c[columna] + fila].Value = lista[i].comentarios_descuadres; columna++;
                        workSheet.Cells[c[columna] + fila].Value = lista[i].comentarios_contratacion; columna++;
                        workSheet.Cells[c[columna] + fila].Value = o[j].fecha_inicio; workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;
                        if(o[j].fecha_fin > DateTime.MinValue)
                        {
                            workSheet.Cells[c[columna] + fila].Value = o[j].fecha_fin;
                            workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; 
                        }

                        columna++;
                        workSheet.Cells[c[columna] + fila].Value = o[j].qd; workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;
                        workSheet.Cells[c[columna] + fila].Value = o[j].tarifa; columna++;
                        workSheet.Cells[c[columna] + fila].Value = o[j].tipo; columna++;
                    }
                }
                else
                {
                    fila++;
                    columna = 0;
                    workSheet.Cells[c[columna] + fila].Value = lista[i].cnifdnic; columna++;
                    workSheet.Cells[c[columna] + fila].Value = lista[i].dapersoc; columna++;
                    workSheet.Cells[c[columna] + fila].Value = lista[i].distribuidora; columna++;
                    workSheet.Cells[c[columna] + fila].Value = lista[i].cups20; columna++;
                    workSheet.Cells[c[columna] + fila].Value = lista[i].comentarios_descuadres; columna++;
                    workSheet.Cells[c[columna] + fila].Value = lista[i].comentarios_contratacion; columna++;
                }
                
            }
            
            var allCells = workSheet.Cells[1, 1, fila, 20];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:K1"].AutoFilter = true;

            headerFont.Bold = true;

                        
            allCells = workSheet.Cells[1, 1, fila, 10];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }

        private void ExportExcelContratosVencimiento(string rutaFichero)
        {
            int fila = 0;
            int columna = 0;

            InicializaColumnas();
            EndesaEntity.Table_atrgas_contratos_detalle dd = new EndesaEntity.Table_atrgas_contratos_detalle();

            FileInfo file = new FileInfo(rutaFichero);

            if (file.Exists)
                file.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(file);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Contratos");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            fila = 1;
            workSheet.Cells[c[columna] + fila].Value = "NIF"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "CLIENTE"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "DISTRIBUIDORA"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "CUPS"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "TIPO"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Fecha Inicio"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Fecha Fin"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Qd (kWh/día)"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "TARIFA"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Gestor"; columna++;
            workSheet.Cells[c[columna] + fila].Value = "Responsable Territorial"; columna++;


            for (int i = 0; i < lista_vencimiento_total.Count(); i++)
            {
               
                fila++;
                columna = 0;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].nif; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].cliente; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].distribuidora; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].cups20; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].tipo; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].fecha_inicio; workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; columna++;
                if (lista_vencimiento_total[i].fecha_fin > DateTime.MinValue)
                {
                    workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].fecha_fin;
                    workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }

                columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].qd; workSheet.Cells[c[columna] + fila].Style.Numberformat.Format = "#,##0"; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].tarifa; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].gestor; columna++;
                workSheet.Cells[c[columna] + fila].Value = lista_vencimiento_total[i].responsable_territorial; columna++;

            }

            var allCells = workSheet.Cells[1, 1, fila, 20];
            allCells.AutoFitColumns();
            workSheet.Cells["A1:K1"].AutoFilter = true;

            headerFont.Bold = true;


            allCells = workSheet.Cells[1, 1, fila, 10];
            workSheet.View.FreezePanes(2, 1);
            allCells.AutoFitColumns();
            excelPackage.Save();

            MessageBox.Show("Informe terminado.",
             "Exportación a Excel",
             MessageBoxButtons.OK,
             MessageBoxIcon.Information);
        }

       

        private void cmdSearchVencimiento_Click(object sender, EventArgs e)
        {
            dgvv.AutoGenerateColumns = false;
            lista_vencimiento_GO = contratos.cgd.Lista_Proximo_Vencimiento(Convert.ToInt32(txt_dias_vencimiento.Text == "" ? "999" : txt_dias_vencimiento.Text));
            lista_vencimiento_total = lista_vencimiento_GO;
            lista_vencimiento_SIGAME = contratos.cgd.Lista_Proximo_Vencimiento_SIGAME(Convert.ToInt32(txt_dias_vencimiento.Text == "" ? "999" : txt_dias_vencimiento.Text));

            foreach (EndesaEntity.Informe_contratos_gas_vencimiento p in lista_vencimiento_SIGAME)
            {
                if (!Existe_Contrato_en_GO(p))
                {
                    lista_vencimiento_total.Add(p);
                }
            }


            dgvv.DataSource = lista_vencimiento_total;
            lbl_total_vencimientos.Text = string.Format("Total Registros: {0:#,##0}", lista_vencimiento_total.Count());
            LoadDataDetail();
        }

        private void txt_dias_vencimiento_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))            
                e.Handled = true;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result_1 = MessageBox.Show("¿Desea exportar los contratos a Excel?", "Expotación de contratos de próximo vencimiento a Excel",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)
            {

                SaveFileDialog save = new SaveFileDialog();
                save.Title = "Ubicación del informe";
                save.AddExtension = true;
                save.DefaultExt = "xlsx";
                save.Filter = "Ficheros xslx (*.xlsx)|*.*";
                DialogResult result = save.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    ExportExcelContratosVencimiento(save.FileName);
                    Cursor.Current = Cursors.Default;

                    DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result2 == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                }
            }
        }

        private void ValidarSelecciónToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle contratoGasDetalle =
                    new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle();
                DialogResult result_1 = MessageBox.Show("¿Desea validar las solicitudes seleccionadas?",
                   "Solicitudes pendientes de aceptación", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result_1 == DialogResult.Yes)
                    foreach (DataGridViewRow row in dgvSol.SelectedRows)
                    {

                        DataGridViewCell id = (DataGridViewCell)
                         row.Cells[1];
                        DataGridViewCell linea = (DataGridViewCell)
                         row.Cells[2];
                        DataGridViewCell nif = (DataGridViewCell)
                         row.Cells[3];
                        DataGridViewCell cups20 = (DataGridViewCell)
                         row.Cells[8];
                        DataGridViewCell fecha_inicio = (DataGridViewCell)
                         row.Cells[11];
                        DataGridViewCell fecha_fin = (DataGridViewCell)
                         row.Cells[12];
                        DataGridViewCell tipo = (DataGridViewCell)
                         row.Cells[9];
                        DataGridViewCell qd = (DataGridViewCell)
                         row.Cells[10];
                        DataGridViewCell tarifa = (DataGridViewCell)
                         row.Cells[13];
                        DataGridViewCell qi = (DataGridViewCell)
                        row.Cells[14];
                        DataGridViewCell hora_inicio = (DataGridViewCell)
                         row.Cells[15];


                        contratoGasDetalle.id_solicitud = Convert.ToInt64(id.Value);
                        contratoGasDetalle.linea_solicitud = Convert.ToInt32(linea.Value);
                        contratoGasDetalle.comentario = null;
                        contratoGasDetalle.tarifa = null;
                        contratoGasDetalle.nif = nif.Value.ToString();
                        contratoGasDetalle.cups20 = cups20.Value.ToString();
                        contratoGasDetalle.fecha_inicio = Convert.ToDateTime(fecha_inicio.Value.ToString());
                        contratoGasDetalle.fecha_fin = Convert.ToDateTime(fecha_fin.Value.ToString());
                        contratoGasDetalle.tipo = tipo.Value.ToString();
                        contratoGasDetalle.qd = Convert.ToDouble(qd.Value.ToString());
                        contratoGasDetalle.created_by = System.Environment.UserName;
                        contratoGasDetalle.creation_date = DateTime.Now;
                        contratoGasDetalle.tarifa = tarifa.Value.ToString();
                        contratoGasDetalle.qi = Convert.ToDouble(qi.Value.ToString());
                        contratoGasDetalle.hora_inicio = Convert.ToDateTime(hora_inicio.Value.ToString());

                        // Si no existe el contrato cabecera del cliente lo creamos
                        if (!contratos.ExisteContrato(contratoGasDetalle.nif, contratoGasDetalle.cups20))
                        {
                            EndesaEntity.sigame.ContratoGas o;
                            if (sigame.dic.TryGetValue(contratoGasDetalle.cups20, out o))
                            {
                                contratos.cnifdnic = contratoGasDetalle.nif;
                                contratos.dapersoc = o.nombre_cliente;
                                contratos.distribuidora = o.distribuidora;
                                contratos.cups20 = contratoGasDetalle.cups20;
                                contratos.comentarios_descuadres = null;
                                contratos.comentarios_contratacion = null;
                                contratos.created_by = "GO";
                                contratos.creation_date = DateTime.Now;
                                contratos.Save();
                            }

                        }

                        if (contratoGasDetalle.ExisteRegistro(contratoGasDetalle.nif, contratoGasDetalle.cups20, contratoGasDetalle.fecha_inicio,
                            contratoGasDetalle.tipo, contratoGasDetalle.qd, ""))
                        {
                            FrmContratosDetalle_Edit f = new FrmContratosDetalle_Edit();
                            f.nuevo_registro = true;
                            f.Text = "Nuevo producto " + contratoGasDetalle.cups20 + " - " + contratoGasDetalle.tipo;
                            f.nif = contratoGasDetalle.nif;
                            f.txt_cups20.Text = contratoGasDetalle.cups20;
                            f.txt_fecha_inicio.Text = contratoGasDetalle.fecha_inicio.ToShortDateString();
                            f.cmb_tarifa.Text = sigame.Tarifa(contratoGasDetalle.cups20);
                            


                            if (contratoGasDetalle.fecha_fin > DateTime.MinValue)
                            {
                                f.txt_fecha_fin.Enabled = true;
                                f.txt_fecha_fin.Text = contratoGasDetalle.fecha_fin.ToShortDateString();
                            }
                            else
                            {
                                f.txt_fecha_fin.Text = "";
                            }

                            f.txt_qd.Text = contratoGasDetalle.qd.ToString("#.##");
                            if (tarifa != null)
                                f.cmb_tarifa.Text = tarifa.ToString();
                            f.cmb_tipo.Text = contratoGasDetalle.tipo;
                            f.con = contratoGasDetalle;
                            f.nif = contratoGasDetalle.nif;
                            f.txt_cups20.Text = contratoGasDetalle.cups20;

                            f.ShowDialog();
                        }
                        else
                        {
                            contratoGasDetalle.New();
                            sol.ValidaSolicitud(Convert.ToInt32(id.Value.ToString()), Convert.ToInt32(linea.Value.ToString()));
                        }


                        // Actualizamos los datos del contrato con las Addendas
                        List<EndesaEntity.contratacion.gas.Addenda> lista_addendas =
                            addendas.GetAddendas(contratoGasDetalle.cups20, contratoGasDetalle.fecha_inicio, contratoGasDetalle.fecha_fin);

                        for (int x = 0; x < lista_addendas.Count; x++)
                        {
                            if (lista_addendas[x].duracion_peaje.ToUpper() == "DIARIO")
                            {
                                for (DateTime j = lista_addendas[x].fecha_desde;
                                    j <= lista_addendas[x].fecha_hasta; j = j.AddDays(1))
                                {
                                    contratoGasDetalle.qd = lista_addendas[x].qd;
                                    contratoGasDetalle.tarifa = lista_addendas[x].tarifa_peaje;
                                    contratoGasDetalle.cups20 = lista_addendas[x].cups;
                                    contratoGasDetalle.tipo = "DIARIO";
                                    contratoGasDetalle.fecha_inicio = j;
                                    contratoGasDetalle.fecha_fin = j;
                                    contratoGasDetalle.comentario = "ADDENDA SIGAME";
                                    contratoGasDetalle.Replace();
                                }

                            }
                            else
                            {
                                contratoGasDetalle.qd = lista_addendas[x].qd;
                                contratoGasDetalle.tarifa = lista_addendas[x].tarifa_peaje;
                                contratoGasDetalle.cups20 = lista_addendas[x].cups;
                                contratoGasDetalle.tipo = lista_addendas[x].duracion_peaje;
                                contratoGasDetalle.fecha_inicio = lista_addendas[x].fecha_desde;                               
                                contratoGasDetalle.fecha_fin = lista_addendas[x].fecha_hasta;
                                contratoGasDetalle.comentario = "ADDENDA SIGAME";
                                contratoGasDetalle.Replace();
                            }

                        }


                        

                    }
                LoadData();
            }catch(Exception ee)
            {
                MessageBox.Show("Error a la hora de tramitar la solicitud",
                    "Se ha producido un error a la hora de tramitar la solicitud",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAdd_d_Click_1(object sender, EventArgs e)
        {

            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            row = dgv.Rows[fila];

            FrmContratosDetalle_Edit f = new FrmContratosDetalle_Edit();
            f.nuevo_registro = true;
            f.Text = "Nuevo producto " + row.Cells[1].Value.ToString() + " - " + row.Cells[3].Value.ToString();
            f.con = contratos.cgd;
            f.nif = row.Cells[0].Value.ToString();
            f.con.nuevo_registro = true;
            f.txt_cups20.Text = row.Cells[3].Value.ToString();
            f.ShowDialog();

            LoadData();
            Filters();
        }

        private void BtnDel_d_Click_1(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            int c = 0;
            row = dgv.Rows[fila];
            contratos.cgd.nif = row.Cells[0].Value.ToString();
            contratos.cgd.cups20 = row.Cells[3].Value.ToString();


            row = new DataGridViewRow();
            fila = dgvd.CurrentRow.Index;
            c = 0;
            row = dgvd.Rows[fila];



            if (row.Cells[3].Value != null)
                contratos.cgd.fecha_inicio = Convert.ToDateTime(row.Cells[3].Value); c++;

            if (row.Cells[4].Value != null)
                contratos.cgd.fecha_fin = Convert.ToDateTime(row.Cells[4].Value); c++;

            if (row.Cells[2].Value != null)
                contratos.cgd.tipo = row.Cells[2].Value.ToString(); c++;

            if (row.Cells[5].Value != null)
                contratos.cgd.qd = Convert.ToDouble(row.Cells[5].Value);

            if (row.Cells[6].Value != null)
                contratos.cgd.qi = Convert.ToDouble(row.Cells[6].Value);

            if (row.Cells[9].Value != null)
                contratos.cgd.comentario = row.Cells[9].Value.ToString();

            contratos.cgd.Del();

            LoadData();
            Filters();
        }

        private void BtnEdit_d_Click_1(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgvd.CurrentRow.Index;
            row = dgvd.Rows[fila];

            EditDetalle(row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString(), row.Cells[2].Value.ToString(),
                Convert.ToDateTime(row.Cells[3].Value), row.Cells[4].Value != null ? Convert.ToDateTime(row.Cells[4].Value) : DateTime.MinValue,
                Convert.ToDouble(row.Cells[5].Value), row.Cells[8].Value != null ? row.Cells[8].Value.ToString() : null,
                row.Cells[9].Value != null ? row.Cells[9].Value.ToString() : null);
        }

        private void cancelarSelecciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle contratoGasDetalle = new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle();
            DialogResult result_1 = MessageBox.Show("¿Desea cancelar las solicitudes seleccionadas?",
               "Solicitudes pendientes de aceptación", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)
                foreach (DataGridViewRow row in dgvSol.SelectedRows)
                {

                    DataGridViewCell id = (DataGridViewCell)
                     row.Cells[1];
                    DataGridViewCell linea = (DataGridViewCell)
                     row.Cells[2];
                    DataGridViewCell nif = (DataGridViewCell)
                     row.Cells[3];
                    DataGridViewCell cups20 = (DataGridViewCell)
                     row.Cells[8];
                    DataGridViewCell fecha_inicio = (DataGridViewCell)
                     row.Cells[11];
                    DataGridViewCell fecha_fin = (DataGridViewCell)
                     row.Cells[12];
                    DataGridViewCell tipo = (DataGridViewCell)
                     row.Cells[9];
                    DataGridViewCell qd = (DataGridViewCell)
                     row.Cells[10];

                    contratoGasDetalle.nif = nif.Value.ToString();
                    contratoGasDetalle.cups20 = cups20.Value.ToString();
                    contratoGasDetalle.fecha_inicio = Convert.ToDateTime(fecha_inicio.Value.ToString());
                    contratoGasDetalle.fecha_fin = Convert.ToDateTime(fecha_fin.Value.ToString());
                    contratoGasDetalle.tipo = tipo.Value.ToString();
                    contratoGasDetalle.qd = Convert.ToDouble(qd.Value.ToString());
                    contratoGasDetalle.created_by = System.Environment.UserName;
                    contratoGasDetalle.creation_date = DateTime.Now;

                    sol.CancelaSolicitud(Convert.ToInt32(id.Value), Convert.ToInt32(linea.Value));

                    

                }
            LoadData();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void solicitudManualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            forms.contratacion.gas.FrmSolicitudManual f = new FrmSolicitudManual();
            Cursor.Current = Cursors.Default;
            f.Show();
        }

        private void actualizarAddendasToolStripMenuItem_Click(object sender, EventArgs e)
        {

            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle contratoGasDetalle =
                   new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle();

            DialogResult result_1 = MessageBox.Show("¿Desea actualizar las addendas (SIGAME) para los contratos seleccionados?",
               "Actualizar Addendas", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)
                foreach (DataGridViewRow row in dgv.SelectedRows)
                {

                    
                    DataGridViewCell cups20 = (DataGridViewCell)
                     row.Cells[3];

                    DataGridViewCell nif = (DataGridViewCell)
                     row.Cells[0];

                    

                    List<EndesaEntity.contratacion.gas.Addenda> lista_addendas = addendas.GetAddendas(cups20.Value.ToString());                    

                    for (int x = 0; x < lista_addendas.Count; x++)
                    {
                        if (lista_addendas[x].duracion_peaje.ToUpper() == "DIARIO")
                        {
                            for (DateTime j = lista_addendas[x].fecha_desde;
                                j <= lista_addendas[x].fecha_hasta; j = j.AddDays(1))
                            {
                                contratoGasDetalle.nif = nif.Value.ToString();
                                contratoGasDetalle.qd = lista_addendas[x].qd;
                                contratoGasDetalle.tarifa = lista_addendas[x].tarifa_peaje;
                                contratoGasDetalle.cups20 = lista_addendas[x].cups;
                                contratoGasDetalle.tipo = "DIARIO";
                                contratoGasDetalle.fecha_inicio = j;
                                contratoGasDetalle.fecha_fin = j;
                                contratoGasDetalle.comentario = "ADDENDA SIGAME";
                                contratoGasDetalle.Replace();
                            }

                        }
                        else
                        {

                            contratoGasDetalle.nif = nif.Value.ToString();
                            contratoGasDetalle.qd = lista_addendas[x].qd;
                            contratoGasDetalle.tarifa = lista_addendas[x].tarifa_peaje;
                            contratoGasDetalle.cups20 = lista_addendas[x].cups;
                            contratoGasDetalle.tipo = lista_addendas[x].duracion_peaje;
                            contratoGasDetalle.fecha_inicio = lista_addendas[x].fecha_desde;
                            contratoGasDetalle.fecha_fin = lista_addendas[x].fecha_hasta;
                            contratoGasDetalle.comentario = "ADDENDA SIGAME";

                            if(contratoGasDetalle.Existe_ContratoDetalle_ConFinMundo(contratoGasDetalle.nif,
                                contratoGasDetalle.cups20,
                                contratoGasDetalle.tipo, contratoGasDetalle.fecha_inicio))

                                contratoGasDetalle.Del(contratoGasDetalle.nif,
                                contratoGasDetalle.cups20,
                                contratoGasDetalle.tipo, 
                                contratoGasDetalle.fecha_inicio,
                                new DateTime(4999,12,31).Date);
                            

                                contratoGasDetalle.Replace();
                        }

                    }







                }
            LoadData();
        }

        private void actualizaAddendasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            actualizarTodasAddendas();

            LoadData();
        }

        private void contextMenuStrip2_Opening(object sender, CancelEventArgs e)
        {

        }

        private void FrmContratos_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void FrmContratos_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmContratos", "Contratos Gas");
        }

        private void actualizarTodasAddendas()
        {
            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle contratoGasDetalle =
                   new EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle();

            DialogResult result_1 = MessageBox.Show("¿Desea actualizar las addendas (SIGAME) para todos los contratos?",
               "Actualizar Addendas", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (KeyValuePair<string, List<EndesaEntity.Table_atrgas_contratos>> p in contratos.dic)
                {
                    Console.WriteLine("Actualizadon Addendas Cliente " + p.Key);
                    for (int i = 0; i < p.Value.Count; i++)
                    {
                        List<EndesaEntity.contratacion.gas.Addenda> lista_addendas = addendas.GetAddendas(p.Value[i].cups20);

                        if (lista_addendas != null)
                            for (int x = 0; x < lista_addendas.Count; x++)
                            {
                                Console.WriteLine("Actualizando Addenda " + x + " para " + p.Value[i].cnifdnic + " - " + p.Value[i].dapersoc);
                                if (lista_addendas[x].duracion_peaje.ToUpper() == "DIARIO")
                                {
                                    for (DateTime j = lista_addendas[x].fecha_desde;
                                        j <= lista_addendas[x].fecha_hasta; j = j.AddDays(1))
                                    {

                                        contratoGasDetalle.nif = p.Value[i].cnifdnic;
                                        contratoGasDetalle.qd = lista_addendas[x].qd;
                                        contratoGasDetalle.tarifa = lista_addendas[x].tarifa_peaje;
                                        contratoGasDetalle.cups20 = lista_addendas[x].cups;
                                        contratoGasDetalle.tipo = "DIARIO";
                                        contratoGasDetalle.fecha_inicio = j;
                                        contratoGasDetalle.fecha_fin = j;
                                        contratoGasDetalle.comentario = "ADDENDA SIGAME";
                                        contratoGasDetalle.Replace();
                                    }

                                }
                                else
                                {

                                    contratoGasDetalle.nif = p.Value[i].cnifdnic;
                                    contratoGasDetalle.qd = lista_addendas[x].qd;
                                    contratoGasDetalle.tarifa = lista_addendas[x].tarifa_peaje;
                                    contratoGasDetalle.cups20 = lista_addendas[x].cups;
                                    contratoGasDetalle.tipo = lista_addendas[x].duracion_peaje;
                                    contratoGasDetalle.fecha_inicio = lista_addendas[x].fecha_desde;
                                    contratoGasDetalle.fecha_fin = lista_addendas[x].fecha_hasta;
                                    contratoGasDetalle.comentario = "ADDENDA SIGAME";

                                    if (contratoGasDetalle.Existe_ContratoDetalle_ConFinMundo(contratoGasDetalle.nif,
                                        contratoGasDetalle.cups20,
                                        contratoGasDetalle.tipo, contratoGasDetalle.fecha_inicio))

                                        contratoGasDetalle.Del(contratoGasDetalle.nif,
                                        contratoGasDetalle.cups20,
                                        contratoGasDetalle.tipo,
                                        contratoGasDetalle.fecha_inicio,
                                        new DateTime(4999, 12, 31).Date);


                                    contratoGasDetalle.Replace();
                                }

                            }
                    }
                }
                Cursor.Current = Cursors.Default;
            }
        }
    }



}
