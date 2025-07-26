using EndesaBusiness.utilidades;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace GestionOperaciones.forms.herramientas
{
    public partial class FrmCalendarios : Form
    {
        EndesaBusiness.calendarios.CalendarioTiposPortugal tc;
        EndesaBusiness.punto_suministro.Territorios territorios;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmCalendarios()
        {
            usage.Start("Herramientas", "FrmCalendarios" ,"N/A");
            InitializeComponent();
        }

        private void FrmCalendarios_Load(object sender, EventArgs e)
        {
            rdb_cuartohoraria.Checked = true;
            txt_fd.Value = new DateTime(DateTime.Now.Year, 1, 1);
            txt_fh.Value = new DateTime(DateTime.Now.Year, 12, 31);
            territorios = new EndesaBusiness.punto_suministro.Territorios();
            CargarCombos();            
        }

        private void CargarCombos()
        {

            for (int i = cmbTarifas.Items.Count - 1; i == 0; i--)
            {
                cmbTarifas.Items.RemoveAt(i);
            }

            EndesaBusiness.punto_suministro.Tarifa tarifas
                = new EndesaBusiness.punto_suministro.Tarifa(txt_fd.Value, txt_fh.Value);
            

            foreach (KeyValuePair<string, EndesaEntity.punto_suministro.Tarifa> p in tarifas.dic)
                cmbTarifas.Items.Add(p.Value.tarifa);

        }

       

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            

            FileInfo fileInfo = new FileInfo(@"c:\Temp\Calendario_ES_"
                + cmbTarifas.SelectedItem + "_"
                + (rdb_horaria.Checked ? "Horario_" : "CuartoHorario_")
                + txt_fd.Value.ToString("yyyyMMdd") + "_"
                + txt_fh.Value.ToString("yyyyMMdd") + "_"
                + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
            
                      

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            save.FileName = fileInfo.Name;
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                ExportarExcel(save.FileName);                
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }

                


            
            else
                MessageBox.Show("No hay datos para exportar a Excel.", "Abrir Excel generado",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);




            //MessageBox.Show("Fin Exportación."
            //    + System.Environment.NewLine
            //    + "Se ha generado el archivo: "
            //    + fileInfo.FullName,
            //    "Calendarios Tarifarios",
            //    MessageBoxButtons.OK,
            //    MessageBoxIcon.Information);

        }

        private void ExportarExcel(string fichero)
        {
            int f = 0;
            int c = 0;

            FileInfo fileInfoExcel = new FileInfo(fichero);
            if (fileInfoExcel.Exists)
                fileInfoExcel.Delete();

            EndesaBusiness.calendarios.Calendario cc =
           new EndesaBusiness.calendarios.Calendario(txt_fd.Value, txt_fh.Value,
           cmbTarifas.SelectedItem.ToString(), cmbTerritorios.SelectedItem.ToString());


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfoExcel);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CALENDARIO");

            var headerCells = workSheet.Cells[1, 1, 1, 6];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;
            workSheet.Cells[f, c].Value = "DIA"; c++;
            workSheet.Cells[f, c].Value = "HORA"; c++;
            workSheet.Cells[f, c].Value = "PERIODO"; c++;

            if (rdb_horaria.Checked)
            {
                for (int i = 1; i <= cc.numPeriodosMedidaHorario; i++)
                {
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = cc.vectorFechas[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = cc.vectorHoras[i]; c++;
                    workSheet.Cells[f, c].Value = cc.vectorPeriodosTarifarios[i]; c++;
                }
            }
            else
            {
                for (int i = 1; i <= cc.numPeriodosMedidaCuartoHorario; i++)
                {
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = cc.vectorFechasCuartoHorarias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = cc.vectorHorasCuartoHorarias[i];
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortTimePattern; c++;
                    workSheet.Cells[f, c].Value = cc.vectorPeriodosTarifariosCuartoHorarios[i]; c++;
                }
            }



            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.Cells["A1:C1"].AutoFilter = true;
            workSheet.View.FreezePanes(2, 1);
            
            allCells.AutoFitColumns();
            excelPackage.Save();
        }

        

        private void cmbTarifas_SelectedIndexChanged(object sender, EventArgs e)
        {

            cmbTerritorios.Items.Clear();

            

            List<string> lista_territorios = territorios.GetTerritorios(cmbTarifas.SelectedItem.ToString());
            foreach(string p in lista_territorios)
                cmbTerritorios.Items.Add(p);
        }

        private void rdb_horaria_CheckedChanged(object sender, EventArgs e)
        {
            rdb_cuartohoraria.Checked = rdb_horaria.Checked == false;
        }

        private void rdb_cuartohoraria_CheckedChanged(object sender, EventArgs e)
        {
            rdb_horaria.Checked = rdb_cuartohoraria.Checked == false;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmCalendarios_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Herramientas", "FrmCalendarios" ,"N/A");
        }
    }
}
