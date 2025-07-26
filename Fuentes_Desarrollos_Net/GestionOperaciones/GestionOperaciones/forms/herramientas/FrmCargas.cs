using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.herramientas
{
    public partial class FrmCargas : Form
    {
        EndesaBusiness.utilidades.Cargas cargas;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();

        public FrmCargas()
        {
            usage.Start("Herramientas", "FrmCargas" ,"N/A");
            InitializeComponent();
        }

        private void FrmCargas_Load(object sender, EventArgs e)
        {

            cargas = new EndesaBusiness.utilidades.Cargas();
            Cursor.Current = Cursors.WaitCursor;
            this.UpdateDates();
            Cursor.Current = Cursors.Default;

            
        }


        private void UpdateDates()
        {
            DateTime fecha = new DateTime();
            int numReg = 0;
            int semanal = 7;
            double free_space = 0;
            DateTime ultimoDiaHabil = new DateTime();

            


            ultimoDiaHabil = EndesaBusiness.utilidades.Fichero.UltimoDiaHabil();

            
            #region Facturacion
            
            fecha = cargas.UltimaActualizacionCuadroMando("CuadroDeMando");
            this.lblCuadroDeMando.Text = string.Format("Cuadro de Mando: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblCuadroDeMando.ForeColor = Color.Green;
            else
                this.lblCuadroDeMando.ForeColor = Color.Red;

            fecha = cargas.UltimaActualizacionCuadroMando("PdteWeb");
            this.lblPdteWeb.Text = string.Format("Pdte Web: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblPdteWeb.ForeColor = Color.Green;
            else
                this.lblPdteWeb.ForeColor = Color.Red;

            fecha = cargas.UltimaActualizacionCuadroMando("FacturasSIGAME");
            this.lblFacturasSIGAME.Text = string.Format("Facturas SIGAME: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblFacturasSIGAME.ForeColor = Color.Green;
            else
                this.lblFacturasSIGAME.ForeColor = Color.Red;

            fecha = cargas.UltimaActualizacionCuadroMando("NRI");
            this.lblNRIs.Text = string.Format("NRI`s: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblNRIs.ForeColor = Color.Green;
            else
                this.lblNRIs.ForeColor = Color.Red;

            fecha = cargas.UltimaActualizacionCuadroMando("PuntosActivosGas");
            this.lblActivosGas.Text = string.Format("Puntos Activos Gas: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblActivosGas.ForeColor = Color.Green;
            else
                this.lblActivosGas.ForeColor = Color.Red;

            fecha = cargas.UltimaActualizacionCuadroMando("ReclamacionesGas");
            this.lblReclamacionesGas.Text = string.Format("Reclamaciones Gas: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblReclamacionesGas.ForeColor = Color.Green;
            else
                this.lblReclamacionesGas.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_Alarmas();
            this.lblAlarmas.Text = string.Format("Alarmas: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblAlarmas.ForeColor = Color.Green;
            else
                this.lblAlarmas.ForeColor = Color.Red;

            // Fechas de facturas

            UltimaFechaFactura();

            // Contratos Ágora

            fecha = cargas.FechaActualizacion_Contratos_Agora();
            this.lbl_contratos_agora.Text = string.Format("Contratos Ágora: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lbl_contratos_agora.ForeColor = Color.Green;
            else
                this.lbl_contratos_agora.ForeColor = Color.Red;

            #endregion
            #region Contratacion

            fecha = EndesaBusiness.utilidades.FechasCargas.UltimaActualizacionPS("gestionATR");
            this.lblGestionPropia.Text = string.Format("Gestión Propia ATR: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblGestionPropia.ForeColor = Color.Green;
            else
                this.lblGestionPropia.ForeColor = Color.Red;

            fecha = EndesaBusiness.utilidades.FechasCargas.UltimaActualizacionPS("PS");
            this.lblPS.Text = string.Format("PS: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblPS.ForeColor = Color.Green;
            else
                this.lblPS.ForeColor = Color.Red;

            fecha = EndesaBusiness.utilidades.FechasCargas.UltimaActualizacionPS("PSAT");
            this.lblPSAT.Text = string.Format("PS_AT: {0}", fecha);
            if (fecha.Date == DateTime.Now.Date)
                this.lblPSAT.ForeColor = Color.Green;
            else
                this.lblPSAT.ForeColor = Color.Red;

            numReg = cargas.NumRegTabla("PS_AT", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
            this.lblPS_AT_Count.Text = string.Format("PS_AT Total Registros: {0:#,##0},", numReg);

            numReg = cargas.NumRegTabla("PS_AT_Ant", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
            this.lbl_PS_AT_Ant_Count.Text = string.Format("PS_AT Anterior Total Registros: {0:#,##0},", numReg);

            fecha = cargas.FechaActualizacion_cont_ps_at_mt();
            this.lblPSATMT.Text = string.Format("PS_AT_MT: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days < semanal)
                this.lblPSATMT.ForeColor = Color.Green;
            else
                this.lblPSATMT.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_contratosPS();
            this.lblcontratosPS.Text = string.Format("contratosPS: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days < semanal)
                this.lblcontratosPS.ForeColor = Color.Green;
            else
                this.lblcontratosPS.ForeColor = Color.Red;


            fecha = cargas.UltimaActualizacionExtraccionAutoconsumo();
            this.lblExtraccionAutoconsumo.Text = string.Format("Extracción Autoconsumo: {0}", fecha.ToString("dd/MM/yyyy"));
            if ((DateTime.Now.Date - fecha.Date).Days < semanal)
                this.lblExtraccionAutoconsumo.ForeColor = Color.Green;
            else
                this.lblExtraccionAutoconsumo.ForeColor = Color.Red;

            #endregion

            #region Medida
            //fecha = FechaActualizacion_CurvasDatamart();
            //this.lbl_dt.Text = string.Format("Última importación: {0}", fecha);
            //if ((DateTime.Now.Date - fecha.Date).Days == 0)
            //    this.lbl_dt.ForeColor = Color.Green;
            //else
            //    this.lbl_dt.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_extraccion_ch");
            this.lbl_kee_reporte_extraccion_ch.Text = string.Format("kee_reporte_extraccion_ch: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_extraccion_ch.ForeColor = Color.Green;
            else if(ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_extraccion_ch.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_extraccion_ch.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_extraccion_cch");
            this.lbl_kee_reporte_extraccion_cch.Text = string.Format("kee_reporte_extraccion_cch: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_extraccion_cch.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_extraccion_cch.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_extraccion_cch.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_exabeat");
            this.lbl_kee_reporte_exabeat.Text = string.Format("kee_reporte_exabeat: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_exabeat.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_exabeat.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_exabeat.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_ftp");
            this.lbl_kee_reporte_ftp.Text = string.Format("kee_reporte_ftp: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_ftp.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_ftp.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_ftp.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_portugal");
            this.lbl_kee_reporte_portugal.Text = string.Format("kee_reporte_portugal: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_portugal.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_portugal.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_portugal.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_publicada");
            this.lbl_kee_reporte_publicada.Text = string.Format("kee_reporte_publicada: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_publicada.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_publicada.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_publicada.ForeColor = Color.Red;

            fecha = cargas.FechaActualizacion_tablas_kee("kee_reporte_starbeat");
            this.lbl_kee_reporte_starbeat.Text = string.Format("kee_reporte_starbeat: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_kee_reporte_starbeat.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_kee_reporte_starbeat.ForeColor = Color.RosyBrown;
            else
                this.lbl_kee_reporte_starbeat.ForeColor = Color.Red;


            fecha = cargas.Fecha_UltimaActualizacionMedida("dt_t_ed_h_rcurvas", "fh_ult_modif_sce");
            this.lbl_dt_t_ed_h_rcurvas.Text = string.Format("dt_t_ed_h_rcurvas: {0}", fecha);
            if ((DateTime.Now.Date - fecha.Date).Days == 0)
                this.lbl_dt_t_ed_h_rcurvas.ForeColor = Color.Green;
            else if (ultimoDiaHabil.Date == fecha.Date)
                this.lbl_dt_t_ed_h_rcurvas.ForeColor = Color.RosyBrown;
            else
                this.lbl_dt_t_ed_h_rcurvas.ForeColor = Color.Red;


            #endregion

            // Espacio Libre Unidades
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady && drive.Name == @"M:\")
                {
                    free_space = Convert.ToDouble(drive.TotalFreeSpace);
                    free_space = free_space / 1024;
                    free_space = free_space / 1024;
                    free_space = free_space / 1024;
                    lbl_unidad_m.Text = string.Format("Unidad M Operaciones: {0:#,##0.00}", free_space) + " GB";
                    //lbl_unidad_m.Text(((Convert.ToDouble(drive.TotalFreeSpace) / 1024) / 1024) / 1024);
                }
            }
            Console.Write(-1);
            

        }
        
        private void UltimaFechaFactura()
        {
            DateTime fecha = new DateTime();

            fecha = cargas.GetFechaUltimaFactura(1);
            this.lbl_MTEspana.Text = string.Format("MT-España: {0}", fecha.ToString("dd/MM/yyyy"));
            if (fecha.Date == DateTime.Now.Date.AddDays(-1))
                this.lbl_MTEspana.ForeColor = Color.Green;
            else
                this.lbl_MTEspana.ForeColor = Color.Red;
            
                           
            fecha = cargas.GetFechaUltimaFactura(2);
            this.lbl_EEXXI.Text = string.Format("EEXXI: {0}", fecha.ToString("dd/MM/yyyy"));
            if (fecha.Date == DateTime.Now.Date.AddDays(-1))
                this.lbl_EEXXI.ForeColor = Color.Green;
            else
                this.lbl_EEXXI.ForeColor = Color.Red;


            fecha = cargas.GetFechaUltimaFactura(3);
            this.lbl_BTN_Portugal.Text = string.Format("BTN-Portugal: {0}", fecha.ToString("dd/MM/yyyy"));
            if (fecha.Date == DateTime.Now.Date.AddDays(-1))
                this.lbl_BTN_Portugal.ForeColor = Color.Green;
            else
                this.lbl_BTN_Portugal.ForeColor = Color.Red;


            fecha = cargas.GetFechaUltimaFactura(4);
            this.lbl_BTE_Portugal.Text = string.Format("BTE-Portugal: {0}", fecha.ToString("dd/MM/yyyy"));
            if (fecha.Date == DateTime.Now.Date.AddDays(-1))
                this.lbl_BTE_Portugal.ForeColor = Color.Green;
            else
                this.lbl_BTE_Portugal.ForeColor = Color.Red;


            fecha = cargas.GetFechaUltimaFactura(5);
            this.lbl_MT_Portugal.Text = string.Format("MT-Portugal: {0}", fecha.ToString("dd/MM/yyyy"));
            if (fecha.Date == DateTime.Now.Date.AddDays(-1))
                this.lbl_MT_Portugal.ForeColor = Color.Green;
            else
                this.lbl_MT_Portugal.ForeColor = Color.Red;
            

        }


        //private DateTime UltimaActualizacionPS(string nombreProceso)
        //{
        //    MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader reader;
        //    string strSql;
        //    DateTime fecha = new DateTime();

        //    try
        //    {
        //        strSql = "select f_ult_mod from ps_fechas_procesos where"
        //            + " proceso = '" + nombreProceso + "'";

        //        db = new MySQLDB(MySQLDB.Esquemas.CON);
        //        command = new MySqlCommand(strSql, db.con);
        //        reader = command.ExecuteReader();
        //        while (reader.Read())
        //        {
        //            fecha = Convert.ToDateTime(reader["f_ult_mod"]);
        //        }

        //        db.CloseConnection();
        //        return fecha;

        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message,
        //     "UltimaActualizacionCuadroMando: " + nombreProceso,
        //     MessageBoxButtons.OK,
        //      MessageBoxIcon.Error);
        //        return new DateTime(1899, 01, 01);

        //    }
        //}

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            this.UpdateDates();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void LblExtraccionAutoconsumo_Click(object sender, EventArgs e)
        {

        }

        private void FrmCargas_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Herramientas", "FrmCargas" ,"N/A");
        }
    }
}
