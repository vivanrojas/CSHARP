namespace GestionOperaciones.forms.facturacion
{
    partial class FrmAgora
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAgora));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.mes_pdte = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.estado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_ps = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sum_tam = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.media_tam = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbl_total_cups = new System.Windows.Forms.Label();
            this.dgvd = new System.Windows.Forms.DataGridView();
            this.nif = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ccounips = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mes = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.estado_d = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tipo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tam = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cerrarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarExcelDetalleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.btnImportar = new System.Windows.Forms.Button();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvd)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 151);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1012, 451);
            this.tabControl1.TabIndex = 28;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1004, 425);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "R E S U M E N";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToOrderColumns = true;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.mes_pdte,
            this.estado,
            this.num_ps,
            this.sum_tam,
            this.media_tam});
            this.dgv.Location = new System.Drawing.Point(6, 6);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(660, 413);
            this.dgv.TabIndex = 67;
            // 
            // mes_pdte
            // 
            this.mes_pdte.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mes_pdte.DataPropertyName = "primer_mes_pdte";
            this.mes_pdte.HeaderText = "PRIMER MES PENDIENTE";
            this.mes_pdte.Name = "mes_pdte";
            this.mes_pdte.ReadOnly = true;
            this.mes_pdte.Width = 151;
            // 
            // estado
            // 
            this.estado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.estado.DataPropertyName = "ltp";
            this.estado.HeaderText = "ESTADO";
            this.estado.Name = "estado";
            this.estado.ReadOnly = true;
            this.estado.Width = 76;
            // 
            // num_ps
            // 
            this.num_ps.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.num_ps.DataPropertyName = "num_cups";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.num_ps.DefaultCellStyle = dataGridViewCellStyle1;
            this.num_ps.HeaderText = "Nº PS";
            this.num_ps.Name = "num_ps";
            this.num_ps.ReadOnly = true;
            this.num_ps.Width = 57;
            // 
            // sum_tam
            // 
            this.sum_tam.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.sum_tam.DataPropertyName = "sum_tam";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            this.sum_tam.DefaultCellStyle = dataGridViewCellStyle2;
            this.sum_tam.HeaderText = "Suma TAM (€)";
            this.sum_tam.Name = "sum_tam";
            this.sum_tam.ReadOnly = true;
            this.sum_tam.Width = 81;
            // 
            // media_tam
            // 
            this.media_tam.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.media_tam.DataPropertyName = "promedio_tam";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Format = "N0";
            dataGridViewCellStyle3.NullValue = null;
            this.media_tam.DefaultCellStyle = dataGridViewCellStyle3;
            this.media_tam.HeaderText = "PROMEDIO TAM (€)";
            this.media_tam.Name = "media_tam";
            this.media_tam.ReadOnly = true;
            this.media_tam.Width = 109;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lbl_total_cups);
            this.tabPage2.Controls.Add(this.dgvd);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1004, 425);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "D E T A L L E";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lbl_total_cups
            // 
            this.lbl_total_cups.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total_cups.AutoSize = true;
            this.lbl_total_cups.Location = new System.Drawing.Point(865, 14);
            this.lbl_total_cups.Name = "lbl_total_cups";
            this.lbl_total_cups.Size = new System.Drawing.Size(35, 13);
            this.lbl_total_cups.TabIndex = 67;
            this.lbl_total_cups.Text = "label1";
            // 
            // dgvd
            // 
            this.dgvd.AllowUserToAddRows = false;
            this.dgvd.AllowUserToOrderColumns = true;
            this.dgvd.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvd.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.nif,
            this.cliente,
            this.ccounips,
            this.mes,
            this.estado_d,
            this.tipo,
            this.tam});
            this.dgvd.Location = new System.Drawing.Point(6, 40);
            this.dgvd.Name = "dgvd";
            this.dgvd.Size = new System.Drawing.Size(992, 379);
            this.dgvd.TabIndex = 66;
            // 
            // nif
            // 
            this.nif.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.nif.DataPropertyName = "nif";
            this.nif.HeaderText = "NIF";
            this.nif.Name = "nif";
            this.nif.Width = 49;
            // 
            // cliente
            // 
            this.cliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cliente.DataPropertyName = "nombre_cliente";
            this.cliente.HeaderText = "CLIENTE";
            this.cliente.Name = "cliente";
            this.cliente.Width = 77;
            // 
            // ccounips
            // 
            this.ccounips.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ccounips.DataPropertyName = "cups13";
            this.ccounips.HeaderText = "CCOUNIPS";
            this.ccounips.Name = "ccounips";
            this.ccounips.Width = 87;
            // 
            // mes
            // 
            this.mes.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mes.DataPropertyName = "ultimo_mes_facturado";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.NullValue = null;
            this.mes.DefaultCellStyle = dataGridViewCellStyle4;
            this.mes.HeaderText = "MES";
            this.mes.Name = "mes";
            this.mes.Width = 55;
            // 
            // estado_d
            // 
            this.estado_d.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.estado_d.DataPropertyName = "estado_ltp";
            this.estado_d.HeaderText = "ESTADO";
            this.estado_d.Name = "estado_d";
            this.estado_d.Width = 76;
            // 
            // tipo
            // 
            this.tipo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tipo.DataPropertyName = "tipo";
            this.tipo.HeaderText = "TIPO";
            this.tipo.Name = "tipo";
            this.tipo.Width = 57;
            // 
            // tam
            // 
            this.tam.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tam.DataPropertyName = "tam";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N2";
            dataGridViewCellStyle5.NullValue = null;
            this.tam.DefaultCellStyle = dataGridViewCellStyle5;
            this.tam.HeaderText = "TAM";
            this.tam.Name = "tam";
            this.tam.Width = 55;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1036, 24);
            this.menuStrip1.TabIndex = 29;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cerrarToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // cerrarToolStripMenuItem
            // 
            this.cerrarToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_dialog_close_29299;
            this.cerrarToolStripMenuItem.Name = "cerrarToolStripMenuItem";
            this.cerrarToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.cerrarToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.cerrarToolStripMenuItem.Text = "Cerrar";
            this.cerrarToolStripMenuItem.Click += new System.EventHandler(this.cerrarToolStripMenuItem_Click);
            // 
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarExcelDetalleToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarExcelDetalleToolStripMenuItem
            // 
            this.importarExcelDetalleToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.import_icon;
            this.importarExcelDetalleToolStripMenuItem.Name = "importarExcelDetalleToolStripMenuItem";
            this.importarExcelDetalleToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.importarExcelDetalleToolStripMenuItem.Text = "Importar Excel Detalle";
            this.importarExcelDetalleToolStripMenuItem.Click += new System.EventHandler(this.importarExcelDetalleToolStripMenuItem_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem});
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem
            // 
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_help_browser_118806;
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem.Name = "acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem";
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem.Size = new System.Drawing.Size(284, 22);
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem.Text = "Acerca de Informe Estado Puntos Ágora";
            this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem.Click += new System.EventHandler(this.acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem_Click);
            // 
            // btnImportar
            // 
            this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImportar.Image = global::GestionOperaciones.Properties.Resources.Export;
            this.btnImportar.Location = new System.Drawing.Point(910, 39);
            this.btnImportar.Name = "btnImportar";
            this.btnImportar.Size = new System.Drawing.Size(110, 94);
            this.btnImportar.TabIndex = 30;
            this.toolTip.SetToolTip(this.btnImportar, "Importa la primera hoja de un excel y a partir de esos datos crea el informe RESU" +
        "MEN");
            this.btnImportar.UseVisualStyleBackColor = true;
            this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
            // 
            // cmdExcel
            // 
            this.cmdExcel.Enabled = false;
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(12, 117);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 27;
            this.toolTip.SetToolTip(this.cmdExcel, "Exporta a Excel el valor contenido en la hoja RESUMEN\r\n");
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(730, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(173, 16);
            this.label1.TabIndex = 31;
            this.label1.Text = "Importar Excel de Datos";
            // 
            // FrmAgora
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1036, 614);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnImportar);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.cmdExcel);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmAgora";
            this.Text = "Informe Estado Puntos Ágora";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmAgora_FormClosing);
            this.Load += new System.EventHandler(this.FrmAgora_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvd)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridView dgvd;
        private System.Windows.Forms.Label lbl_total_cups;
        private System.Windows.Forms.DataGridViewTextBoxColumn nif;
        private System.Windows.Forms.DataGridViewTextBoxColumn cliente;
        private System.Windows.Forms.DataGridViewTextBoxColumn ccounips;
        private System.Windows.Forms.DataGridViewTextBoxColumn mes;
        private System.Windows.Forms.DataGridViewTextBoxColumn estado_d;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipo;
        private System.Windows.Forms.DataGridViewTextBoxColumn tam;
        private System.Windows.Forms.DataGridViewTextBoxColumn mes_pdte;
        private System.Windows.Forms.DataGridViewTextBoxColumn estado;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_ps;
        private System.Windows.Forms.DataGridViewTextBoxColumn sum_tam;
        private System.Windows.Forms.DataGridViewTextBoxColumn media_tam;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarExcelDetalleToolStripMenuItem;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.ToolStripMenuItem cerrarToolStripMenuItem;
        private System.Windows.Forms.Button btnImportar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem acercaDeInformeEstadoPuntosÁgoraToolStripMenuItem;
    }
}