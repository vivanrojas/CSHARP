namespace GestionOperaciones.forms.facturacion
{
    partial class FrmFacturasBTN_CAPB
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFacturasBTN_CAPB));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lbl_total = new System.Windows.Forms.Label();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cnifdnic = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dapersoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.distribuidora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentarios_descuadres = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentarios_contratacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tramitacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ffacthas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CAPB = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.perfil = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.calendario = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tarifa = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chk_ultima_factura_periodo = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_cups20 = new System.Windows.Forms.TextBox();
            this.btn_buscar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_fecha_factura_hasta = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_fecha_factura_desde = new System.Windows.Forms.DateTimePicker();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1241, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.opcionesToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // opcionesToolStripMenuItem
            // 
            this.opcionesToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_settings_2199094;
            this.opcionesToolStripMenuItem.Name = "opcionesToolStripMenuItem";
            this.opcionesToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.opcionesToolStripMenuItem.Text = "Opciones";
            this.opcionesToolStripMenuItem.Click += new System.EventHandler(this.opcionesToolStripMenuItem_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(12, 170);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1217, 580);
            this.tabControl1.TabIndex = 32;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lbl_total);
            this.tabPage1.Controls.Add(this.cmdExcel);
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1209, 554);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Facturas";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // lbl_total
            // 
            this.lbl_total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total.AutoSize = true;
            this.lbl_total.Location = new System.Drawing.Point(1052, 14);
            this.lbl_total.Name = "lbl_total";
            this.lbl_total.Size = new System.Drawing.Size(160, 13);
            this.lbl_total.TabIndex = 66;
            this.lbl_total.Text = "Total registros:                 ";
            this.lbl_total.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(6, 6);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 63;
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cnifdnic,
            this.dapersoc,
            this.distribuidora,
            this.cups20,
            this.comentarios_descuadres,
            this.comentarios_contratacion,
            this.tramitacion,
            this.ffacthas,
            this.CAPB,
            this.perfil,
            this.calendario,
            this.tarifa});
            this.dgv.Location = new System.Drawing.Point(6, 40);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(1197, 508);
            this.dgv.TabIndex = 62;
            // 
            // cnifdnic
            // 
            this.cnifdnic.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cnifdnic.DataPropertyName = "cupsree";
            this.cnifdnic.HeaderText = "CUPS";
            this.cnifdnic.Name = "cnifdnic";
            this.cnifdnic.ReadOnly = true;
            this.cnifdnic.Width = 65;
            // 
            // dapersoc
            // 
            this.dapersoc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dapersoc.DataPropertyName = "cfactura";
            this.dapersoc.HeaderText = "CFACTURA";
            this.dapersoc.Name = "dapersoc";
            this.dapersoc.ReadOnly = true;
            this.dapersoc.Width = 97;
            // 
            // distribuidora
            // 
            this.distribuidora.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.distribuidora.DataPropertyName = "creferen";
            this.distribuidora.HeaderText = "Ref,";
            this.distribuidora.Name = "distribuidora";
            this.distribuidora.ReadOnly = true;
            this.distribuidora.Width = 56;
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "secfactu";
            this.cups20.HeaderText = "Sec.";
            this.cups20.Name = "cups20";
            this.cups20.ReadOnly = true;
            this.cups20.Width = 58;
            // 
            // comentarios_descuadres
            // 
            this.comentarios_descuadres.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios_descuadres.DataPropertyName = "vcuovafa";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            this.comentarios_descuadres.DefaultCellStyle = dataGridViewCellStyle2;
            this.comentarios_descuadres.HeaderText = "Consumo";
            this.comentarios_descuadres.Name = "comentarios_descuadres";
            this.comentarios_descuadres.ReadOnly = true;
            this.comentarios_descuadres.Width = 83;
            // 
            // comentarios_contratacion
            // 
            this.comentarios_contratacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios_contratacion.DataPropertyName = "ffactura";
            this.comentarios_contratacion.HeaderText = "Fecha emisión";
            this.comentarios_contratacion.Name = "comentarios_contratacion";
            this.comentarios_contratacion.ReadOnly = true;
            this.comentarios_contratacion.Width = 104;
            // 
            // tramitacion
            // 
            this.tramitacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tramitacion.DataPropertyName = "ffactdes";
            this.tramitacion.HeaderText = "Periodo Desde";
            this.tramitacion.Name = "tramitacion";
            this.tramitacion.ReadOnly = true;
            this.tramitacion.Width = 106;
            // 
            // ffacthas
            // 
            this.ffacthas.DataPropertyName = "ffacthas";
            this.ffacthas.HeaderText = "Periodo Hasta";
            this.ffacthas.Name = "ffacthas";
            this.ffacthas.ReadOnly = true;
            // 
            // CAPB
            // 
            this.CAPB.DataPropertyName = "ifactura";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Format = "C2";
            dataGridViewCellStyle3.NullValue = null;
            this.CAPB.DefaultCellStyle = dataGridViewCellStyle3;
            this.CAPB.HeaderText = "CAPB";
            this.CAPB.Name = "CAPB";
            this.CAPB.ReadOnly = true;
            // 
            // perfil
            // 
            this.perfil.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.perfil.DataPropertyName = "perfil";
            this.perfil.HeaderText = "PERFIL";
            this.perfil.Name = "perfil";
            this.perfil.ReadOnly = true;
            this.perfil.Width = 75;
            // 
            // calendario
            // 
            this.calendario.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.calendario.DataPropertyName = "calendario";
            this.calendario.HeaderText = "CALENDARIO";
            this.calendario.Name = "calendario";
            this.calendario.ReadOnly = true;
            this.calendario.Width = 111;
            // 
            // tarifa
            // 
            this.tarifa.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tarifa.DataPropertyName = "tarifa";
            this.tarifa.HeaderText = "TARIFA";
            this.tarifa.Name = "tarifa";
            this.tarifa.ReadOnly = true;
            this.tarifa.Width = 76;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chk_ultima_factura_periodo);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.txt_cups20);
            this.groupBox2.Controls.Add(this.btn_buscar);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txt_fecha_factura_hasta);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txt_fecha_factura_desde);
            this.groupBox2.Location = new System.Drawing.Point(12, 27);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(567, 137);
            this.groupBox2.TabIndex = 33;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Filtros";
            // 
            // chk_ultima_factura_periodo
            // 
            this.chk_ultima_factura_periodo.AutoSize = true;
            this.chk_ultima_factura_periodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chk_ultima_factura_periodo.Location = new System.Drawing.Point(150, 99);
            this.chk_ultima_factura_periodo.Name = "chk_ultima_factura_periodo";
            this.chk_ultima_factura_periodo.Size = new System.Drawing.Size(244, 17);
            this.chk_ultima_factura_periodo.TabIndex = 30;
            this.chk_ultima_factura_periodo.Text = "Sólo última factura emitida por periodo";
            this.chk_ultima_factura_periodo.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(90, 76);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "CUPS20";
            // 
            // txt_cups20
            // 
            this.txt_cups20.Location = new System.Drawing.Point(150, 73);
            this.txt_cups20.Name = "txt_cups20";
            this.txt_cups20.Size = new System.Drawing.Size(146, 20);
            this.txt_cups20.TabIndex = 28;
            // 
            // btn_buscar
            // 
            this.btn_buscar.Image = global::GestionOperaciones.Properties.Resources.buscar;
            this.btn_buscar.Location = new System.Drawing.Point(486, 38);
            this.btn_buscar.Name = "btn_buscar";
            this.btn_buscar.Size = new System.Drawing.Size(44, 42);
            this.btn_buscar.TabIndex = 27;
            this.btn_buscar.UseVisualStyleBackColor = true;
            this.btn_buscar.Click += new System.EventHandler(this.btn_buscar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(247, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 13);
            this.label1.TabIndex = 24;
            this.label1.Text = "Hasta Fecha Consumo";
            // 
            // txt_fecha_factura_hasta
            // 
            this.txt_fecha_factura_hasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_factura_hasta.Location = new System.Drawing.Point(387, 47);
            this.txt_fecha_factura_hasta.Name = "txt_fecha_factura_hasta";
            this.txt_fecha_factura_hasta.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_factura_hasta.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(7, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Desde Fecha Consumo";
            // 
            // txt_fecha_factura_desde
            // 
            this.txt_fecha_factura_desde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_factura_desde.Location = new System.Drawing.Point(150, 47);
            this.txt_fecha_factura_desde.Name = "txt_fecha_factura_desde";
            this.txt_fecha_factura_desde.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_factura_desde.TabIndex = 21;
            // 
            // FrmFacturasBTN_CAPB
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1241, 762);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmFacturasBTN_CAPB";
            this.Text = "Facturas BTN - CAPB";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmFacturasBTN_CAPB_FormClosing);
            this.Load += new System.EventHandler(this.FrmFacturasBTN_CAPB_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label lbl_total;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_cups20;
        private System.Windows.Forms.Button btn_buscar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker txt_fecha_factura_hasta;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker txt_fecha_factura_desde;
        private System.Windows.Forms.DataGridViewTextBoxColumn cnifdnic;
        private System.Windows.Forms.DataGridViewTextBoxColumn dapersoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn distribuidora;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios_descuadres;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios_contratacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn tramitacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn ffacthas;
        private System.Windows.Forms.DataGridViewTextBoxColumn CAPB;
        private System.Windows.Forms.DataGridViewTextBoxColumn perfil;
        private System.Windows.Forms.DataGridViewTextBoxColumn calendario;
        private System.Windows.Forms.DataGridViewTextBoxColumn tarifa;
        private System.Windows.Forms.CheckBox chk_ultima_factura_periodo;
    }
}