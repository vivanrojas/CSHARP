namespace GestionOperaciones.forms.contratacion.gas
{
    partial class FrmReubicaciones
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmReubicaciones));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarXMLToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_fecha_hasta = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_fecha_desde = new System.Windows.Forms.DateTimePicker();
            this.cmdSearch = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.lbl_total_reubicaciones = new System.Windows.Forms.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cnifdnic = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dapersoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.distribuidora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentarios_descuadres = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentarios_contratacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_reubicacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lbl_total_registros_cargas = new System.Windows.Forms.Label();
            this.dgvCargas = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_archivos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_minima_solicitud = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_maxima_solicitud = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chk_all = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCargas)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1020, 24);
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
            this.importarXMLToolStripMenuItem,
            this.importarXMLToolStripMenuItem1,
            this.toolStripSeparator1,
            this.opcionesToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarXMLToolStripMenuItem
            // 
            this.importarXMLToolStripMenuItem.Name = "importarXMLToolStripMenuItem";
            this.importarXMLToolStripMenuItem.Size = new System.Drawing.Size(223, 22);
            this.importarXMLToolStripMenuItem.Text = "Importar XML desde ZIP";
            this.importarXMLToolStripMenuItem.Click += new System.EventHandler(this.importarXMLToolStripMenuItem_Click);
            // 
            // importarXMLToolStripMenuItem1
            // 
            this.importarXMLToolStripMenuItem1.Name = "importarXMLToolStripMenuItem1";
            this.importarXMLToolStripMenuItem1.Size = new System.Drawing.Size(223, 22);
            this.importarXMLToolStripMenuItem1.Text = "Importar XML desde carpeta";
            this.importarXMLToolStripMenuItem1.Click += new System.EventHandler(this.importarXMLToolStripMenuItem1_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(220, 6);
            // 
            // opcionesToolStripMenuItem
            // 
            this.opcionesToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_settings_2199094;
            this.opcionesToolStripMenuItem.Name = "opcionesToolStripMenuItem";
            this.opcionesToolStripMenuItem.Size = new System.Drawing.Size(223, 22);
            this.opcionesToolStripMenuItem.Text = "Opciones";
            this.opcionesToolStripMenuItem.Click += new System.EventHandler(this.opcionesToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chk_all);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txt_fecha_hasta);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txt_fecha_desde);
            this.groupBox1.Controls.Add(this.cmdSearch);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(732, 97);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(285, 42);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 13);
            this.label5.TabIndex = 100;
            this.label5.Text = "Vigor Hasta:";
            // 
            // txt_fecha_hasta
            // 
            this.txt_fecha_hasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_hasta.Location = new System.Drawing.Point(353, 39);
            this.txt_fecha_hasta.Name = "txt_fecha_hasta";
            this.txt_fecha_hasta.Size = new System.Drawing.Size(102, 20);
            this.txt_fecha_hasta.TabIndex = 99;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(50, 42);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 98;
            this.label6.Text = "Vigor Desde:";
            // 
            // txt_fecha_desde
            // 
            this.txt_fecha_desde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_desde.Location = new System.Drawing.Point(121, 39);
            this.txt_fecha_desde.Name = "txt_fecha_desde";
            this.txt_fecha_desde.Size = new System.Drawing.Size(102, 20);
            this.txt_fecha_desde.TabIndex = 97;
            // 
            // cmdSearch
            // 
            this.cmdSearch.Image = ((System.Drawing.Image)(resources.GetObject("cmdSearch.Image")));
            this.cmdSearch.Location = new System.Drawing.Point(485, 26);
            this.cmdSearch.Name = "cmdSearch";
            this.cmdSearch.Size = new System.Drawing.Size(69, 46);
            this.cmdSearch.TabIndex = 96;
            this.cmdSearch.UseVisualStyleBackColor = true;
            this.cmdSearch.Click += new System.EventHandler(this.cmdSearch_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 130);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1001, 462);
            this.tabControl1.TabIndex = 5;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.cmdExcel);
            this.tabPage2.Controls.Add(this.lbl_total_reubicaciones);
            this.tabPage2.Controls.Add(this.dgv);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(993, 436);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Contratos";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(6, 6);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 62;
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // lbl_total_reubicaciones
            // 
            this.lbl_total_reubicaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total_reubicaciones.AutoSize = true;
            this.lbl_total_reubicaciones.Location = new System.Drawing.Point(824, 14);
            this.lbl_total_reubicaciones.Name = "lbl_total_reubicaciones";
            this.lbl_total_reubicaciones.Size = new System.Drawing.Size(127, 13);
            this.lbl_total_reubicaciones.TabIndex = 61;
            this.lbl_total_reubicaciones.Text = "Total registros:                 ";
            this.lbl_total_reubicaciones.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
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
            this.comentarios_descuadres,
            this.comentarios_contratacion,
            this.cups20,
            this.fecha_reubicacion});
            this.dgv.Location = new System.Drawing.Point(6, 40);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(981, 390);
            this.dgv.TabIndex = 60;
            // 
            // cnifdnic
            // 
            this.cnifdnic.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cnifdnic.DataPropertyName = "cups";
            this.cnifdnic.HeaderText = "CUPS";
            this.cnifdnic.Name = "cnifdnic";
            this.cnifdnic.ReadOnly = true;
            this.cnifdnic.Width = 61;
            // 
            // dapersoc
            // 
            this.dapersoc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dapersoc.DataPropertyName = "tipo";
            this.dapersoc.HeaderText = "Tipo";
            this.dapersoc.Name = "dapersoc";
            this.dapersoc.ReadOnly = true;
            this.dapersoc.Width = 53;
            // 
            // distribuidora
            // 
            this.distribuidora.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.distribuidora.DataPropertyName = "tarifa_sigame";
            this.distribuidora.HeaderText = "Tarifa SIGAME";
            this.distribuidora.Name = "distribuidora";
            this.distribuidora.ReadOnly = true;
            this.distribuidora.Width = 95;
            // 
            // comentarios_descuadres
            // 
            this.comentarios_descuadres.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios_descuadres.DataPropertyName = "fecha_desde";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.comentarios_descuadres.DefaultCellStyle = dataGridViewCellStyle2;
            this.comentarios_descuadres.HeaderText = "Fecha Desde Addenda";
            this.comentarios_descuadres.Name = "comentarios_descuadres";
            this.comentarios_descuadres.ReadOnly = true;
            this.comentarios_descuadres.Width = 130;
            // 
            // comentarios_contratacion
            // 
            this.comentarios_contratacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios_contratacion.DataPropertyName = "fecha_hasta";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.comentarios_contratacion.DefaultCellStyle = dataGridViewCellStyle3;
            this.comentarios_contratacion.HeaderText = "Fecha Hasta Addenda";
            this.comentarios_contratacion.Name = "comentarios_contratacion";
            this.comentarios_contratacion.ReadOnly = true;
            this.comentarios_contratacion.Width = 127;
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "tarifa_reubicacion";
            this.cups20.HeaderText = "Tarifa Reubicación";
            this.cups20.Name = "cups20";
            this.cups20.ReadOnly = true;
            this.cups20.Width = 112;
            // 
            // fecha_reubicacion
            // 
            this.fecha_reubicacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_reubicacion.DataPropertyName = "fecha_reubicacion";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fecha_reubicacion.DefaultCellStyle = dataGridViewCellStyle4;
            this.fecha_reubicacion.HeaderText = "Fecha Reubicación";
            this.fecha_reubicacion.Name = "fecha_reubicacion";
            this.fecha_reubicacion.ReadOnly = true;
            this.fecha_reubicacion.Width = 115;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.lbl_total_registros_cargas);
            this.tabPage1.Controls.Add(this.dgvCargas);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(993, 436);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "Cargas";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(50, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 13);
            this.label2.TabIndex = 66;
            this.label2.Text = "Añadir registro sin archivos ";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button1
            // 
            this.button1.Image = global::GestionOperaciones.Properties.Resources.if_f_right_256_282463;
            this.button1.Location = new System.Drawing.Point(6, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(38, 35);
            this.button1.TabIndex = 65;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lbl_total_registros_cargas
            // 
            this.lbl_total_registros_cargas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total_registros_cargas.AutoSize = true;
            this.lbl_total_registros_cargas.Location = new System.Drawing.Point(860, 44);
            this.lbl_total_registros_cargas.Name = "lbl_total_registros_cargas";
            this.lbl_total_registros_cargas.Size = new System.Drawing.Size(127, 13);
            this.lbl_total_registros_cargas.TabIndex = 64;
            this.lbl_total_registros_cargas.Text = "Total registros:                 ";
            this.lbl_total_registros_cargas.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dgvCargas
            // 
            this.dgvCargas.AllowUserToAddRows = false;
            this.dgvCargas.AllowUserToDeleteRows = false;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dgvCargas.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle5;
            this.dgvCargas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvCargas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCargas.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.num_archivos,
            this.fecha_minima_solicitud,
            this.fecha_maxima_solicitud});
            this.dgvCargas.Location = new System.Drawing.Point(6, 63);
            this.dgvCargas.MultiSelect = false;
            this.dgvCargas.Name = "dgvCargas";
            this.dgvCargas.Size = new System.Drawing.Size(981, 367);
            this.dgvCargas.TabIndex = 63;
            this.dgvCargas.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvCargas_CellValueChanged);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridViewTextBoxColumn1.DataPropertyName = "fecha_carga";
            this.dataGridViewTextBoxColumn1.HeaderText = "Fecha Carga";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.Width = 86;
            // 
            // num_archivos
            // 
            this.num_archivos.DataPropertyName = "num_archivos";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.num_archivos.DefaultCellStyle = dataGridViewCellStyle6;
            this.num_archivos.HeaderText = "Nº Archivos";
            this.num_archivos.Name = "num_archivos";
            // 
            // fecha_minima_solicitud
            // 
            this.fecha_minima_solicitud.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_minima_solicitud.DataPropertyName = "fecha_minima_solicitud";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fecha_minima_solicitud.DefaultCellStyle = dataGridViewCellStyle7;
            this.fecha_minima_solicitud.HeaderText = "Fecha mínima solicitud";
            this.fecha_minima_solicitud.Name = "fecha_minima_solicitud";
            this.fecha_minima_solicitud.Width = 128;
            // 
            // fecha_maxima_solicitud
            // 
            this.fecha_maxima_solicitud.DataPropertyName = "fecha_maxima_solicitud";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fecha_maxima_solicitud.DefaultCellStyle = dataGridViewCellStyle8;
            this.fecha_maxima_solicitud.HeaderText = "Fecha máxima solicitud";
            this.fecha_maxima_solicitud.Name = "fecha_maxima_solicitud";
            // 
            // chk_all
            // 
            this.chk_all.AutoSize = true;
            this.chk_all.Location = new System.Drawing.Point(121, 67);
            this.chk_all.Name = "chk_all";
            this.chk_all.Size = new System.Drawing.Size(271, 17);
            this.chk_all.TabIndex = 101;
            this.chk_all.Text = "Mostrar reubicaciones con tarifas iguales a SIGAME";
            this.chk_all.UseVisualStyleBackColor = true;
            // 
            // FrmReubicaciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1020, 604);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmReubicaciones";
            this.Text = "Reubicaciones";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmReubicaciones_FormClosing);
            this.Load += new System.EventHandler(this.FrmReubicaciones_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCargas)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button cmdSearch;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker txt_fecha_hasta;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker txt_fecha_desde;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Label lbl_total_reubicaciones;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarXMLToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarXMLToolStripMenuItem1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lbl_total_registros_cargas;
        private System.Windows.Forms.DataGridView dgvCargas;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_archivos;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_minima_solicitud;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_maxima_solicitud;
        private System.Windows.Forms.DataGridViewTextBoxColumn cnifdnic;
        private System.Windows.Forms.DataGridViewTextBoxColumn dapersoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn distribuidora;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios_descuadres;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios_contratacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_reubicacion;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
        private System.Windows.Forms.CheckBox chk_all;
    }
}