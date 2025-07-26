namespace GestionOperaciones.forms.medida
{
    partial class FrmKeeReporteExtraccion
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmKeeReporteExtraccion));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.lbl_registros = new System.Windows.Forms.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_sol_desde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_sol_hasta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cups22 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_desde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_hasta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.origen = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tipo_curva = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.motivo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.listBoxSufijo_CUPS22 = new System.Windows.Forms.ListBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btn_po1011 = new System.Windows.Forms.Button();
            this.listBoxExtraccion = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 179);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1158, 437);
            this.tabControl1.TabIndex = 64;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.cmdExcel);
            this.tabPage2.Controls.Add(this.lbl_registros);
            this.tabPage2.Controls.Add(this.dgv);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1150, 411);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Solicitudes de extracción";
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
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click_1);
            // 
            // lbl_registros
            // 
            this.lbl_registros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_registros.AutoSize = true;
            this.lbl_registros.Location = new System.Drawing.Point(1007, 14);
            this.lbl_registros.Name = "lbl_registros";
            this.lbl_registros.Size = new System.Drawing.Size(105, 13);
            this.lbl_registros.TabIndex = 61;
            this.lbl_registros.Text = "Registros:                 ";
            this.lbl_registros.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
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
            this.dgv.AutoGenerateColumns = false;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cups20,
            this.fecha_sol_desde,
            this.fecha_sol_hasta,
            this.cups22,
            this.fecha_desde,
            this.fecha_hasta,
            this.origen,
            this.tipo_curva,
            this.motivo});
            this.dgv.DataSource = this.listBoxSufijo_CUPS22.CustomTabOffsets;
            this.dgv.Location = new System.Drawing.Point(6, 40);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(1138, 365);
            this.dgv.TabIndex = 60;
            this.dgv.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgv_ColumnHeaderMouseClick);
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "cups20";
            this.cups20.HeaderText = "CUPS20";
            this.cups20.Name = "cups20";
            this.cups20.ReadOnly = true;
            this.cups20.Width = 73;
            // 
            // fecha_sol_desde
            // 
            this.fecha_sol_desde.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_sol_desde.DataPropertyName = "fecha_sol_desde";
            this.fecha_sol_desde.HeaderText = "FECHA SOL. DESDE";
            this.fecha_sol_desde.Name = "fecha_sol_desde";
            this.fecha_sol_desde.ReadOnly = true;
            this.fecha_sol_desde.Width = 123;
            // 
            // fecha_sol_hasta
            // 
            this.fecha_sol_hasta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_sol_hasta.DataPropertyName = "fecha_sol_hasta";
            this.fecha_sol_hasta.HeaderText = "FECHA SOL. HASTA";
            this.fecha_sol_hasta.Name = "fecha_sol_hasta";
            this.fecha_sol_hasta.ReadOnly = true;
            this.fecha_sol_hasta.Width = 122;
            // 
            // cups22
            // 
            this.cups22.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups22.DataPropertyName = "cups22";
            this.cups22.HeaderText = "CUPS22";
            this.cups22.Name = "cups22";
            this.cups22.ReadOnly = true;
            this.cups22.Width = 73;
            // 
            // fecha_desde
            // 
            this.fecha_desde.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_desde.DataPropertyName = "fecha_desde";
            this.fecha_desde.HeaderText = "FECHA DESDE";
            this.fecha_desde.Name = "fecha_desde";
            this.fecha_desde.ReadOnly = true;
            this.fecha_desde.Width = 98;
            // 
            // fecha_hasta
            // 
            this.fecha_hasta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_hasta.DataPropertyName = "fecha_hasta";
            this.fecha_hasta.HeaderText = "FECHA HASTA";
            this.fecha_hasta.Name = "fecha_hasta";
            this.fecha_hasta.ReadOnly = true;
            this.fecha_hasta.Width = 97;
            // 
            // origen
            // 
            this.origen.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.origen.DataPropertyName = "tipo";
            this.origen.HeaderText = "ORIGEN";
            this.origen.Name = "origen";
            this.origen.ReadOnly = true;
            this.origen.Width = 74;
            // 
            // tipo_curva
            // 
            this.tipo_curva.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tipo_curva.DataPropertyName = "fuente";
            this.tipo_curva.HeaderText = "TIPO CURVA";
            this.tipo_curva.Name = "tipo_curva";
            this.tipo_curva.ReadOnly = true;
            this.tipo_curva.Width = 89;
            // 
            // motivo
            // 
            this.motivo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.motivo.DataPropertyName = "extraccion";
            this.motivo.HeaderText = "MOTIVO";
            this.motivo.Name = "motivo";
            this.motivo.ReadOnly = true;
            this.motivo.Width = 74;
            // 
            // listBoxSufijo_CUPS22
            // 
            this.listBoxSufijo_CUPS22.FormattingEnabled = true;
            this.listBoxSufijo_CUPS22.Location = new System.Drawing.Point(152, 63);
            this.listBoxSufijo_CUPS22.Name = "listBoxSufijo_CUPS22";
            this.listBoxSufijo_CUPS22.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxSufijo_CUPS22.Size = new System.Drawing.Size(78, 95);
            this.listBoxSufijo_CUPS22.TabIndex = 68;
            this.listBoxSufijo_CUPS22.SelectedIndexChanged += new System.EventHandler(this.listBoxSufijo_CUPS22_SelectedIndexChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1182, 24);
            this.menuStrip1.TabIndex = 65;
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
            this.ayudaToolStripMenuItem.Click += new System.EventHandler(this.ayudaToolStripMenuItem_Click);
            // 
            // btn_po1011
            // 
            this.btn_po1011.Location = new System.Drawing.Point(293, 108);
            this.btn_po1011.Name = "btn_po1011";
            this.btn_po1011.Size = new System.Drawing.Size(168, 50);
            this.btn_po1011.TabIndex = 67;
            this.btn_po1011.Text = "    Exportación PO1011        (Red Eléctrica)";
            this.toolTip1.SetToolTip(this.btn_po1011, "\r\nExportar información a Excel");
            this.btn_po1011.UseVisualStyleBackColor = true;
            this.btn_po1011.Click += new System.EventHandler(this.btn_po1011_Click);
            // 
            // listBoxExtraccion
            // 
            this.listBoxExtraccion.FormattingEnabled = true;
            this.listBoxExtraccion.Location = new System.Drawing.Point(12, 63);
            this.listBoxExtraccion.Name = "listBoxExtraccion";
            this.listBoxExtraccion.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxExtraccion.Size = new System.Drawing.Size(134, 95);
            this.listBoxExtraccion.TabIndex = 66;
            this.listBoxExtraccion.SelectedIndexChanged += new System.EventHandler(this.listBoxExtraccion_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 69;
            this.label1.Text = "Motivo";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(149, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 70;
            this.label2.Text = "Sufijo CUPS";
            // 
            // FrmKeeReporteExtraccion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1182, 628);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBoxSufijo_CUPS22);
            this.Controls.Add(this.btn_po1011);
            this.Controls.Add(this.listBoxExtraccion);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmKeeReporteExtraccion";
            this.Text = "Consulta Kronos peticiones extracción";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmKeeReporteExtraccion_FormClosing);
            this.Load += new System.EventHandler(this.FrmKeeReporteExtraccion_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Label lbl_registros;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ListBox listBoxExtraccion;
        private System.Windows.Forms.Button btn_po1011;
        private System.Windows.Forms.ListBox listBoxSufijo_CUPS22;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_sol_desde;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_sol_hasta;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups22;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_desde;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_hasta;
        private System.Windows.Forms.DataGridViewTextBoxColumn origen;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipo_curva;
        private System.Windows.Forms.DataGridViewTextBoxColumn motivo;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
    }
}