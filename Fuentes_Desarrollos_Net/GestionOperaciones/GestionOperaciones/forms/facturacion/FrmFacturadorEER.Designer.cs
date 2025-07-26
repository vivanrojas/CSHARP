namespace GestionOperaciones.forms.facturacion
{
    partial class FrmFacturadorEER
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFacturadorEER));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cerrarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarCurvasDeAccessAMySQLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.informeDeFacturasAExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.parámetrosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.facturasManualesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chk_exportar_calculos = new System.Windows.Forms.CheckBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_fecha_consumo_hasta = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_fecha_consumo_desde = new System.Windows.Forms.DateTimePicker();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.lbl_registros_i = new System.Windows.Forms.Label();
            this.dgv_Inventario = new System.Windows.Forms.DataGridView();
            this.nif_i = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cliente_i = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cpe_i = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tarifa = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_consumo_desde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_consumo_hasta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.estado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_factura = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_factura = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.facturarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.generarPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Inventario)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
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
            this.menuStrip1.Size = new System.Drawing.Size(1082, 24);
            this.menuStrip1.TabIndex = 0;
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
            this.cerrarToolStripMenuItem.Name = "cerrarToolStripMenuItem";
            this.cerrarToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.cerrarToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.cerrarToolStripMenuItem.Text = "Salir";
            this.cerrarToolStripMenuItem.Click += new System.EventHandler(this.cerrarToolStripMenuItem_Click);
            // 
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarCurvasDeAccessAMySQLToolStripMenuItem,
            this.toolStripSeparator2,
            this.informeDeFacturasAExcelToolStripMenuItem,
            this.toolStripSeparator3,
            this.parámetrosToolStripMenuItem,
            this.toolStripSeparator4,
            this.facturasManualesToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarCurvasDeAccessAMySQLToolStripMenuItem
            // 
            this.importarCurvasDeAccessAMySQLToolStripMenuItem.Name = "importarCurvasDeAccessAMySQLToolStripMenuItem";
            this.importarCurvasDeAccessAMySQLToolStripMenuItem.Size = new System.Drawing.Size(264, 22);
            this.importarCurvasDeAccessAMySQLToolStripMenuItem.Text = "Importar Curvas de Access a MySQL";
            this.importarCurvasDeAccessAMySQLToolStripMenuItem.Click += new System.EventHandler(this.ImportarCurvasDeAccessAMySQLToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(261, 6);
            // 
            // informeDeFacturasAExcelToolStripMenuItem
            // 
            this.informeDeFacturasAExcelToolStripMenuItem.Name = "informeDeFacturasAExcelToolStripMenuItem";
            this.informeDeFacturasAExcelToolStripMenuItem.Size = new System.Drawing.Size(264, 22);
            this.informeDeFacturasAExcelToolStripMenuItem.Text = "Informe de Facturas a Excel";
            this.informeDeFacturasAExcelToolStripMenuItem.Click += new System.EventHandler(this.informeDeFacturasAExcelToolStripMenuItem_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(261, 6);
            // 
            // parámetrosToolStripMenuItem
            // 
            this.parámetrosToolStripMenuItem.Name = "parámetrosToolStripMenuItem";
            this.parámetrosToolStripMenuItem.Size = new System.Drawing.Size(264, 22);
            this.parámetrosToolStripMenuItem.Text = "Parámetros";
            this.parámetrosToolStripMenuItem.Click += new System.EventHandler(this.parámetrosToolStripMenuItem_Click);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(261, 6);
            // 
            // facturasManualesToolStripMenuItem
            // 
            this.facturasManualesToolStripMenuItem.Name = "facturasManualesToolStripMenuItem";
            this.facturasManualesToolStripMenuItem.Size = new System.Drawing.Size(264, 22);
            this.facturasManualesToolStripMenuItem.Text = "Facturas Manuales";
            this.facturasManualesToolStripMenuItem.Click += new System.EventHandler(this.facturasManualesToolStripMenuItem_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chk_exportar_calculos);
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_fecha_consumo_hasta);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txt_fecha_consumo_desde);
            this.groupBox1.Location = new System.Drawing.Point(12, 38);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(567, 137);
            this.groupBox1.TabIndex = 23;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // chk_exportar_calculos
            // 
            this.chk_exportar_calculos.AutoSize = true;
            this.chk_exportar_calculos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chk_exportar_calculos.Location = new System.Drawing.Point(10, 87);
            this.chk_exportar_calculos.Name = "chk_exportar_calculos";
            this.chk_exportar_calculos.Size = new System.Drawing.Size(124, 17);
            this.chk_exportar_calculos.TabIndex = 28;
            this.chk_exportar_calculos.Text = "Exportar cálculos";
            this.chk_exportar_calculos.UseVisualStyleBackColor = true;
            // 
            // btnRefresh
            // 
            this.btnRefresh.Image = global::GestionOperaciones.Properties.Resources.if_view_refresh_118801;
            this.btnRefresh.Location = new System.Drawing.Point(465, 33);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 52);
            this.btnRefresh.TabIndex = 27;
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.BtnRefresh_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(239, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(134, 13);
            this.label3.TabIndex = 24;
            this.label3.Text = "Hasta Fecha Consumo";
            // 
            // txt_fecha_consumo_hasta
            // 
            this.txt_fecha_consumo_hasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_consumo_hasta.Location = new System.Drawing.Point(379, 47);
            this.txt_fecha_consumo_hasta.Name = "txt_fecha_consumo_hasta";
            this.txt_fecha_consumo_hasta.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_consumo_hasta.TabIndex = 23;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(7, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(137, 13);
            this.label4.TabIndex = 22;
            this.label4.Text = "Desde Fecha Consumo";
            // 
            // txt_fecha_consumo_desde
            // 
            this.txt_fecha_consumo_desde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_consumo_desde.Location = new System.Drawing.Point(150, 47);
            this.txt_fecha_consumo_desde.Name = "txt_fecha_consumo_desde";
            this.txt_fecha_consumo_desde.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_consumo_desde.TabIndex = 21;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 181);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1058, 424);
            this.tabControl1.TabIndex = 29;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.cmdExcel);
            this.tabPage1.Controls.Add(this.lbl_registros_i);
            this.tabPage1.Controls.Add(this.dgv_Inventario);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1050, 398);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Inventario";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(6, 8);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 63;
            this.cmdExcel.UseVisualStyleBackColor = true;
            // 
            // lbl_registros_i
            // 
            this.lbl_registros_i.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_registros_i.AutoSize = true;
            this.lbl_registros_i.Location = new System.Drawing.Point(936, 13);
            this.lbl_registros_i.Name = "lbl_registros_i";
            this.lbl_registros_i.Size = new System.Drawing.Size(105, 13);
            this.lbl_registros_i.TabIndex = 62;
            this.lbl_registros_i.Text = "Registros:                 ";
            this.lbl_registros_i.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dgv_Inventario
            // 
            this.dgv_Inventario.AllowUserToAddRows = false;
            this.dgv_Inventario.AllowUserToDeleteRows = false;
            this.dgv_Inventario.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv_Inventario.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Inventario.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.nif_i,
            this.cliente_i,
            this.cpe_i,
            this.tarifa,
            this.fecha_consumo_desde,
            this.fecha_consumo_hasta,
            this.estado,
            this.num_factura,
            this.fecha_factura});
            this.dgv_Inventario.ContextMenuStrip = this.contextMenuStrip1;
            this.dgv_Inventario.Location = new System.Drawing.Point(3, 42);
            this.dgv_Inventario.Name = "dgv_Inventario";
            this.dgv_Inventario.ReadOnly = true;
            this.dgv_Inventario.Size = new System.Drawing.Size(1038, 350);
            this.dgv_Inventario.TabIndex = 1;
            this.dgv_Inventario.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_Inventario_CellContentClick);
            this.dgv_Inventario.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_Inventario_CellDoubleClick);
            // 
            // nif_i
            // 
            this.nif_i.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.nif_i.DataPropertyName = "nif";
            this.nif_i.HeaderText = "NIF";
            this.nif_i.Name = "nif_i";
            this.nif_i.ReadOnly = true;
            this.nif_i.Width = 49;
            // 
            // cliente_i
            // 
            this.cliente_i.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cliente_i.DataPropertyName = "cliente";
            this.cliente_i.HeaderText = "Cliente";
            this.cliente_i.Name = "cliente_i";
            this.cliente_i.ReadOnly = true;
            this.cliente_i.Width = 64;
            // 
            // cpe_i
            // 
            this.cpe_i.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cpe_i.DataPropertyName = "cups20";
            this.cpe_i.HeaderText = "CUPS20";
            this.cpe_i.Name = "cpe_i";
            this.cpe_i.ReadOnly = true;
            this.cpe_i.Width = 73;
            // 
            // tarifa
            // 
            this.tarifa.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tarifa.DataPropertyName = "tarifa";
            this.tarifa.HeaderText = "TARIFA";
            this.tarifa.Name = "tarifa";
            this.tarifa.ReadOnly = true;
            this.tarifa.Width = 70;
            // 
            // fecha_consumo_desde
            // 
            this.fecha_consumo_desde.DataPropertyName = "fecha_consumo_desde";
            this.fecha_consumo_desde.HeaderText = "Fecha consumo Desde";
            this.fecha_consumo_desde.Name = "fecha_consumo_desde";
            this.fecha_consumo_desde.ReadOnly = true;
            // 
            // fecha_consumo_hasta
            // 
            this.fecha_consumo_hasta.DataPropertyName = "fecha_consumo_hasta";
            this.fecha_consumo_hasta.HeaderText = "Fecha Consumo Hasta";
            this.fecha_consumo_hasta.Name = "fecha_consumo_hasta";
            this.fecha_consumo_hasta.ReadOnly = true;
            // 
            // estado
            // 
            this.estado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.estado.DataPropertyName = "estado";
            this.estado.HeaderText = "Estado";
            this.estado.Name = "estado";
            this.estado.ReadOnly = true;
            this.estado.Width = 65;
            // 
            // num_factura
            // 
            this.num_factura.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.num_factura.DataPropertyName = "num_factura";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.num_factura.DefaultCellStyle = dataGridViewCellStyle1;
            this.num_factura.HeaderText = "Código Factura";
            this.num_factura.Name = "num_factura";
            this.num_factura.ReadOnly = true;
            this.num_factura.Width = 96;
            // 
            // fecha_factura
            // 
            this.fecha_factura.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_factura.DataPropertyName = "fecha_factura";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fecha_factura.DefaultCellStyle = dataGridViewCellStyle2;
            this.fecha_factura.HeaderText = "Fecha Factura";
            this.fecha_factura.Name = "fecha_factura";
            this.fecha_factura.ReadOnly = true;
            this.fecha_factura.Width = 93;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.facturarToolStripMenuItem,
            this.toolStripSeparator1,
            this.generarPDFToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(173, 54);
            // 
            // facturarToolStripMenuItem
            // 
            this.facturarToolStripMenuItem.Name = "facturarToolStripMenuItem";
            this.facturarToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.facturarToolStripMenuItem.Text = "Generar Prefactura";
            this.facturarToolStripMenuItem.Click += new System.EventHandler(this.FacturarToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(169, 6);
            // 
            // generarPDFToolStripMenuItem
            // 
            this.generarPDFToolStripMenuItem.Enabled = false;
            this.generarPDFToolStripMenuItem.Name = "generarPDFToolStripMenuItem";
            this.generarPDFToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.generarPDFToolStripMenuItem.Text = "Generar PDF";
            this.generarPDFToolStripMenuItem.Click += new System.EventHandler(this.generarPDFToolStripMenuItem_Click);
            // 
            // FrmFacturadorEER
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1082, 617);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmFacturadorEER";
            this.Text = "Facturador";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmFacturadorEER_FormClosing);
            this.Load += new System.EventHandler(this.Facturador_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Inventario)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker txt_fecha_consumo_hasta;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker txt_fecha_consumo_desde;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Label lbl_registros_i;
        private System.Windows.Forms.DataGridView dgv_Inventario;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem facturarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarCurvasDeAccessAMySQLToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem generarPDFToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem parámetrosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem informeDeFacturasAExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripMenuItem cerrarToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn nif_i;
        private System.Windows.Forms.DataGridViewTextBoxColumn cliente_i;
        private System.Windows.Forms.DataGridViewTextBoxColumn cpe_i;
        private System.Windows.Forms.DataGridViewTextBoxColumn tarifa;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_consumo_desde;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_consumo_hasta;
        private System.Windows.Forms.DataGridViewTextBoxColumn estado;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_factura;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_factura;
        private System.Windows.Forms.CheckBox chk_exportar_calculos;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripMenuItem facturasManualesToolStripMenuItem;
    }
}