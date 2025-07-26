namespace GestionOperaciones.forms.facturacion
{
    partial class FrmFacturadorPortugal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFacturadorPortugal));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtNif = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_fecha_consumo_hasta = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_fecha_consumo_desde = new System.Windows.Forms.DateTimePicker();
            this.btnFacturar = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarPreciosSpotToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.importarClicksToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.lbl_registros_i = new System.Windows.Forms.Label();
            this.dgv_Inventario = new System.Windows.Forms.DataGridView();
            this.actualizado = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.nif_i = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cliente_i = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cpe_i = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ltp = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.error = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.estado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.generarFacturaATravésDePlantillaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generarPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbl_precio_medio = new System.Windows.Forms.Label();
            this.lbl_registros_s = new System.Windows.Forms.Label();
            this.dgv_Spot = new System.Windows.Forms.DataGridView();
            this.fecha_s = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.precio_s = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgv_Clicks = new System.Windows.Forms.DataGridView();
            this.nif = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cpe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.click = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mercado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.operacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.producto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_desde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_hasta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bl = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fee = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.volumen = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Inventario)).BeginInit();
            this.contextMenuStrip.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Spot)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Clicks)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtNif);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_fecha_consumo_hasta);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txt_fecha_consumo_desde);
            this.groupBox1.Location = new System.Drawing.Point(22, 44);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(567, 137);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Image = global::GestionOperaciones.Properties.Resources.if_view_refresh_118801;
            this.btnRefresh.Location = new System.Drawing.Point(465, 33);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 52);
            this.btnRefresh.TabIndex = 27;
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(117, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(27, 13);
            this.label1.TabIndex = 26;
            this.label1.Text = "NIF";
            // 
            // txtNif
            // 
            this.txtNif.Location = new System.Drawing.Point(150, 75);
            this.txtNif.Name = "txtNif";
            this.txtNif.Size = new System.Drawing.Size(137, 20);
            this.txtNif.TabIndex = 25;
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
            // btnFacturar
            // 
            this.btnFacturar.Location = new System.Drawing.Point(607, 83);
            this.btnFacturar.Name = "btnFacturar";
            this.btnFacturar.Size = new System.Drawing.Size(118, 56);
            this.btnFacturar.TabIndex = 27;
            this.btnFacturar.Text = "Facturar DIA";
            this.btnFacturar.UseVisualStyleBackColor = true;
            this.btnFacturar.Click += new System.EventHandler(this.btnFacturar_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1246, 24);
            this.menuStrip1.TabIndex = 23;
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
            this.salirToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_dialog_close_29299;
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarPreciosSpotToolStripMenuItem,
            this.toolStripSeparator1,
            this.importarClicksToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarPreciosSpotToolStripMenuItem
            // 
            this.importarPreciosSpotToolStripMenuItem.Name = "importarPreciosSpotToolStripMenuItem";
            this.importarPreciosSpotToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.importarPreciosSpotToolStripMenuItem.Text = "Importar Precios Spot";
            this.importarPreciosSpotToolStripMenuItem.Click += new System.EventHandler(this.importarPreciosSpotToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(185, 6);
            // 
            // importarClicksToolStripMenuItem
            // 
            this.importarClicksToolStripMenuItem.Name = "importarClicksToolStripMenuItem";
            this.importarClicksToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.importarClicksToolStripMenuItem.Text = "Importar Clicks";
            this.importarClicksToolStripMenuItem.Click += new System.EventHandler(this.importarClicksToolStripMenuItem_Click);
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
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(23, 187);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1211, 442);
            this.tabControl1.TabIndex = 28;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.cmdExcel);
            this.tabPage1.Controls.Add(this.lbl_registros_i);
            this.tabPage1.Controls.Add(this.dgv_Inventario);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1203, 416);
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
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // lbl_registros_i
            // 
            this.lbl_registros_i.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_registros_i.AutoSize = true;
            this.lbl_registros_i.Location = new System.Drawing.Point(1089, 13);
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
            this.actualizado,
            this.nif_i,
            this.cliente_i,
            this.cpe_i,
            this.ltp,
            this.error,
            this.estado});
            this.dgv_Inventario.ContextMenuStrip = this.contextMenuStrip;
            this.dgv_Inventario.Location = new System.Drawing.Point(3, 42);
            this.dgv_Inventario.Name = "dgv_Inventario";
            this.dgv_Inventario.ReadOnly = true;
            this.dgv_Inventario.Size = new System.Drawing.Size(1191, 368);
            this.dgv_Inventario.TabIndex = 1;
            // 
            // actualizado
            // 
            this.actualizado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.actualizado.DataPropertyName = "actualizado";
            this.actualizado.HeaderText = "Actualizado";
            this.actualizado.Name = "actualizado";
            this.actualizado.ReadOnly = true;
            this.actualizado.Width = 68;
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
            this.cpe_i.DataPropertyName = "cpe";
            this.cpe_i.HeaderText = "CPE";
            this.cpe_i.Name = "cpe_i";
            this.cpe_i.ReadOnly = true;
            this.cpe_i.Width = 53;
            // 
            // ltp
            // 
            this.ltp.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ltp.DataPropertyName = "ltp";
            this.ltp.HeaderText = "LTP";
            this.ltp.Name = "ltp";
            this.ltp.ReadOnly = true;
            this.ltp.Width = 52;
            // 
            // error
            // 
            this.error.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.error.DataPropertyName = "error";
            this.error.HeaderText = "Aviso";
            this.error.Name = "error";
            this.error.ReadOnly = true;
            this.error.Width = 58;
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
            // contextMenuStrip
            // 
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.generarFacturaATravésDePlantillaToolStripMenuItem,
            this.generarPDFToolStripMenuItem});
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(260, 48);
            // 
            // generarFacturaATravésDePlantillaToolStripMenuItem
            // 
            this.generarFacturaATravésDePlantillaToolStripMenuItem.Name = "generarFacturaATravésDePlantillaToolStripMenuItem";
            this.generarFacturaATravésDePlantillaToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.generarFacturaATravésDePlantillaToolStripMenuItem.Text = "Generar factura a través de plantilla";
            this.generarFacturaATravésDePlantillaToolStripMenuItem.Click += new System.EventHandler(this.generarFacturaATravésDePlantillaToolStripMenuItem_Click);
            // 
            // generarPDFToolStripMenuItem
            // 
            this.generarPDFToolStripMenuItem.Name = "generarPDFToolStripMenuItem";
            this.generarPDFToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.generarPDFToolStripMenuItem.Text = "Generar PDF";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lbl_precio_medio);
            this.tabPage2.Controls.Add(this.lbl_registros_s);
            this.tabPage2.Controls.Add(this.dgv_Spot);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1203, 416);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Precios Spot";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lbl_precio_medio
            // 
            this.lbl_precio_medio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_precio_medio.AutoSize = true;
            this.lbl_precio_medio.Location = new System.Drawing.Point(494, 15);
            this.lbl_precio_medio.Name = "lbl_precio_medio";
            this.lbl_precio_medio.Size = new System.Drawing.Size(122, 13);
            this.lbl_precio_medio.TabIndex = 64;
            this.lbl_precio_medio.Text = "Precio medio:                 ";
            this.lbl_precio_medio.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lbl_registros_s
            // 
            this.lbl_registros_s.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_registros_s.AutoSize = true;
            this.lbl_registros_s.Location = new System.Drawing.Point(646, 15);
            this.lbl_registros_s.Name = "lbl_registros_s";
            this.lbl_registros_s.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_registros_s.Size = new System.Drawing.Size(105, 13);
            this.lbl_registros_s.TabIndex = 63;
            this.lbl_registros_s.Text = "Registros:                 ";
            this.lbl_registros_s.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dgv_Spot
            // 
            this.dgv_Spot.AllowUserToAddRows = false;
            this.dgv_Spot.AllowUserToDeleteRows = false;
            this.dgv_Spot.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv_Spot.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Spot.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.fecha_s,
            this.precio_s});
            this.dgv_Spot.Location = new System.Drawing.Point(6, 40);
            this.dgv_Spot.Name = "dgv_Spot";
            this.dgv_Spot.ReadOnly = true;
            this.dgv_Spot.Size = new System.Drawing.Size(745, 179);
            this.dgv_Spot.TabIndex = 0;
            // 
            // fecha_s
            // 
            this.fecha_s.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_s.DataPropertyName = "fecha_hora";
            this.fecha_s.HeaderText = "Fecha";
            this.fecha_s.Name = "fecha_s";
            this.fecha_s.ReadOnly = true;
            this.fecha_s.Width = 62;
            // 
            // precio_s
            // 
            this.precio_s.DataPropertyName = "precio";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N2";
            dataGridViewCellStyle1.NullValue = null;
            this.precio_s.DefaultCellStyle = dataGridViewCellStyle1;
            this.precio_s.HeaderText = "Precio € MW";
            this.precio_s.Name = "precio_s";
            this.precio_s.ReadOnly = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgv_Clicks);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1203, 416);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Clicks";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgv_Clicks
            // 
            this.dgv_Clicks.AllowUserToAddRows = false;
            this.dgv_Clicks.AllowUserToDeleteRows = false;
            this.dgv_Clicks.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv_Clicks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Clicks.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.nif,
            this.cliente,
            this.cpe,
            this.click,
            this.fecha,
            this.mercado,
            this.operacion,
            this.producto,
            this.fecha_desde,
            this.fecha_hasta,
            this.bl,
            this.fee,
            this.volumen});
            this.dgv_Clicks.Location = new System.Drawing.Point(5, 76);
            this.dgv_Clicks.Name = "dgv_Clicks";
            this.dgv_Clicks.ReadOnly = true;
            this.dgv_Clicks.Size = new System.Drawing.Size(745, 146);
            this.dgv_Clicks.TabIndex = 2;
            // 
            // nif
            // 
            this.nif.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.nif.DataPropertyName = "nif";
            this.nif.HeaderText = "nif";
            this.nif.Name = "nif";
            this.nif.ReadOnly = true;
            this.nif.Width = 43;
            // 
            // cliente
            // 
            this.cliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cliente.DataPropertyName = "cliente";
            this.cliente.HeaderText = "cliente";
            this.cliente.Name = "cliente";
            this.cliente.ReadOnly = true;
            this.cliente.Width = 63;
            // 
            // cpe
            // 
            this.cpe.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cpe.DataPropertyName = "cpe";
            this.cpe.HeaderText = "cpe";
            this.cpe.Name = "cpe";
            this.cpe.ReadOnly = true;
            this.cpe.Width = 50;
            // 
            // click
            // 
            this.click.DataPropertyName = "click";
            this.click.HeaderText = "Click";
            this.click.Name = "click";
            this.click.ReadOnly = true;
            // 
            // fecha
            // 
            this.fecha.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha.DataPropertyName = "fecha_operacion";
            this.fecha.HeaderText = "Fecha";
            this.fecha.Name = "fecha";
            this.fecha.ReadOnly = true;
            this.fecha.Width = 62;
            // 
            // mercado
            // 
            this.mercado.DataPropertyName = "mercado";
            this.mercado.HeaderText = "Mercado";
            this.mercado.Name = "mercado";
            this.mercado.ReadOnly = true;
            // 
            // operacion
            // 
            this.operacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.operacion.DataPropertyName = "operacion";
            this.operacion.HeaderText = "Operación";
            this.operacion.Name = "operacion";
            this.operacion.ReadOnly = true;
            this.operacion.Width = 81;
            // 
            // producto
            // 
            this.producto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.producto.HeaderText = "Producto";
            this.producto.Name = "producto";
            this.producto.ReadOnly = true;
            this.producto.Width = 75;
            // 
            // fecha_desde
            // 
            this.fecha_desde.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_desde.HeaderText = "Fecha Desde";
            this.fecha_desde.Name = "fecha_desde";
            this.fecha_desde.ReadOnly = true;
            this.fecha_desde.Width = 96;
            // 
            // fecha_hasta
            // 
            this.fecha_hasta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_hasta.HeaderText = "Fecha Hasta";
            this.fecha_hasta.Name = "fecha_hasta";
            this.fecha_hasta.ReadOnly = true;
            this.fecha_hasta.Width = 93;
            // 
            // bl
            // 
            this.bl.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.bl.DataPropertyName = "bl";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N2";
            dataGridViewCellStyle2.NullValue = null;
            this.bl.DefaultCellStyle = dataGridViewCellStyle2;
            this.bl.HeaderText = "BL";
            this.bl.Name = "bl";
            this.bl.ReadOnly = true;
            this.bl.Width = 45;
            // 
            // fee
            // 
            this.fee.DataPropertyName = "fee";
            dataGridViewCellStyle3.Format = "N2";
            dataGridViewCellStyle3.NullValue = null;
            this.fee.DefaultCellStyle = dataGridViewCellStyle3;
            this.fee.HeaderText = "Fee";
            this.fee.Name = "fee";
            this.fee.ReadOnly = true;
            // 
            // volumen
            // 
            this.volumen.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.volumen.DataPropertyName = "volumen";
            dataGridViewCellStyle4.Format = "N2";
            dataGridViewCellStyle4.NullValue = null;
            this.volumen.DefaultCellStyle = dataGridViewCellStyle4;
            this.volumen.HeaderText = "Volumen";
            this.volumen.Name = "volumen";
            this.volumen.ReadOnly = true;
            this.volumen.Width = 73;
            // 
            // FrmFacturadorPortugal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1246, 641);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnFacturar);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmFacturadorPortugal";
            this.Text = "Facturador productos sofisticados de Portugal";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmFacturadorPortugal_FormClosing);
            this.Load += new System.EventHandler(this.FrmFacturadorPortugal_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Inventario)).EndInit();
            this.contextMenuStrip.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Spot)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Clicks)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnFacturar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtNif;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarPreciosSpotToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem importarClicksToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dgv_Spot;
        private System.Windows.Forms.DataGridView dgv_Inventario;
        private System.Windows.Forms.DataGridView dgv_Clicks;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.DataGridViewTextBoxColumn nif;
        private System.Windows.Forms.DataGridViewTextBoxColumn cliente;
        private System.Windows.Forms.DataGridViewTextBoxColumn cpe;
        private System.Windows.Forms.DataGridViewTextBoxColumn click;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha;
        private System.Windows.Forms.DataGridViewTextBoxColumn mercado;
        private System.Windows.Forms.DataGridViewTextBoxColumn operacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn producto;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_desde;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_hasta;
        private System.Windows.Forms.DataGridViewTextBoxColumn bl;
        private System.Windows.Forms.DataGridViewTextBoxColumn fee;
        private System.Windows.Forms.DataGridViewTextBoxColumn volumen;
        private System.Windows.Forms.Label lbl_registros_i;
        private System.Windows.Forms.Label lbl_registros_s;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_s;
        private System.Windows.Forms.DataGridViewTextBoxColumn precio_s;
        private System.Windows.Forms.Label lbl_precio_medio;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem generarFacturaATravésDePlantillaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem generarPDFToolStripMenuItem;
        private System.Windows.Forms.DataGridViewCheckBoxColumn actualizado;
        private System.Windows.Forms.DataGridViewTextBoxColumn nif_i;
        private System.Windows.Forms.DataGridViewTextBoxColumn cliente_i;
        private System.Windows.Forms.DataGridViewTextBoxColumn cpe_i;
        private System.Windows.Forms.DataGridViewTextBoxColumn ltp;
        private System.Windows.Forms.DataGridViewTextBoxColumn error;
        private System.Windows.Forms.DataGridViewTextBoxColumn estado;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker txt_fecha_consumo_hasta;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker txt_fecha_consumo_desde;
    }
}