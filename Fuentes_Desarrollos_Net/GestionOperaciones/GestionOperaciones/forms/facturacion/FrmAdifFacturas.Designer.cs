namespace GestionOperaciones.forms.facturacion
{
    partial class FrmAdifFacturas
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAdifFacturas));
            this.dgvCups = new System.Windows.Forms.DataGridView();
            this.CUPSREE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LOTE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.medida_en_baja = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.devolucion_de_energia = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.cierres_energia = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.existe_factura_adif = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.existe_factura_sce = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.CREFEREN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SECFACTU = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FFACTURA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FFACTDES = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FFACTHAS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TFACTURA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TESTFACT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CONSUMO_ADIF = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CONSUMO_SCE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DIF_CONSUMO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TOTAL_ADIF = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TOTAL_SCE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DIF_TOTAL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuCopyPaste = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuCopy = new System.Windows.Forms.ToolStripMenuItem();
            this.btnExcel = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.archivosFacturasADIFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.archivosFacturasAdif_REG_toolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.cierresDeEnergíaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkDiff = new System.Windows.Forms.CheckBox();
            this.txtlote = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnBuscar = new System.Windows.Forms.Button();
            this.txtcupsree = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtFFACTHAS = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.dtFFACTDES = new System.Windows.Forms.DateTimePicker();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.parámetrosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCups)).BeginInit();
            this.menuCopyPaste.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvCups
            // 
            this.dgvCups.AllowUserToAddRows = false;
            this.dgvCups.AllowUserToDeleteRows = false;
            this.dgvCups.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.SandyBrown;
            this.dgvCups.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvCups.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvCups.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvCups.ColumnHeadersHeight = 29;
            this.dgvCups.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CUPSREE,
            this.LOTE,
            this.medida_en_baja,
            this.devolucion_de_energia,
            this.cierres_energia,
            this.existe_factura_adif,
            this.existe_factura_sce,
            this.CREFEREN,
            this.SECFACTU,
            this.FFACTURA,
            this.FFACTDES,
            this.FFACTHAS,
            this.TFACTURA,
            this.TESTFACT,
            this.CONSUMO_ADIF,
            this.CONSUMO_SCE,
            this.DIF_CONSUMO,
            this.TOTAL_ADIF,
            this.TOTAL_SCE,
            this.DIF_TOTAL});
            this.dgvCups.Location = new System.Drawing.Point(12, 181);
            this.dgvCups.Name = "dgvCups";
            this.dgvCups.ReadOnly = true;
            this.dgvCups.RowHeadersWidth = 51;
            this.dgvCups.Size = new System.Drawing.Size(1330, 522);
            this.dgvCups.TabIndex = 25;
            this.dgvCups.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvCups_ColumnHeaderMouseClick);
            // 
            // CUPSREE
            // 
            this.CUPSREE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.CUPSREE.DataPropertyName = "CUPSREE";
            this.CUPSREE.HeaderText = "CUPS20";
            this.CUPSREE.MinimumWidth = 6;
            this.CUPSREE.Name = "CUPSREE";
            this.CUPSREE.ReadOnly = true;
            this.CUPSREE.Width = 73;
            // 
            // LOTE
            // 
            this.LOTE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.LOTE.DataPropertyName = "LOTE";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            this.LOTE.DefaultCellStyle = dataGridViewCellStyle2;
            this.LOTE.HeaderText = "LOTE";
            this.LOTE.MinimumWidth = 6;
            this.LOTE.Name = "LOTE";
            this.LOTE.ReadOnly = true;
            this.LOTE.Width = 60;
            // 
            // medida_en_baja
            // 
            this.medida_en_baja.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.medida_en_baja.DataPropertyName = "medida_en_baja";
            this.medida_en_baja.HeaderText = "Medida en baja";
            this.medida_en_baja.MinimumWidth = 6;
            this.medida_en_baja.Name = "medida_en_baja";
            this.medida_en_baja.ReadOnly = true;
            this.medida_en_baja.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.medida_en_baja.Width = 105;
            // 
            // devolucion_de_energia
            // 
            this.devolucion_de_energia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.devolucion_de_energia.DataPropertyName = "devolucion_de_energia";
            this.devolucion_de_energia.HeaderText = "Devolución de energía";
            this.devolucion_de_energia.MinimumWidth = 6;
            this.devolucion_de_energia.Name = "devolucion_de_energia";
            this.devolucion_de_energia.ReadOnly = true;
            this.devolucion_de_energia.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.devolucion_de_energia.Width = 141;
            // 
            // cierres_energia
            // 
            this.cierres_energia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cierres_energia.DataPropertyName = "cierres_energia";
            this.cierres_energia.HeaderText = "Cierres energía";
            this.cierres_energia.MinimumWidth = 6;
            this.cierres_energia.Name = "cierres_energia";
            this.cierres_energia.ReadOnly = true;
            this.cierres_energia.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.cierres_energia.Width = 104;
            // 
            // existe_factura_adif
            // 
            this.existe_factura_adif.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.existe_factura_adif.DataPropertyName = "existe_factura_adif";
            this.existe_factura_adif.HeaderText = "Existe factura ADIF";
            this.existe_factura_adif.MinimumWidth = 6;
            this.existe_factura_adif.Name = "existe_factura_adif";
            this.existe_factura_adif.ReadOnly = true;
            this.existe_factura_adif.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.existe_factura_adif.Width = 123;
            // 
            // existe_factura_sce
            // 
            this.existe_factura_sce.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.existe_factura_sce.DataPropertyName = "existe_factura_sce";
            this.existe_factura_sce.HeaderText = "Existe factura ENDESA";
            this.existe_factura_sce.MinimumWidth = 6;
            this.existe_factura_sce.Name = "existe_factura_sce";
            this.existe_factura_sce.ReadOnly = true;
            this.existe_factura_sce.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.existe_factura_sce.Width = 143;
            // 
            // CREFEREN
            // 
            this.CREFEREN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.CREFEREN.DataPropertyName = "CREFEREN";
            this.CREFEREN.HeaderText = "CREFEREN";
            this.CREFEREN.MinimumWidth = 6;
            this.CREFEREN.Name = "CREFEREN";
            this.CREFEREN.ReadOnly = true;
            this.CREFEREN.Width = 90;
            // 
            // SECFACTU
            // 
            this.SECFACTU.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.SECFACTU.DataPropertyName = "SECFACTU";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.SECFACTU.DefaultCellStyle = dataGridViewCellStyle3;
            this.SECFACTU.HeaderText = "SECFACTU";
            this.SECFACTU.MinimumWidth = 6;
            this.SECFACTU.Name = "SECFACTU";
            this.SECFACTU.ReadOnly = true;
            this.SECFACTU.Width = 88;
            // 
            // FFACTURA
            // 
            this.FFACTURA.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FFACTURA.DataPropertyName = "FFACTURA";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FFACTURA.DefaultCellStyle = dataGridViewCellStyle4;
            this.FFACTURA.HeaderText = "F. FACTURA";
            this.FFACTURA.MinimumWidth = 6;
            this.FFACTURA.Name = "FFACTURA";
            this.FFACTURA.ReadOnly = true;
            this.FFACTURA.Width = 94;
            // 
            // FFACTDES
            // 
            this.FFACTDES.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FFACTDES.DataPropertyName = "FFACTDES";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FFACTDES.DefaultCellStyle = dataGridViewCellStyle5;
            this.FFACTDES.HeaderText = "F. CONSUMO DESDE";
            this.FFACTDES.MinimumWidth = 6;
            this.FFACTDES.Name = "FFACTDES";
            this.FFACTDES.ReadOnly = true;
            this.FFACTDES.Width = 139;
            // 
            // FFACTHAS
            // 
            this.FFACTHAS.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FFACTHAS.DataPropertyName = "FFACTHAS";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FFACTHAS.DefaultCellStyle = dataGridViewCellStyle6;
            this.FFACTHAS.HeaderText = "F. CONSUMO HASTA";
            this.FFACTHAS.MinimumWidth = 6;
            this.FFACTHAS.Name = "FFACTHAS";
            this.FFACTHAS.ReadOnly = true;
            this.FFACTHAS.Width = 138;
            // 
            // TFACTURA
            // 
            this.TFACTURA.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.TFACTURA.DataPropertyName = "de_tfactura";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.TFACTURA.DefaultCellStyle = dataGridViewCellStyle7;
            this.TFACTURA.HeaderText = "T. FACTURA";
            this.TFACTURA.MinimumWidth = 6;
            this.TFACTURA.Name = "TFACTURA";
            this.TFACTURA.ReadOnly = true;
            this.TFACTURA.Width = 95;
            // 
            // TESTFACT
            // 
            this.TESTFACT.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.TESTFACT.DataPropertyName = "testfact";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.TESTFACT.DefaultCellStyle = dataGridViewCellStyle8;
            this.TESTFACT.HeaderText = "TESTFACT";
            this.TESTFACT.MinimumWidth = 6;
            this.TESTFACT.Name = "TESTFACT";
            this.TESTFACT.ReadOnly = true;
            this.TESTFACT.Width = 87;
            // 
            // CONSUMO_ADIF
            // 
            this.CONSUMO_ADIF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.CONSUMO_ADIF.DataPropertyName = "CONSUMO_ADIF";
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle9.Format = "N0";
            dataGridViewCellStyle9.NullValue = null;
            this.CONSUMO_ADIF.DefaultCellStyle = dataGridViewCellStyle9;
            this.CONSUMO_ADIF.HeaderText = "CONSUMO ADIF";
            this.CONSUMO_ADIF.MinimumWidth = 6;
            this.CONSUMO_ADIF.Name = "CONSUMO_ADIF";
            this.CONSUMO_ADIF.ReadOnly = true;
            this.CONSUMO_ADIF.Width = 114;
            // 
            // CONSUMO_SCE
            // 
            this.CONSUMO_SCE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.CONSUMO_SCE.DataPropertyName = "CONSUMO_SCE";
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle10.Format = "N0";
            dataGridViewCellStyle10.NullValue = null;
            this.CONSUMO_SCE.DefaultCellStyle = dataGridViewCellStyle10;
            this.CONSUMO_SCE.HeaderText = "CONSUMO ENDESA";
            this.CONSUMO_SCE.MinimumWidth = 6;
            this.CONSUMO_SCE.Name = "CONSUMO_SCE";
            this.CONSUMO_SCE.ReadOnly = true;
            this.CONSUMO_SCE.Width = 134;
            // 
            // DIF_CONSUMO
            // 
            this.DIF_CONSUMO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.DIF_CONSUMO.DataPropertyName = "DIF_CONSUMO";
            dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle11.Format = "N0";
            dataGridViewCellStyle11.NullValue = null;
            this.DIF_CONSUMO.DefaultCellStyle = dataGridViewCellStyle11;
            this.DIF_CONSUMO.HeaderText = "DIF CONSUMO";
            this.DIF_CONSUMO.MinimumWidth = 6;
            this.DIF_CONSUMO.Name = "DIF_CONSUMO";
            this.DIF_CONSUMO.ReadOnly = true;
            this.DIF_CONSUMO.Width = 107;
            // 
            // TOTAL_ADIF
            // 
            this.TOTAL_ADIF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.TOTAL_ADIF.DataPropertyName = "cnpr_adif";
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle12.Format = "C2";
            dataGridViewCellStyle12.NullValue = null;
            this.TOTAL_ADIF.DefaultCellStyle = dataGridViewCellStyle12;
            this.TOTAL_ADIF.HeaderText = "CNPR+CPRE ADIF";
            this.TOTAL_ADIF.MinimumWidth = 6;
            this.TOTAL_ADIF.Name = "TOTAL_ADIF";
            this.TOTAL_ADIF.ReadOnly = true;
            this.TOTAL_ADIF.Width = 124;
            // 
            // TOTAL_SCE
            // 
            this.TOTAL_SCE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.TOTAL_SCE.DataPropertyName = "cnpr_endesa";
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle13.Format = "C2";
            dataGridViewCellStyle13.NullValue = null;
            this.TOTAL_SCE.DefaultCellStyle = dataGridViewCellStyle13;
            this.TOTAL_SCE.HeaderText = "CNPR+CPRE ENDESA";
            this.TOTAL_SCE.MinimumWidth = 6;
            this.TOTAL_SCE.Name = "TOTAL_SCE";
            this.TOTAL_SCE.ReadOnly = true;
            this.TOTAL_SCE.Width = 144;
            // 
            // DIF_TOTAL
            // 
            this.DIF_TOTAL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.DIF_TOTAL.DataPropertyName = "DIF_TOTAL";
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle14.Format = "C2";
            dataGridViewCellStyle14.NullValue = null;
            this.DIF_TOTAL.DefaultCellStyle = dataGridViewCellStyle14;
            this.DIF_TOTAL.HeaderText = "DIF TOTAL";
            this.DIF_TOTAL.MinimumWidth = 6;
            this.DIF_TOTAL.Name = "DIF_TOTAL";
            this.DIF_TOTAL.ReadOnly = true;
            this.DIF_TOTAL.Width = 87;
            // 
            // menuCopyPaste
            // 
            this.menuCopyPaste.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuCopyPaste.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuCopy});
            this.menuCopyPaste.Name = "menuCopyPaste";
            this.menuCopyPaste.Size = new System.Drawing.Size(152, 26);
            this.menuCopyPaste.Opening += new System.ComponentModel.CancelEventHandler(this.menuCopyPaste_Opening);
            // 
            // toolStripMenuCopy
            // 
            this.toolStripMenuCopy.Name = "toolStripMenuCopy";
            this.toolStripMenuCopy.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.C)));
            this.toolStripMenuCopy.Size = new System.Drawing.Size(151, 22);
            this.toolStripMenuCopy.Text = "Copiar";
            // 
            // btnExcel
            // 
            this.btnExcel.Image = global::GestionOperaciones.Properties.Resources.if_microsoft_office_excel_1784856;
            this.btnExcel.Location = new System.Drawing.Point(12, 147);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(28, 28);
            this.btnExcel.TabIndex = 26;
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.importarToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1354, 24);
            this.menuStrip1.TabIndex = 29;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(60, 20);
            this.toolStripMenuItem1.Text = "Archivo";
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_dialog_close_29299;
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(138, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click_1);
            // 
            // importarToolStripMenuItem
            // 
            this.importarToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivosFacturasADIFToolStripMenuItem,
            this.archivosFacturasAdif_REG_toolStripMenuItem,
            this.toolStripSeparator1,
            this.cierresDeEnergíaToolStripMenuItem,
            this.toolStripSeparator2,
            this.parámetrosToolStripMenuItem});
            this.importarToolStripMenuItem.Name = "importarToolStripMenuItem";
            this.importarToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.importarToolStripMenuItem.Text = "Herramientas";
            // 
            // archivosFacturasADIFToolStripMenuItem
            // 
            this.archivosFacturasADIFToolStripMenuItem.Name = "archivosFacturasADIFToolStripMenuItem";
            this.archivosFacturasADIFToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.archivosFacturasADIFToolStripMenuItem.Text = "Importar archivos Facturas ADIF";
            this.archivosFacturasADIFToolStripMenuItem.Click += new System.EventHandler(this.archivosFacturasADIFToolStripMenuItem_Click);
            // 
            // archivosFacturasAdif_REG_toolStripMenuItem
            // 
            this.archivosFacturasAdif_REG_toolStripMenuItem.Enabled = false;
            this.archivosFacturasAdif_REG_toolStripMenuItem.Name = "archivosFacturasAdif_REG_toolStripMenuItem";
            this.archivosFacturasAdif_REG_toolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.archivosFacturasAdif_REG_toolStripMenuItem.Text = "Importar Regularizaciones ADIF";
            this.archivosFacturasAdif_REG_toolStripMenuItem.Click += new System.EventHandler(this.archivosFacturasAdif_REG_toolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(239, 6);
            // 
            // cierresDeEnergíaToolStripMenuItem
            // 
            this.cierresDeEnergíaToolStripMenuItem.Name = "cierresDeEnergíaToolStripMenuItem";
            this.cierresDeEnergíaToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.cierresDeEnergíaToolStripMenuItem.Text = "Cierres de energía";
            this.cierresDeEnergíaToolStripMenuItem.Click += new System.EventHandler(this.cierresDeEnergíaToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkDiff);
            this.groupBox1.Controls.Add(this.txtlote);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnBuscar);
            this.groupBox1.Controls.Add(this.txtcupsree);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dtFFACTHAS);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.dtFFACTDES);
            this.groupBox1.Location = new System.Drawing.Point(12, 41);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(656, 100);
            this.groupBox1.TabIndex = 30;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // chkDiff
            // 
            this.chkDiff.AutoSize = true;
            this.chkDiff.Location = new System.Drawing.Point(81, 70);
            this.chkDiff.Name = "chkDiff";
            this.chkDiff.Size = new System.Drawing.Size(137, 17);
            this.chkDiff.TabIndex = 34;
            this.chkDiff.Text = "Mostrar sólo diferencias";
            this.chkDiff.UseVisualStyleBackColor = true;
            // 
            // txtlote
            // 
            this.txtlote.Location = new System.Drawing.Point(305, 44);
            this.txtlote.MaxLength = 2;
            this.txtlote.Name = "txtlote";
            this.txtlote.Size = new System.Drawing.Size(47, 20);
            this.txtlote.TabIndex = 33;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(264, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 32;
            this.label4.Text = "LOTE";
            // 
            // btnBuscar
            // 
            this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
            this.btnBuscar.Location = new System.Drawing.Point(599, 19);
            this.btnBuscar.Name = "btnBuscar";
            this.btnBuscar.Size = new System.Drawing.Size(38, 39);
            this.btnBuscar.TabIndex = 31;
            this.btnBuscar.UseVisualStyleBackColor = true;
            this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click_1);
            // 
            // txtcupsree
            // 
            this.txtcupsree.Location = new System.Drawing.Point(81, 44);
            this.txtcupsree.Name = "txtcupsree";
            this.txtcupsree.Size = new System.Drawing.Size(144, 20);
            this.txtcupsree.TabIndex = 30;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "CUPSREE";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(316, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 13);
            this.label2.TabIndex = 28;
            this.label2.Text = "Hasta Fecha Consumo";
            // 
            // dtFFACTHAS
            // 
            this.dtFFACTHAS.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtFFACTHAS.Location = new System.Drawing.Point(437, 18);
            this.dtFFACTHAS.Name = "dtFFACTHAS";
            this.dtFFACTHAS.Size = new System.Drawing.Size(102, 20);
            this.dtFFACTHAS.TabIndex = 27;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 13);
            this.label1.TabIndex = 26;
            this.label1.Text = "Desde Fecha Consumo";
            // 
            // dtFFACTDES
            // 
            this.dtFFACTDES.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtFFACTDES.Location = new System.Drawing.Point(137, 19);
            this.dtFFACTDES.Name = "dtFFACTDES";
            this.dtFFACTDES.Size = new System.Drawing.Size(102, 20);
            this.dtFFACTDES.TabIndex = 25;
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(239, 6);
            // 
            // parámetrosToolStripMenuItem
            // 
            this.parámetrosToolStripMenuItem.Name = "parámetrosToolStripMenuItem";
            this.parámetrosToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.parámetrosToolStripMenuItem.Text = "Parámetros";
            this.parámetrosToolStripMenuItem.Click += new System.EventHandler(this.parámetrosToolStripMenuItem_Click);
            // 
            // FrmAdifFacturas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1354, 714);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.dgvCups);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmAdifFacturas";
            this.Text = "FACTURAS ADIF vs FACTURAS ENDESA";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmAdifFacturas_FormClosing);
            this.Load += new System.EventHandler(this.FrmAdifFacturas_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCups)).EndInit();
            this.menuCopyPaste.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dgvCups;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.ContextMenuStrip menuCopyPaste;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuCopy;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem importarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem archivosFacturasADIFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem archivosFacturasAdif_REG_toolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkDiff;
        private System.Windows.Forms.TextBox txtlote;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnBuscar;
        private System.Windows.Forms.TextBox txtcupsree;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtFFACTHAS;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtFFACTDES;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem cierresDeEnergíaToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn CUPSREE;
        private System.Windows.Forms.DataGridViewTextBoxColumn LOTE;
        private System.Windows.Forms.DataGridViewCheckBoxColumn medida_en_baja;
        private System.Windows.Forms.DataGridViewCheckBoxColumn devolucion_de_energia;
        private System.Windows.Forms.DataGridViewCheckBoxColumn cierres_energia;
        private System.Windows.Forms.DataGridViewCheckBoxColumn existe_factura_adif;
        private System.Windows.Forms.DataGridViewCheckBoxColumn existe_factura_sce;
        private System.Windows.Forms.DataGridViewTextBoxColumn CREFEREN;
        private System.Windows.Forms.DataGridViewTextBoxColumn SECFACTU;
        private System.Windows.Forms.DataGridViewTextBoxColumn FFACTURA;
        private System.Windows.Forms.DataGridViewTextBoxColumn FFACTDES;
        private System.Windows.Forms.DataGridViewTextBoxColumn FFACTHAS;
        private System.Windows.Forms.DataGridViewTextBoxColumn TFACTURA;
        private System.Windows.Forms.DataGridViewTextBoxColumn TESTFACT;
        private System.Windows.Forms.DataGridViewTextBoxColumn CONSUMO_ADIF;
        private System.Windows.Forms.DataGridViewTextBoxColumn CONSUMO_SCE;
        private System.Windows.Forms.DataGridViewTextBoxColumn DIF_CONSUMO;
        private System.Windows.Forms.DataGridViewTextBoxColumn TOTAL_ADIF;
        private System.Windows.Forms.DataGridViewTextBoxColumn TOTAL_SCE;
        private System.Windows.Forms.DataGridViewTextBoxColumn DIF_TOTAL;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem parámetrosToolStripMenuItem;
    }
}