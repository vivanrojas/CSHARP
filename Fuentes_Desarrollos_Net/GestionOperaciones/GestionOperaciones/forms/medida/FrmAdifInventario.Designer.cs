namespace GestionOperaciones.forms.medida
{
    partial class FrmAdifInventario
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAdifInventario));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chk_perdidas = new System.Windows.Forms.CheckBox();
            this.chk_cierres_energia = new System.Windows.Forms.CheckBox();
            this.chk_devolucion_energia = new System.Windows.Forms.CheckBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.chk_medida_en_baja = new System.Windows.Forms.CheckBox();
            this.txtFH = new System.Windows.Forms.DateTimePicker();
            this.txtFD = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.listBoxLotes = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtcups20 = new System.Windows.Forms.TextBox();
            this.dgvInventario = new System.Windows.Forms.DataGridView();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lote = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_desde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_hasta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tarifa = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.estado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.devolucion_energia = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.medida_en_baja = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.cierres_energia = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.perdidas = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.comentarios = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.lbl_total_cups = new System.Windows.Forms.Label();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dvgLotes = new System.Windows.Forms.DataGridView();
            this.dgv2_lote = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_2_numcups = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvInventario)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dvgLotes)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chk_perdidas);
            this.groupBox1.Controls.Add(this.chk_cierres_energia);
            this.groupBox1.Controls.Add(this.chk_devolucion_energia);
            this.groupBox1.Controls.Add(this.btnSearch);
            this.groupBox1.Controls.Add(this.chk_medida_en_baja);
            this.groupBox1.Controls.Add(this.txtFH);
            this.groupBox1.Controls.Add(this.txtFD);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.listBoxLotes);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtcups20);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 40);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(639, 100);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // chk_perdidas
            // 
            this.chk_perdidas.AutoSize = true;
            this.chk_perdidas.Location = new System.Drawing.Point(313, 72);
            this.chk_perdidas.Name = "chk_perdidas";
            this.chk_perdidas.Size = new System.Drawing.Size(60, 16);
            this.chk_perdidas.TabIndex = 46;
            this.chk_perdidas.Text = "Pérdidas";
            this.toolTip.SetToolTip(this.chk_perdidas, "Filtra los registros que tienen pérdidas informadas > 0\r\n");
            this.chk_perdidas.UseVisualStyleBackColor = true;
            // 
            // chk_cierres_energia
            // 
            this.chk_cierres_energia.AutoSize = true;
            this.chk_cierres_energia.Location = new System.Drawing.Point(218, 72);
            this.chk_cierres_energia.Name = "chk_cierres_energia";
            this.chk_cierres_energia.Size = new System.Drawing.Size(86, 16);
            this.chk_cierres_energia.TabIndex = 45;
            this.chk_cierres_energia.Text = "Cierres energía";
            this.chk_cierres_energia.UseVisualStyleBackColor = true;
            this.chk_cierres_energia.CheckedChanged += new System.EventHandler(this.chk_cierres_energia_CheckedChanged);
            // 
            // chk_devolucion_energia
            // 
            this.chk_devolucion_energia.AutoSize = true;
            this.chk_devolucion_energia.Location = new System.Drawing.Point(106, 72);
            this.chk_devolucion_energia.Name = "chk_devolucion_energia";
            this.chk_devolucion_energia.Size = new System.Drawing.Size(103, 16);
            this.chk_devolucion_energia.TabIndex = 44;
            this.chk_devolucion_energia.Text = "Devolución energía";
            this.chk_devolucion_energia.UseVisualStyleBackColor = true;
            this.chk_devolucion_energia.CheckedChanged += new System.EventHandler(this.chk_devolucion_energia_CheckedChanged);
            // 
            // btnSearch
            // 
            this.btnSearch.Image = ((System.Drawing.Image)(resources.GetObject("btnSearch.Image")));
            this.btnSearch.Location = new System.Drawing.Point(563, 24);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(59, 58);
            this.btnSearch.TabIndex = 16;
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // chk_medida_en_baja
            // 
            this.chk_medida_en_baja.AutoSize = true;
            this.chk_medida_en_baja.Location = new System.Drawing.Point(11, 72);
            this.chk_medida_en_baja.Name = "chk_medida_en_baja";
            this.chk_medida_en_baja.Size = new System.Drawing.Size(86, 16);
            this.chk_medida_en_baja.TabIndex = 43;
            this.chk_medida_en_baja.Text = "Medida en baja";
            this.chk_medida_en_baja.UseVisualStyleBackColor = true;
            // 
            // txtFH
            // 
            this.txtFH.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFH.Location = new System.Drawing.Point(289, 42);
            this.txtFH.Name = "txtFH";
            this.txtFH.Size = new System.Drawing.Size(88, 18);
            this.txtFH.TabIndex = 12;
            // 
            // txtFD
            // 
            this.txtFD.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFD.Location = new System.Drawing.Point(90, 42);
            this.txtFD.Name = "txtFD";
            this.txtFD.Size = new System.Drawing.Size(88, 18);
            this.txtFD.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(427, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(34, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "LOTES";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(211, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "FECHA HASTA";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "FECHA DESDE";
            // 
            // listBoxLotes
            // 
            this.listBoxLotes.FormattingEnabled = true;
            this.listBoxLotes.ItemHeight = 12;
            this.listBoxLotes.Location = new System.Drawing.Point(467, 13);
            this.listBoxLotes.Name = "listBoxLotes";
            this.listBoxLotes.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxLotes.Size = new System.Drawing.Size(85, 76);
            this.listBoxLotes.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "CUPS20";
            // 
            // txtcups20
            // 
            this.txtcups20.Location = new System.Drawing.Point(90, 18);
            this.txtcups20.Name = "txtcups20";
            this.txtcups20.Size = new System.Drawing.Size(145, 18);
            this.txtcups20.TabIndex = 0;
            // 
            // dgvInventario
            // 
            this.dgvInventario.AllowUserToAddRows = false;
            this.dgvInventario.AllowUserToDeleteRows = false;
            this.dgvInventario.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dgvInventario.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvInventario.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvInventario.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvInventario.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvInventario.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cups20,
            this.lote,
            this.fecha_desde,
            this.fecha_hasta,
            this.tarifa,
            this.estado,
            this.devolucion_energia,
            this.medida_en_baja,
            this.cierres_energia,
            this.perdidas,
            this.comentarios});
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvInventario.DefaultCellStyle = dataGridViewCellStyle9;
            this.dgvInventario.Location = new System.Drawing.Point(6, 48);
            this.dgvInventario.Name = "dgvInventario";
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvInventario.RowHeadersDefaultCellStyle = dataGridViewCellStyle10;
            this.dgvInventario.Size = new System.Drawing.Size(1056, 479);
            this.dgvInventario.TabIndex = 1;
            this.dgvInventario.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvInventario_CellClick);
            this.dgvInventario.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvInventario_CellContentClick);
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "cups20";
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cups20.DefaultCellStyle = dataGridViewCellStyle3;
            this.cups20.HeaderText = "CUPS20";
            this.cups20.Name = "cups20";
            this.cups20.Width = 73;
            // 
            // lote
            // 
            this.lote.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.lote.DataPropertyName = "lote";
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lote.DefaultCellStyle = dataGridViewCellStyle4;
            this.lote.HeaderText = "LOTE";
            this.lote.Name = "lote";
            this.lote.Width = 60;
            // 
            // fecha_desde
            // 
            this.fecha_desde.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_desde.DataPropertyName = "vigencia_desde";
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fecha_desde.DefaultCellStyle = dataGridViewCellStyle5;
            this.fecha_desde.HeaderText = "FECHA DESDE";
            this.fecha_desde.Name = "fecha_desde";
            this.fecha_desde.Width = 98;
            // 
            // fecha_hasta
            // 
            this.fecha_hasta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_hasta.DataPropertyName = "vigencia_hasta";
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fecha_hasta.DefaultCellStyle = dataGridViewCellStyle6;
            this.fecha_hasta.HeaderText = "FECHA HASTA";
            this.fecha_hasta.Name = "fecha_hasta";
            this.fecha_hasta.Width = 97;
            // 
            // tarifa
            // 
            this.tarifa.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tarifa.DataPropertyName = "tarifa";
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tarifa.DefaultCellStyle = dataGridViewCellStyle7;
            this.tarifa.HeaderText = "TARIFA";
            this.tarifa.Name = "tarifa";
            this.tarifa.Width = 70;
            // 
            // estado
            // 
            this.estado.DataPropertyName = "estado_contrato";
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.estado.DefaultCellStyle = dataGridViewCellStyle8;
            this.estado.HeaderText = "ESTADO";
            this.estado.Name = "estado";
            // 
            // devolucion_energia
            // 
            this.devolucion_energia.DataPropertyName = "devolucion_de_energia";
            this.devolucion_energia.HeaderText = "Devolución energía";
            this.devolucion_energia.Name = "devolucion_energia";
            this.devolucion_energia.ReadOnly = true;
            this.devolucion_energia.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // medida_en_baja
            // 
            this.medida_en_baja.DataPropertyName = "medida_en_baja";
            this.medida_en_baja.HeaderText = "Medida en baja";
            this.medida_en_baja.Name = "medida_en_baja";
            this.medida_en_baja.ReadOnly = true;
            this.medida_en_baja.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // cierres_energia
            // 
            this.cierres_energia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cierres_energia.DataPropertyName = "cierres_energia";
            this.cierres_energia.HeaderText = "Cierres de energía";
            this.cierres_energia.Name = "cierres_energia";
            this.cierres_energia.ReadOnly = true;
            this.cierres_energia.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.cierres_energia.Width = 109;
            // 
            // perdidas
            // 
            this.perdidas.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.perdidas.DataPropertyName = "tiene_perdidas";
            this.perdidas.HeaderText = "Pérdidas";
            this.perdidas.Name = "perdidas";
            this.perdidas.Width = 54;
            // 
            // comentarios
            // 
            this.comentarios.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios.DataPropertyName = "comentarios";
            this.comentarios.HeaderText = "Comentarios";
            this.comentarios.Name = "comentarios";
            this.comentarios.Width = 90;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1100, 24);
            this.menuStrip1.TabIndex = 4;
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
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 178);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1076, 559);
            this.tabControl1.TabIndex = 5;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnEdit);
            this.tabPage2.Controls.Add(this.btnAdd);
            this.tabPage2.Controls.Add(this.button2);
            this.tabPage2.Controls.Add(this.lbl_total_cups);
            this.tabPage2.Controls.Add(this.cmdExcel);
            this.tabPage2.Controls.Add(this.dgvInventario);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1068, 533);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Inventario";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnEdit
            // 
            this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
            this.btnEdit.Location = new System.Drawing.Point(43, 10);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(28, 28);
            this.btnEdit.TabIndex = 66;
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click_1);
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.Location = new System.Drawing.Point(9, 10);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(28, 28);
            this.btnAdd.TabIndex = 63;
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Location = new System.Drawing.Point(77, 10);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(28, 28);
            this.button2.TabIndex = 65;
            this.button2.UseVisualStyleBackColor = true;
            // 
            // lbl_total_cups
            // 
            this.lbl_total_cups.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total_cups.AutoSize = true;
            this.lbl_total_cups.Location = new System.Drawing.Point(876, 14);
            this.lbl_total_cups.Name = "lbl_total_cups";
            this.lbl_total_cups.Size = new System.Drawing.Size(127, 13);
            this.lbl_total_cups.TabIndex = 62;
            this.lbl_total_cups.Text = "Total registros:                 ";
            this.lbl_total_cups.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(111, 10);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 26;
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dvgLotes);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1068, 533);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "Control inventario";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dvgLotes
            // 
            this.dvgLotes.AllowUserToAddRows = false;
            this.dvgLotes.AllowUserToDeleteRows = false;
            this.dvgLotes.AllowUserToOrderColumns = true;
            dataGridViewCellStyle11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dvgLotes.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle11;
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dvgLotes.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle12;
            this.dvgLotes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dvgLotes.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dgv2_lote,
            this.dgv_2_numcups});
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dvgLotes.DefaultCellStyle = dataGridViewCellStyle15;
            this.dvgLotes.Location = new System.Drawing.Point(6, 6);
            this.dvgLotes.Name = "dvgLotes";
            this.dvgLotes.Size = new System.Drawing.Size(196, 521);
            this.dvgLotes.TabIndex = 1;
            // 
            // dgv2_lote
            // 
            this.dgv2_lote.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCellStyle13.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgv2_lote.DefaultCellStyle = dataGridViewCellStyle13;
            this.dgv2_lote.HeaderText = "LOTE";
            this.dgv2_lote.Name = "dgv2_lote";
            this.dgv2_lote.Width = 53;
            // 
            // dgv_2_numcups
            // 
            this.dgv_2_numcups.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCellStyle14.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgv_2_numcups.DefaultCellStyle = dataGridViewCellStyle14;
            this.dgv_2_numcups.HeaderText = "NUM. CUPS";
            this.dgv_2_numcups.Name = "dgv_2_numcups";
            this.dgv_2_numcups.Width = 84;
            // 
            // FrmAdifInventario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1100, 749);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmAdifInventario";
            this.Text = "Inventario";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmAdifInventario_FormClosing);
            this.Load += new System.EventHandler(this.FrmAdifInventario_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvInventario)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dvgLotes)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox listBoxLotes;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtcups20;
        private System.Windows.Forms.DataGridView dgvInventario;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DateTimePicker txtFH;
        private System.Windows.Forms.DateTimePicker txtFD;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.CheckBox chk_devolucion_energia;
        private System.Windows.Forms.CheckBox chk_medida_en_baja;
        private System.Windows.Forms.CheckBox chk_cierres_energia;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dvgLotes;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv2_lote;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_2_numcups;
        private System.Windows.Forms.Label lbl_total_cups;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.CheckBox chk_perdidas;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn lote;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_desde;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_hasta;
        private System.Windows.Forms.DataGridViewTextBoxColumn tarifa;
        private System.Windows.Forms.DataGridViewTextBoxColumn estado;
        private System.Windows.Forms.DataGridViewCheckBoxColumn devolucion_energia;
        private System.Windows.Forms.DataGridViewCheckBoxColumn medida_en_baja;
        private System.Windows.Forms.DataGridViewCheckBoxColumn cierres_energia;
        private System.Windows.Forms.DataGridViewCheckBoxColumn perdidas;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios;
    }
}