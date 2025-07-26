namespace GestionOperaciones.forms.medida
{
    partial class FrmCurvas
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCurvas));
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btn_generar_excels = new System.Windows.Forms.Button();
            this.btnImportExcel = new System.Windows.Forms.Button();
            this.fh = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cusp20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.conexiónDataMartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilidadesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.p1TelefonicaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fechaDesde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fechaHasta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblTotalRegistros = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radio_facturada = new System.Windows.Forms.RadioButton();
            this.radio_registrada = new System.Windows.Forms.RadioButton();
            this.cmb_formato_salida = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.tabControl2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_generar_excels
            // 
            this.btn_generar_excels.BackColor = System.Drawing.Color.NavajoWhite;
            this.btn_generar_excels.ForeColor = System.Drawing.Color.Black;
            this.btn_generar_excels.Location = new System.Drawing.Point(171, 8);
            this.btn_generar_excels.Name = "btn_generar_excels";
            this.btn_generar_excels.Size = new System.Drawing.Size(136, 55);
            this.btn_generar_excels.TabIndex = 44;
            this.btn_generar_excels.Text = "Generar Ficheros Excels";
            this.toolTip1.SetToolTip(this.btn_generar_excels, "Generará un Excel por cada CUPS  en la carpeta c:\\Temp");
            this.btn_generar_excels.UseVisualStyleBackColor = false;
            this.btn_generar_excels.Click += new System.EventHandler(this.btn_generar_excels_Click);
            // 
            // btnImportExcel
            // 
            this.btnImportExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnImportExcel.Location = new System.Drawing.Point(6, 8);
            this.btnImportExcel.Name = "btnImportExcel";
            this.btnImportExcel.Size = new System.Drawing.Size(159, 55);
            this.btnImportExcel.TabIndex = 27;
            this.btnImportExcel.Text = "Importar CUPS desde Excel";
            this.toolTip1.SetToolTip(this.btnImportExcel, "Importa un Excel con el formato en 3 columnas de:\r\n- CUPS20 \r\n- FECHA DESDE\r\n- FE" +
        "CHA HASTA\r\n\r\n");
            this.btnImportExcel.UseVisualStyleBackColor = true;
            this.btnImportExcel.Click += new System.EventHandler(this.btnImportExcel_Click);
            // 
            // fh
            // 
            this.fh.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fh.DataPropertyName = "fh";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fh.DefaultCellStyle = dataGridViewCellStyle1;
            this.fh.HeaderText = "Fecha hasta";
            this.fh.Name = "fh";
            // 
            // fd
            // 
            this.fd.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fd.DataPropertyName = "fd";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fd.DefaultCellStyle = dataGridViewCellStyle2;
            this.fd.HeaderText = "Fecha desde";
            this.fd.Name = "fd";
            // 
            // cusp20
            // 
            this.cusp20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cusp20.DataPropertyName = "cups20";
            this.cusp20.HeaderText = "CUPS20";
            this.cusp20.Name = "cusp20";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.editarToolStripMenuItem,
            this.utilidadesToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(593, 24);
            this.menuStrip1.TabIndex = 44;
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
            // editarToolStripMenuItem
            // 
            this.editarToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.conexiónDataMartToolStripMenuItem});
            this.editarToolStripMenuItem.Name = "editarToolStripMenuItem";
            this.editarToolStripMenuItem.Size = new System.Drawing.Size(49, 20);
            this.editarToolStripMenuItem.Text = "Editar";
            // 
            // conexiónDataMartToolStripMenuItem
            // 
            this.conexiónDataMartToolStripMenuItem.Enabled = false;
            this.conexiónDataMartToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.password_icon;
            this.conexiónDataMartToolStripMenuItem.Name = "conexiónDataMartToolStripMenuItem";
            this.conexiónDataMartToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.conexiónDataMartToolStripMenuItem.Text = "Conexión Datamart";
            this.conexiónDataMartToolStripMenuItem.Click += new System.EventHandler(this.conexiónDataMartToolStripMenuItem_Click);
            // 
            // utilidadesToolStripMenuItem
            // 
            this.utilidadesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.p1TelefonicaToolStripMenuItem});
            this.utilidadesToolStripMenuItem.Name = "utilidadesToolStripMenuItem";
            this.utilidadesToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
            this.utilidadesToolStripMenuItem.Text = "Utilidades";
            // 
            // p1TelefonicaToolStripMenuItem
            // 
            this.p1TelefonicaToolStripMenuItem.Name = "p1TelefonicaToolStripMenuItem";
            this.p1TelefonicaToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.p1TelefonicaToolStripMenuItem.Text = "P1 Telefonica";
            // 
            // tabControl2
            // 
            this.tabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl2.Controls.Add(this.tabPage3);
            this.tabControl2.Location = new System.Drawing.Point(12, 202);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(569, 487);
            this.tabControl2.TabIndex = 46;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.tabControl1);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(561, 461);
            this.tabPage3.TabIndex = 1;
            this.tabPage3.Text = "Consulta masiva";
            this.tabPage3.UseVisualStyleBackColor = true;
            this.tabPage3.Click += new System.EventHandler(this.tabPage3_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(6, 6);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(529, 449);
            this.tabControl1.TabIndex = 45;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btn_generar_excels);
            this.tabPage1.Controls.Add(this.btnImportExcel);
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Controls.Add(this.lblTotalRegistros);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(521, 423);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Lista de CUPS";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            this.dgv.AllowUserToOrderColumns = true;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cups20,
            this.fechaDesde,
            this.fechaHasta});
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle6;
            this.dgv.Location = new System.Drawing.Point(6, 69);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(509, 348);
            this.dgv.TabIndex = 43;
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "cups20";
            this.cups20.HeaderText = "CUPS20";
            this.cups20.Name = "cups20";
            this.cups20.Width = 73;
            // 
            // fechaDesde
            // 
            this.fechaDesde.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fechaDesde.DataPropertyName = "fd";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fechaDesde.DefaultCellStyle = dataGridViewCellStyle4;
            this.fechaDesde.HeaderText = "Fecha Desde";
            this.fechaDesde.Name = "fechaDesde";
            this.fechaDesde.Width = 96;
            // 
            // fechaHasta
            // 
            this.fechaHasta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fechaHasta.DataPropertyName = "fh";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fechaHasta.DefaultCellStyle = dataGridViewCellStyle5;
            this.fechaHasta.HeaderText = "fechaHasta";
            this.fechaHasta.Name = "fechaHasta";
            this.fechaHasta.Width = 87;
            // 
            // lblTotalRegistros
            // 
            this.lblTotalRegistros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTotalRegistros.AutoSize = true;
            this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalRegistros.Location = new System.Drawing.Point(367, 10);
            this.lblTotalRegistros.Name = "lblTotalRegistros";
            this.lblTotalRegistros.Size = new System.Drawing.Size(97, 13);
            this.lblTotalRegistros.TabIndex = 42;
            this.lblTotalRegistros.Text = "Total Registros:";
            this.lblTotalRegistros.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radio_facturada);
            this.groupBox1.Controls.Add(this.radio_registrada);
            this.groupBox1.Location = new System.Drawing.Point(16, 52);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(111, 105);
            this.groupBox1.TabIndex = 47;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tipo de Curva";
            // 
            // radio_facturada
            // 
            this.radio_facturada.AutoSize = true;
            this.radio_facturada.Location = new System.Drawing.Point(10, 57);
            this.radio_facturada.Name = "radio_facturada";
            this.radio_facturada.Size = new System.Drawing.Size(73, 17);
            this.radio_facturada.TabIndex = 3;
            this.radio_facturada.TabStop = true;
            this.radio_facturada.Text = "Facturada";
            this.radio_facturada.UseVisualStyleBackColor = true;
            // 
            // radio_registrada
            // 
            this.radio_registrada.AutoSize = true;
            this.radio_registrada.Location = new System.Drawing.Point(10, 34);
            this.radio_registrada.Name = "radio_registrada";
            this.radio_registrada.Size = new System.Drawing.Size(76, 17);
            this.radio_registrada.TabIndex = 2;
            this.radio_registrada.TabStop = true;
            this.radio_registrada.Text = "Registrada";
            this.radio_registrada.UseVisualStyleBackColor = true;
            // 
            // cmb_formato_salida
            // 
            this.cmb_formato_salida.FormattingEnabled = true;
            this.cmb_formato_salida.Items.AddRange(new object[] {
            "Formato Excel Gestores",
            "Formato Excel CuartoHoraria (SCE)",
            "Formato ADIF"});
            this.cmb_formato_salida.Location = new System.Drawing.Point(380, 186);
            this.cmb_formato_salida.Name = "cmb_formato_salida";
            this.cmb_formato_salida.Size = new System.Drawing.Size(194, 21);
            this.cmb_formato_salida.TabIndex = 48;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(297, 189);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 49;
            this.label1.Text = "Formato Salida";
            // 
            // FrmCurvas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 701);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmb_formato_salida);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.tabControl2);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmCurvas";
            this.Text = "Curvas  Datamart";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmCurvas_FormClosing);
            this.Load += new System.EventHandler(this.FrmCurvas_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.DataGridViewTextBoxColumn fh;
        private System.Windows.Forms.DataGridViewTextBoxColumn fd;
        private System.Windows.Forms.DataGridViewTextBoxColumn cusp20;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem conexiónDataMartToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem utilidadesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem p1TelefonicaToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button btn_generar_excels;
        private System.Windows.Forms.Button btnImportExcel;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn fechaDesde;
        private System.Windows.Forms.DataGridViewTextBoxColumn fechaHasta;
        private System.Windows.Forms.Label lblTotalRegistros;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radio_registrada;
        private System.Windows.Forms.RadioButton radio_facturada;
        private System.Windows.Forms.ComboBox cmb_formato_salida;
        private System.Windows.Forms.Label label1;
    }
}