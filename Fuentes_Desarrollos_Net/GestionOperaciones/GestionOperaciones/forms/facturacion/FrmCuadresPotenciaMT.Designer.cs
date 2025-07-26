
namespace GestionOperaciones.forms.facturacion
{
    partial class FrmCuadresPotenciaMT
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCuadresPotenciaMT));
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_cnifdnic = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_cups20 = new System.Windows.Forms.TextBox();
            this.cmdSearch = new System.Windows.Forms.Button();
            this.txt_fh = new System.Windows.Forms.DateTimePicker();
            this.txt_fd = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.lbl_total_registros = new System.Windows.Forms.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.dapersoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cnifdnic = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.distribuidora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ffactura = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ffactdes = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ffacthas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ifactura = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tconfac4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.iconfac4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.potencia_contratada = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1012, 24);
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_cnifdnic);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txt_cups20);
            this.groupBox1.Controls.Add(this.cmdSearch);
            this.groupBox1.Controls.Add(this.txt_fh);
            this.groupBox1.Controls.Add(this.txt_fd);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(616, 95);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(337, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 13);
            this.label3.TabIndex = 74;
            this.label3.Text = "NIF:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txt_cnifdnic
            // 
            this.txt_cnifdnic.Location = new System.Drawing.Point(370, 45);
            this.txt_cnifdnic.Name = "txt_cnifdnic";
            this.txt_cnifdnic.Size = new System.Drawing.Size(109, 20);
            this.txt_cnifdnic.TabIndex = 73;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(35, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 72;
            this.label4.Text = "CUPS20:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txt_cups20
            // 
            this.txt_cups20.Location = new System.Drawing.Point(92, 45);
            this.txt_cups20.Name = "txt_cups20";
            this.txt_cups20.Size = new System.Drawing.Size(189, 20);
            this.txt_cups20.TabIndex = 71;
            // 
            // cmdSearch
            // 
            this.cmdSearch.Image = ((System.Drawing.Image)(resources.GetObject("cmdSearch.Image")));
            this.cmdSearch.Location = new System.Drawing.Point(525, 22);
            this.cmdSearch.Name = "cmdSearch";
            this.cmdSearch.Size = new System.Drawing.Size(69, 46);
            this.cmdSearch.TabIndex = 70;
            this.cmdSearch.UseVisualStyleBackColor = true;
            this.cmdSearch.Click += new System.EventHandler(this.cmdSearch_Click);
            // 
            // txt_fh
            // 
            this.txt_fh.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fh.Location = new System.Drawing.Point(383, 19);
            this.txt_fh.Name = "txt_fh";
            this.txt_fh.Size = new System.Drawing.Size(96, 20);
            this.txt_fh.TabIndex = 1;
            // 
            // txt_fd
            // 
            this.txt_fd.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fd.Location = new System.Drawing.Point(144, 19);
            this.txt_fd.Name = "txt_fd";
            this.txt_fd.Size = new System.Drawing.Size(96, 20);
            this.txt_fd.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Fecha Consumo Desde";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(262, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Fecha Consumo Hasta";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(6, 128);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1001, 688);
            this.tabControl1.TabIndex = 18;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.cmdExcel);
            this.tabPage2.Controls.Add(this.lbl_total_registros);
            this.tabPage2.Controls.Add(this.dgv);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(993, 662);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Facturas";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(8, 6);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 62;
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // lbl_total_registros
            // 
            this.lbl_total_registros.AutoSize = true;
            this.lbl_total_registros.Location = new System.Drawing.Point(814, 14);
            this.lbl_total_registros.Name = "lbl_total_registros";
            this.lbl_total_registros.Size = new System.Drawing.Size(121, 13);
            this.lbl_total_registros.TabIndex = 61;
            this.lbl_total_registros.Text = "Total registros:               ";
            this.lbl_total_registros.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dapersoc,
            this.cnifdnic,
            this.distribuidora,
            this.ffactura,
            this.ffactdes,
            this.ffacthas,
            this.ifactura,
            this.tconfac4,
            this.iconfac4,
            this.cups20,
            this.potencia_contratada});
            this.dgv.Location = new System.Drawing.Point(6, 40);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(981, 616);
            this.dgv.TabIndex = 60;
            // 
            // dapersoc
            // 
            this.dapersoc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dapersoc.DataPropertyName = "dapersoc";
            this.dapersoc.HeaderText = "DAPERSOC";
            this.dapersoc.Name = "dapersoc";
            this.dapersoc.ReadOnly = true;
            this.dapersoc.Width = 91;
            // 
            // cnifdnic
            // 
            this.cnifdnic.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cnifdnic.DataPropertyName = "cnifdnic";
            this.cnifdnic.HeaderText = "CNIFDNIC";
            this.cnifdnic.Name = "cnifdnic";
            this.cnifdnic.ReadOnly = true;
            this.cnifdnic.Width = 82;
            // 
            // distribuidora
            // 
            this.distribuidora.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.distribuidora.DataPropertyName = "cfactura";
            this.distribuidora.HeaderText = "CFACTURA";
            this.distribuidora.Name = "distribuidora";
            this.distribuidora.ReadOnly = true;
            this.distribuidora.Width = 89;
            // 
            // ffactura
            // 
            this.ffactura.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ffactura.DataPropertyName = "ffactura";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.ffactura.DefaultCellStyle = dataGridViewCellStyle2;
            this.ffactura.HeaderText = "FFACTURA";
            this.ffactura.Name = "ffactura";
            this.ffactura.ReadOnly = true;
            this.ffactura.Width = 88;
            // 
            // ffactdes
            // 
            this.ffactdes.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ffactdes.DataPropertyName = "ffactdes";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.ffactdes.DefaultCellStyle = dataGridViewCellStyle3;
            this.ffactdes.HeaderText = "FFACTDES";
            this.ffactdes.Name = "ffactdes";
            this.ffactdes.ReadOnly = true;
            this.ffactdes.Width = 87;
            // 
            // ffacthas
            // 
            this.ffacthas.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ffacthas.DataPropertyName = "ffacthas";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.ffacthas.DefaultCellStyle = dataGridViewCellStyle4;
            this.ffacthas.HeaderText = "FFACTHAS";
            this.ffacthas.Name = "ffacthas";
            this.ffacthas.ReadOnly = true;
            this.ffacthas.Width = 87;
            // 
            // ifactura
            // 
            this.ifactura.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ifactura.DataPropertyName = "ifactura";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N2";
            dataGridViewCellStyle5.NullValue = null;
            this.ifactura.DefaultCellStyle = dataGridViewCellStyle5;
            this.ifactura.HeaderText = "IFACTURA";
            this.ifactura.Name = "ifactura";
            this.ifactura.ReadOnly = true;
            this.ifactura.Width = 85;
            // 
            // tconfac4
            // 
            this.tconfac4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tconfac4.DataPropertyName = "tconfac4";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            this.tconfac4.DefaultCellStyle = dataGridViewCellStyle6;
            this.tconfac4.HeaderText = "TCONFAC4";
            this.tconfac4.Name = "tconfac4";
            this.tconfac4.ReadOnly = true;
            this.tconfac4.Width = 88;
            // 
            // iconfac4
            // 
            this.iconfac4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.iconfac4.DataPropertyName = "iconfac4";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle7.Format = "N2";
            dataGridViewCellStyle7.NullValue = null;
            this.iconfac4.DefaultCellStyle = dataGridViewCellStyle7;
            this.iconfac4.HeaderText = "ICONFAC4";
            this.iconfac4.Name = "iconfac4";
            this.iconfac4.ReadOnly = true;
            this.iconfac4.Width = 84;
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "cupsree";
            this.cups20.HeaderText = "CUPSREE";
            this.cups20.Name = "cups20";
            this.cups20.ReadOnly = true;
            this.cups20.Width = 83;
            // 
            // potencia_contratada
            // 
            this.potencia_contratada.DataPropertyName = "potencia_contratada";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.potencia_contratada.DefaultCellStyle = dataGridViewCellStyle8;
            this.potencia_contratada.HeaderText = "POTENCIA_CONTRATADA";
            this.potencia_contratada.Name = "potencia_contratada";
            this.potencia_contratada.ReadOnly = true;
            // 
            // FrmCuadresPotenciaMT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1012, 821);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmCuadresPotenciaMT";
            this.Text = "Cuadres de Potencia MT";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmCuadresPotenciaMT_FormClosing);
            this.Load += new System.EventHandler(this.FrmCuadresPotenciaMT_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.DateTimePicker txt_fh;
        public System.Windows.Forms.DateTimePicker txt_fd;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button cmdSearch;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_cnifdnic;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_cups20;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Label lbl_total_registros;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn dapersoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn cnifdnic;
        private System.Windows.Forms.DataGridViewTextBoxColumn distribuidora;
        private System.Windows.Forms.DataGridViewTextBoxColumn ffactura;
        private System.Windows.Forms.DataGridViewTextBoxColumn ffactdes;
        private System.Windows.Forms.DataGridViewTextBoxColumn ffacthas;
        private System.Windows.Forms.DataGridViewTextBoxColumn ifactura;
        private System.Windows.Forms.DataGridViewTextBoxColumn tconfac4;
        private System.Windows.Forms.DataGridViewTextBoxColumn iconfac4;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn potencia_contratada;
    }
}