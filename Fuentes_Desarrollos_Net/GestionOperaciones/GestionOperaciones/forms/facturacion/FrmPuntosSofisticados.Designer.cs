namespace GestionOperaciones.forms.facturacion
{
    partial class FrmPuntosSofisticados
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
#pragma warning disable CS0114 // 'FrmPuntosSofisticados.Dispose(bool)' oculta el miembro heredado 'Form.Dispose(bool)'. Para hacer que el miembro actual invalide esa implementación, agregue la palabra clave override. Si no, agregue la palabra clave new.
        protected void Dispose(bool disposing)
#pragma warning restore CS0114 // 'FrmPuntosSofisticados.Dispose(bool)' oculta el miembro heredado 'Form.Dispose(bool)'. Para hacer que el miembro actual invalide esa implementación, agregue la palabra clave override. Si no, agregue la palabra clave new.
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPuntosSofisticados));
            this.txtCCOUNIPS = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.dgvFacturas = new System.Windows.Forms.DataGridView();
            this.CCOUNIPS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cusp20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.codintegr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GRUPO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRECIOS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FACTURAS_A_CUENTA = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.toolTipNIF = new System.Windows.Forms.ToolTip(this.components);
            this.btnDel = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.toolTipCUPSREE = new System.Windows.Forms.ToolTip(this.components);
            this.cmdSearch = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFacturas)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtCCOUNIPS
            // 
            this.txtCCOUNIPS.Location = new System.Drawing.Point(89, 50);
            this.txtCCOUNIPS.Name = "txtCCOUNIPS";
            this.txtCCOUNIPS.Size = new System.Drawing.Size(111, 20);
            this.txtCCOUNIPS.TabIndex = 8;
            this.toolTipCUPSREE.SetToolTip(this.txtCCOUNIPS, "Filtro por CUPSREE");
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(21, 53);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "CCOUNIPS";
            // 
            // dgvFacturas
            // 
            this.dgvFacturas.AllowUserToAddRows = false;
            this.dgvFacturas.AllowUserToDeleteRows = false;
            this.dgvFacturas.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.SandyBrown;
            this.dgvFacturas.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvFacturas.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvFacturas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFacturas.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CCOUNIPS,
            this.cusp20,
            this.codintegr,
            this.GRUPO,
            this.FD,
            this.FH,
            this.PRECIOS,
            this.FACTURAS_A_CUENTA});
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvFacturas.DefaultCellStyle = dataGridViewCellStyle6;
            this.dgvFacturas.Location = new System.Drawing.Point(6, 40);
            this.dgvFacturas.Name = "dgvFacturas";
            this.dgvFacturas.ReadOnly = true;
            this.dgvFacturas.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvFacturas.Size = new System.Drawing.Size(927, 285);
            this.dgvFacturas.TabIndex = 15;
            // 
            // CCOUNIPS
            // 
            this.CCOUNIPS.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.CCOUNIPS.DataPropertyName = "cups13";
            this.CCOUNIPS.HeaderText = "CUPS13";
            this.CCOUNIPS.Name = "CCOUNIPS";
            this.CCOUNIPS.ReadOnly = true;
            this.CCOUNIPS.Width = 79;
            // 
            // cusp20
            // 
            this.cusp20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cusp20.DataPropertyName = "cups20";
            this.cusp20.HeaderText = "CUPS20";
            this.cusp20.Name = "cusp20";
            this.cusp20.ReadOnly = true;
            this.cusp20.Width = 79;
            // 
            // codintegr
            // 
            this.codintegr.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.codintegr.DataPropertyName = "dapersoc";
            this.codintegr.HeaderText = "DAPERSOC";
            this.codintegr.Name = "codintegr";
            this.codintegr.ReadOnly = true;
            this.codintegr.Width = 98;
            // 
            // GRUPO
            // 
            this.GRUPO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.GRUPO.DataPropertyName = "grupo";
            this.GRUPO.HeaderText = "GRUPO";
            this.GRUPO.Name = "GRUPO";
            this.GRUPO.ReadOnly = true;
            this.GRUPO.Width = 76;
            // 
            // FD
            // 
            this.FD.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FD.DataPropertyName = "fd";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FD.DefaultCellStyle = dataGridViewCellStyle3;
            this.FD.HeaderText = "FD";
            this.FD.Name = "FD";
            this.FD.ReadOnly = true;
            this.FD.Width = 48;
            // 
            // FH
            // 
            this.FH.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FH.DataPropertyName = "fh";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FH.DefaultCellStyle = dataGridViewCellStyle4;
            this.FH.HeaderText = "FH";
            this.FH.Name = "FH";
            this.FH.ReadOnly = true;
            this.FH.Width = 48;
            // 
            // PRECIOS
            // 
            this.PRECIOS.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.PRECIOS.DataPropertyName = "precios";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.PRECIOS.DefaultCellStyle = dataGridViewCellStyle5;
            this.PRECIOS.HeaderText = "PRECIOS";
            this.PRECIOS.Name = "PRECIOS";
            this.PRECIOS.ReadOnly = true;
            this.PRECIOS.Width = 85;
            // 
            // FACTURAS_A_CUENTA
            // 
            this.FACTURAS_A_CUENTA.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FACTURAS_A_CUENTA.DataPropertyName = "facturas_a_cuenta";
            this.FACTURAS_A_CUENTA.FalseValue = "N";
            this.FACTURAS_A_CUENTA.HeaderText = "FACTURAS A CUENTA";
            this.FACTURAS_A_CUENTA.Name = "FACTURAS_A_CUENTA";
            this.FACTURAS_A_CUENTA.ReadOnly = true;
            this.FACTURAS_A_CUENTA.TrueValue = "S";
            this.FACTURAS_A_CUENTA.Width = 79;
            // 
            // toolTipNIF
            // 
            this.toolTipNIF.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTipNIF_Popup);
            // 
            // btnDel
            // 
            this.btnDel.Image = ((System.Drawing.Image)(resources.GetObject("btnDel.Image")));
            this.btnDel.Location = new System.Drawing.Point(73, 6);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(28, 28);
            this.btnDel.TabIndex = 44;
            this.toolTipNIF.SetToolTip(this.btnDel, "Elimina el elemento actual de la lista");
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
            this.btnEdit.Location = new System.Drawing.Point(39, 6);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(28, 28);
            this.btnEdit.TabIndex = 43;
            this.toolTipNIF.SetToolTip(this.btnEdit, "Muestra un cuadro de diálogo para editar el elemento actual de la lista");
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.Location = new System.Drawing.Point(5, 6);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(28, 28);
            this.btnAdd.TabIndex = 42;
            this.toolTipNIF.SetToolTip(this.btnAdd, "Agrega un elemento nuevo a la lista y muestra un cuadro diálogo para editarlo");
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmdExcel.Image")));
            this.cmdExcel.Location = new System.Drawing.Point(125, 6);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(28, 28);
            this.cmdExcel.TabIndex = 41;
            this.toolTipNIF.SetToolTip(this.cmdExcel, "Exportar a Excel");
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // cmdSearch
            // 
            this.cmdSearch.Image = ((System.Drawing.Image)(resources.GetObject("cmdSearch.Image")));
            this.cmdSearch.Location = new System.Drawing.Point(206, 40);
            this.cmdSearch.Name = "cmdSearch";
            this.cmdSearch.Size = new System.Drawing.Size(62, 39);
            this.cmdSearch.TabIndex = 14;
            this.toolTipCUPSREE.SetToolTip(this.cmdSearch, "Buscar datos");
            this.cmdSearch.UseVisualStyleBackColor = true;
            this.cmdSearch.Click += new System.EventHandler(this.cmdSearch_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 142);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(947, 357);
            this.tabControl1.TabIndex = 41;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnDel);
            this.tabPage2.Controls.Add(this.btnEdit);
            this.tabPage2.Controls.Add(this.btnAdd);
            this.tabPage2.Controls.Add(this.cmdExcel);
            this.tabPage2.Controls.Add(this.dgvFacturas);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(939, 331);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Puntos Sofisticados Manuales";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // FrmPuntosSofisticados
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(971, 511);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.cmdSearch);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtCCOUNIPS);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmPuntosSofisticados";
            this.Text = "Puntos Sofisticados";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmPuntosSofisticados_FormClosing);
            this.Load += new System.EventHandler(this.FrmFacturasOperaciones_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFacturas)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtCCOUNIPS;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button cmdSearch;
        private System.Windows.Forms.DataGridView dgvFacturas;
        private System.Windows.Forms.ToolTip toolTipNIF;
        private System.Windows.Forms.ToolTip toolTipCUPSREE;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.DataGridViewTextBoxColumn CCOUNIPS;
        private System.Windows.Forms.DataGridViewTextBoxColumn cusp20;
        private System.Windows.Forms.DataGridViewTextBoxColumn codintegr;
        private System.Windows.Forms.DataGridViewTextBoxColumn GRUPO;
        private System.Windows.Forms.DataGridViewTextBoxColumn FD;
        private System.Windows.Forms.DataGridViewTextBoxColumn FH;
        private System.Windows.Forms.DataGridViewTextBoxColumn PRECIOS;
        private System.Windows.Forms.DataGridViewCheckBoxColumn FACTURAS_A_CUENTA;
    }
}