namespace GestionOperaciones.forms.contratacion
{
    partial class FrmMotivosRechazo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMotivosRechazo));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.rtxt_Observaciones = new System.Windows.Forms.RichTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmb_Motivos = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_RechazoPdte = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_TipoSol = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_FecRechazoSol = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_NumSolAtr = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_Cliente_Actualizado = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_CUPS = new System.Windows.Forms.TextBox();
            this.lbl_total_registros = new System.Windows.Forms.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cups = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clienteActualizado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.numSolAtr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecRechazoSol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tipoSolicitud = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.rechadoPdte = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.motivos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.observaciones = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1634, 36);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(88, 30);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_dialog_close_29299;
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(212, 34);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(133, 30);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(79, 30);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(18, 62);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1598, 895);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btnCancel);
            this.tabPage1.Controls.Add(this.btnOK);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.rtxt_Observaciones);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.cmb_Motivos);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.txt_RechazoPdte);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.txt_TipoSol);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.txt_FecRechazoSol);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.txt_NumSolAtr);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txt_Cliente_Actualizado);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.txt_CUPS);
            this.tabPage1.Controls.Add(this.lbl_total_registros);
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Size = new System.Drawing.Size(1590, 862);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Rechazos";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(1134, 922);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 35);
            this.btnCancel.TabIndex = 80;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Enabled = false;
            this.btnOK.Location = new System.Drawing.Point(1256, 922);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(112, 35);
            this.btnOK.TabIndex = 79;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click_1);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(202, 706);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(118, 20);
            this.label8.TabIndex = 78;
            this.label8.Text = "Observaciones:";
            // 
            // rtxt_Observaciones
            // 
            this.rtxt_Observaciones.Location = new System.Drawing.Point(333, 702);
            this.rtxt_Observaciones.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rtxt_Observaciones.Name = "rtxt_Observaciones";
            this.rtxt_Observaciones.Size = new System.Drawing.Size(1033, 186);
            this.rtxt_Observaciones.TabIndex = 77;
            this.rtxt_Observaciones.Text = "";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(254, 665);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(67, 20);
            this.label7.TabIndex = 76;
            this.label7.Text = "Motivos:";
            // 
            // cmb_Motivos
            // 
            this.cmb_Motivos.FormattingEnabled = true;
            this.cmb_Motivos.Location = new System.Drawing.Point(333, 660);
            this.cmb_Motivos.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmb_Motivos.Name = "cmb_Motivos";
            this.cmb_Motivos.Size = new System.Drawing.Size(1033, 28);
            this.cmb_Motivos.TabIndex = 75;
            this.cmb_Motivos.SelectedIndexChanged += new System.EventHandler(this.cmb_Motivos_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(670, 614);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(114, 20);
            this.label5.TabIndex = 74;
            this.label5.Text = "Rechazo Pdte:";
            // 
            // txt_RechazoPdte
            // 
            this.txt_RechazoPdte.Location = new System.Drawing.Point(796, 609);
            this.txt_RechazoPdte.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_RechazoPdte.Name = "txt_RechazoPdte";
            this.txt_RechazoPdte.Size = new System.Drawing.Size(410, 26);
            this.txt_RechazoPdte.TabIndex = 73;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(213, 614);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(107, 20);
            this.label6.TabIndex = 72;
            this.label6.Text = "Tipo Solicitud:";
            // 
            // txt_TipoSol
            // 
            this.txt_TipoSol.Location = new System.Drawing.Point(333, 609);
            this.txt_TipoSol.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_TipoSol.Name = "txt_TipoSol";
            this.txt_TipoSol.Size = new System.Drawing.Size(211, 26);
            this.txt_TipoSol.TabIndex = 71;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(594, 574);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(190, 20);
            this.label4.TabIndex = 70;
            this.label4.Text = "Fecha Rechazo Solicitud:";
            // 
            // txt_FecRechazoSol
            // 
            this.txt_FecRechazoSol.Location = new System.Drawing.Point(796, 569);
            this.txt_FecRechazoSol.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_FecRechazoSol.Name = "txt_FecRechazoSol";
            this.txt_FecRechazoSol.Size = new System.Drawing.Size(154, 26);
            this.txt_FecRechazoSol.TabIndex = 69;
            this.txt_FecRechazoSol.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(234, 574);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(90, 20);
            this.label3.TabIndex = 68;
            this.label3.Text = "NumSolAtr:";
            // 
            // txt_NumSolAtr
            // 
            this.txt_NumSolAtr.Location = new System.Drawing.Point(333, 569);
            this.txt_NumSolAtr.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_NumSolAtr.Name = "txt_NumSolAtr";
            this.txt_NumSolAtr.Size = new System.Drawing.Size(211, 26);
            this.txt_NumSolAtr.TabIndex = 67;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(638, 538);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(149, 20);
            this.label2.TabIndex = 66;
            this.label2.Text = "Cliente Actualizado:";
            // 
            // txt_Cliente_Actualizado
            // 
            this.txt_Cliente_Actualizado.Location = new System.Drawing.Point(796, 529);
            this.txt_Cliente_Actualizado.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Cliente_Actualizado.Name = "txt_Cliente_Actualizado";
            this.txt_Cliente_Actualizado.Size = new System.Drawing.Size(410, 26);
            this.txt_Cliente_Actualizado.TabIndex = 65;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(266, 534);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 20);
            this.label1.TabIndex = 64;
            this.label1.Text = "CUPS:";
            // 
            // txt_CUPS
            // 
            this.txt_CUPS.Location = new System.Drawing.Point(333, 529);
            this.txt_CUPS.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_CUPS.Name = "txt_CUPS";
            this.txt_CUPS.Size = new System.Drawing.Size(211, 26);
            this.txt_CUPS.TabIndex = 63;
            // 
            // lbl_total_registros
            // 
            this.lbl_total_registros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total_registros.AutoSize = true;
            this.lbl_total_registros.Location = new System.Drawing.Point(882, 32);
            this.lbl_total_registros.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbl_total_registros.Name = "lbl_total_registros";
            this.lbl_total_registros.Size = new System.Drawing.Size(128, 20);
            this.lbl_total_registros.TabIndex = 62;
            this.lbl_total_registros.Text = "Nº de Rechazos:";
            this.lbl_total_registros.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cups,
            this.clienteActualizado,
            this.numSolAtr,
            this.fecRechazoSol,
            this.tipoSolicitud,
            this.rechadoPdte,
            this.motivos,
            this.observaciones});
            this.dgv.Location = new System.Drawing.Point(9, 71);
            this.dgv.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.RowHeadersWidth = 62;
            this.dgv.Size = new System.Drawing.Size(1040, 429);
            this.dgv.TabIndex = 0;
            this.dgv.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellClick);
            this.dgv.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgv_KeyDown);
            this.dgv.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgv_KeyUp);
            // 
            // cups
            // 
            this.cups.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups.DataPropertyName = "cups";
            this.cups.HeaderText = "CUPS";
            this.cups.MinimumWidth = 8;
            this.cups.Name = "cups";
            this.cups.ReadOnly = true;
            this.cups.Width = 89;
            // 
            // clienteActualizado
            // 
            this.clienteActualizado.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.clienteActualizado.DataPropertyName = "clienteActualizado";
            this.clienteActualizado.HeaderText = "Cliente Actualizado";
            this.clienteActualizado.MinimumWidth = 8;
            this.clienteActualizado.Name = "clienteActualizado";
            this.clienteActualizado.ReadOnly = true;
            this.clienteActualizado.Width = 166;
            // 
            // numSolAtr
            // 
            this.numSolAtr.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.numSolAtr.DataPropertyName = "numSolAtr";
            this.numSolAtr.HeaderText = "NumSolAtr";
            this.numSolAtr.MinimumWidth = 8;
            this.numSolAtr.Name = "numSolAtr";
            this.numSolAtr.ReadOnly = true;
            this.numSolAtr.Width = 122;
            // 
            // fecRechazoSol
            // 
            this.fecRechazoSol.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecRechazoSol.DataPropertyName = "fechaRechazoSol";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fecRechazoSol.DefaultCellStyle = dataGridViewCellStyle2;
            this.fecRechazoSol.HeaderText = "FecRechazoSol";
            this.fecRechazoSol.MinimumWidth = 8;
            this.fecRechazoSol.Name = "fecRechazoSol";
            this.fecRechazoSol.ReadOnly = true;
            this.fecRechazoSol.Width = 159;
            // 
            // tipoSolicitud
            // 
            this.tipoSolicitud.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tipoSolicitud.DataPropertyName = "tipoSolicitud";
            this.tipoSolicitud.HeaderText = "Tipo Solicitud";
            this.tipoSolicitud.MinimumWidth = 8;
            this.tipoSolicitud.Name = "tipoSolicitud";
            this.tipoSolicitud.ReadOnly = true;
            this.tipoSolicitud.Width = 128;
            // 
            // rechadoPdte
            // 
            this.rechadoPdte.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.rechadoPdte.DataPropertyName = "rechazoPdte";
            this.rechadoPdte.HeaderText = "RechadoPdte";
            this.rechadoPdte.MinimumWidth = 8;
            this.rechadoPdte.Name = "rechadoPdte";
            this.rechadoPdte.ReadOnly = true;
            this.rechadoPdte.Width = 143;
            // 
            // motivos
            // 
            this.motivos.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.motivos.DataPropertyName = "motivos";
            this.motivos.HeaderText = "Motivos";
            this.motivos.MinimumWidth = 8;
            this.motivos.Name = "motivos";
            this.motivos.ReadOnly = true;
            this.motivos.Width = 99;
            // 
            // observaciones
            // 
            this.observaciones.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.observaciones.DataPropertyName = "comentario";
            this.observaciones.HeaderText = "Observaciones";
            this.observaciones.MinimumWidth = 8;
            this.observaciones.Name = "observaciones";
            this.observaciones.ReadOnly = true;
            this.observaciones.Width = 150;
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // FrmMotivosRechazo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1634, 1050);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FrmMotivosRechazo";
            this.Text = "Motivos Rechazo";
            this.Load += new System.EventHandler(this.FrmMotivosRechazo_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lbl_total_registros;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups;
        private System.Windows.Forms.DataGridViewTextBoxColumn clienteActualizado;
        private System.Windows.Forms.DataGridViewTextBoxColumn numSolAtr;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecRechazoSol;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipoSolicitud;
        private System.Windows.Forms.DataGridViewTextBoxColumn rechadoPdte;
        private System.Windows.Forms.DataGridViewTextBoxColumn motivos;
        private System.Windows.Forms.DataGridViewTextBoxColumn observaciones;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.RichTextBox rtxt_Observaciones;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmb_Motivos;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_RechazoPdte;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_TipoSol;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_FecRechazoSol;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_NumSolAtr;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_Cliente_Actualizado;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_CUPS;
    }
}