namespace GestionOperaciones.forms.contratacion
{
    partial class FrmDisuasoria
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmDisuasoria));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuCopyPaste = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuCopy = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.salirMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.C70CCNAE_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.paramétricasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excComplToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cNAEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMaxFecha = new System.Windows.Forms.TextBox();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgvResumen = new System.Windows.Forms.DataGridView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.btnExcel = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.NUM_CUPS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.R_ESENCIAL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.R_TIPO_EMPRESA_AAPP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gestor = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sin_segmento = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.total_general = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuCopyPaste.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResumen)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuCopyPaste
            // 
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
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirMenuItem,
            this.importarToolStripMenuItem,
            this.paramétricasToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(904, 24);
            this.menuStrip1.TabIndex = 29;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // salirMenuItem
            // 
            this.salirMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.salirMenuItem.Name = "salirMenuItem";
            this.salirMenuItem.Size = new System.Drawing.Size(60, 20);
            this.salirMenuItem.Text = "Archivo";
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
            this.C70CCNAE_ToolStripMenuItem});
            this.importarToolStripMenuItem.Name = "importarToolStripMenuItem";
            this.importarToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.importarToolStripMenuItem.Text = "Importar";
            // 
            // C70CCNAE_ToolStripMenuItem
            // 
            this.C70CCNAE_ToolStripMenuItem.Name = "C70CCNAE_ToolStripMenuItem";
            this.C70CCNAE_ToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.C70CCNAE_ToolStripMenuItem.Text = "Extracción C70CCNAE";
            this.C70CCNAE_ToolStripMenuItem.Click += new System.EventHandler(this.C70CCNAE_ToolStripMenuItem_Click);
            // 
            // paramétricasToolStripMenuItem
            // 
            this.paramétricasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.excComplToolStripMenuItem,
            this.cNAEToolStripMenuItem,
            this.toolStripSeparator1,
            this.opcionesToolStripMenuItem});
            this.paramétricasToolStripMenuItem.Name = "paramétricasToolStripMenuItem";
            this.paramétricasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.paramétricasToolStripMenuItem.Text = "Herramientas";
            // 
            // excComplToolStripMenuItem
            // 
            this.excComplToolStripMenuItem.Name = "excComplToolStripMenuItem";
            this.excComplToolStripMenuItem.Size = new System.Drawing.Size(223, 22);
            this.excComplToolStripMenuItem.Text = "Exclusión de complementos";
            this.excComplToolStripMenuItem.Click += new System.EventHandler(this.excComplToolStripMenuItem_Click);
            // 
            // cNAEToolStripMenuItem
            // 
            this.cNAEToolStripMenuItem.Name = "cNAEToolStripMenuItem";
            this.cNAEToolStripMenuItem.Size = new System.Drawing.Size(223, 22);
            this.cNAEToolStripMenuItem.Text = "CNAE";
            this.cNAEToolStripMenuItem.Click += new System.EventHandler(this.cNAEToolStripMenuItem_Click);
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
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(552, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(206, 13);
            this.label2.TabIndex = 33;
            this.label2.Text = "Última importación del archivo C70CCNAE";
            // 
            // txtMaxFecha
            // 
            this.txtMaxFecha.Enabled = false;
            this.txtMaxFecha.Location = new System.Drawing.Point(764, 38);
            this.txtMaxFecha.Name = "txtMaxFecha";
            this.txtMaxFecha.Size = new System.Drawing.Size(124, 20);
            this.txtMaxFecha.TabIndex = 34;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgvResumen);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(869, 302);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Resumen";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgvResumen
            // 
            this.dgvResumen.AllowUserToAddRows = false;
            this.dgvResumen.AllowUserToDeleteRows = false;
            this.dgvResumen.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.SandyBrown;
            this.dgvResumen.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvResumen.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvResumen.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvResumen.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.NUM_CUPS,
            this.R_ESENCIAL,
            this.R_TIPO_EMPRESA_AAPP,
            this.gestor,
            this.sin_segmento,
            this.total_general});
            this.dgvResumen.Location = new System.Drawing.Point(6, 6);
            this.dgvResumen.Name = "dgvResumen";
            this.dgvResumen.ReadOnly = true;
            this.dgvResumen.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvResumen.Size = new System.Drawing.Size(857, 290);
            this.dgvResumen.TabIndex = 26;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(15, 131);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(877, 328);
            this.tabControl1.TabIndex = 32;
            // 
            // btnExcel
            // 
            this.btnExcel.Image = global::GestionOperaciones.Properties.Resources._1493038278_excel;
            this.btnExcel.Location = new System.Drawing.Point(12, 77);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(38, 39);
            this.btnExcel.TabIndex = 26;
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(267, 67);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(615, 80);
            this.richTextBox1.TabIndex = 35;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // NUM_CUPS
            // 
            this.NUM_CUPS.DataPropertyName = "etiqueta";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.NUM_CUPS.DefaultCellStyle = dataGridViewCellStyle2;
            this.NUM_CUPS.HeaderText = "Etiquetas de Fila";
            this.NUM_CUPS.Name = "NUM_CUPS";
            this.NUM_CUPS.ReadOnly = true;
            // 
            // R_ESENCIAL
            // 
            this.R_ESENCIAL.DataPropertyName = "aapp";
            this.R_ESENCIAL.HeaderText = "AAPP";
            this.R_ESENCIAL.Name = "R_ESENCIAL";
            this.R_ESENCIAL.ReadOnly = true;
            this.R_ESENCIAL.ToolTipText = "Administraciones Públicas";
            // 
            // R_TIPO_EMPRESA_AAPP
            // 
            this.R_TIPO_EMPRESA_AAPP.DataPropertyName = "kam";
            this.R_TIPO_EMPRESA_AAPP.HeaderText = "KAM";
            this.R_TIPO_EMPRESA_AAPP.Name = "R_TIPO_EMPRESA_AAPP";
            this.R_TIPO_EMPRESA_AAPP.ReadOnly = true;
            // 
            // gestor
            // 
            this.gestor.DataPropertyName = "gestor";
            this.gestor.HeaderText = "GESTOR";
            this.gestor.Name = "gestor";
            this.gestor.ReadOnly = true;
            // 
            // sin_segmento
            // 
            this.sin_segmento.DataPropertyName = "sin_segmento";
            this.sin_segmento.HeaderText = "Sin segmento";
            this.sin_segmento.Name = "sin_segmento";
            this.sin_segmento.ReadOnly = true;
            // 
            // total_general
            // 
            this.total_general.DataPropertyName = "total_general";
            this.total_general.HeaderText = "Total General";
            this.total_general.Name = "total_general";
            this.total_general.ReadOnly = true;
            // 
            // FrmDisuasoria
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(904, 471);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.txtMaxFecha);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.btnExcel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmDisuasoria";
            this.Text = "Tarifa disuasoria";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmDisuasoria_FormClosing);
            this.Load += new System.EventHandler(this.FrmDisuasoria_Load);
            this.menuCopyPaste.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvResumen)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.ContextMenuStrip menuCopyPaste;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuCopy;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem salirMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem C70CCNAE_ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem paramétricasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excComplToolStripMenuItem;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMaxFecha;
        private System.Windows.Forms.ToolStripMenuItem cNAEToolStripMenuItem;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dgvResumen;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn NUM_CUPS;
        private System.Windows.Forms.DataGridViewTextBoxColumn R_ESENCIAL;
        private System.Windows.Forms.DataGridViewTextBoxColumn R_TIPO_EMPRESA_AAPP;
        private System.Windows.Forms.DataGridViewTextBoxColumn gestor;
        private System.Windows.Forms.DataGridViewTextBoxColumn sin_segmento;
        private System.Windows.Forms.DataGridViewTextBoxColumn total_general;
    }
}