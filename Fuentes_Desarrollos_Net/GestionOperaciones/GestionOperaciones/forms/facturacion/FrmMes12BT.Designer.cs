namespace GestionOperaciones.forms.facturacion
{
    partial class FrmMes12BT
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMes12BT));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarExcelMes12ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.importarAdjudicaciónToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarPrevisiónToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_sol_desde = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btn_lanzar_proceso = new System.Windows.Forms.Button();
            this.txtInfo = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.lbl_total_agrupadas = new System.Windows.Forms.Label();
            this.lbl_total_individuales = new System.Windows.Forms.Label();
            this.lbl_factoring_adjudicacion = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(909, 24);
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
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarExcelMes12ToolStripMenuItem,
            this.toolStripSeparator1,
            this.importarAdjudicaciónToolStripMenuItem,
            this.importarPrevisiónToolStripMenuItem,
            this.toolStripSeparator2,
            this.opcionesToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarExcelMes12ToolStripMenuItem
            // 
            this.importarExcelMes12ToolStripMenuItem.Name = "importarExcelMes12ToolStripMenuItem";
            this.importarExcelMes12ToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.importarExcelMes12ToolStripMenuItem.Text = "Importar Excel Mes 12";
            this.importarExcelMes12ToolStripMenuItem.Click += new System.EventHandler(this.importarExcelMes12ToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(188, 6);
            // 
            // importarAdjudicaciónToolStripMenuItem
            // 
            this.importarAdjudicaciónToolStripMenuItem.Name = "importarAdjudicaciónToolStripMenuItem";
            this.importarAdjudicaciónToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.importarAdjudicaciónToolStripMenuItem.Text = "Importar adjudicación";
            this.importarAdjudicaciónToolStripMenuItem.Click += new System.EventHandler(this.importarAdjudicaciónToolStripMenuItem_Click);
            // 
            // importarPrevisiónToolStripMenuItem
            // 
            this.importarPrevisiónToolStripMenuItem.Enabled = false;
            this.importarPrevisiónToolStripMenuItem.Name = "importarPrevisiónToolStripMenuItem";
            this.importarPrevisiónToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.importarPrevisiónToolStripMenuItem.Text = "Importar Previsión";
            this.importarPrevisiónToolStripMenuItem.Click += new System.EventHandler(this.importarPrevisiónToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(188, 6);
            // 
            // opcionesToolStripMenuItem
            // 
            this.opcionesToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_settings_2199094;
            this.opcionesToolStripMenuItem.Name = "opcionesToolStripMenuItem";
            this.opcionesToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.opcionesToolStripMenuItem.Text = "Opciones";
            this.opcionesToolStripMenuItem.Click += new System.EventHandler(this.opcionesToolStripMenuItem_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(10, 255);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(351, 267);
            this.tabControl1.TabIndex = 73;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgv);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(343, 241);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Datos Excel Mes12 BT";
            this.tabPage2.UseVisualStyleBackColor = true;
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
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cups20,
            this.fecha_sol_desde});
            this.dgv.Location = new System.Drawing.Point(6, 6);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.RowHeadersWidth = 51;
            this.dgv.Size = new System.Drawing.Size(331, 233);
            this.dgv.TabIndex = 60;
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "hoja";
            this.cups20.HeaderText = "HOJA";
            this.cups20.MinimumWidth = 6;
            this.cups20.Name = "cups20";
            this.cups20.ReadOnly = true;
            this.cups20.Width = 60;
            // 
            // fecha_sol_desde
            // 
            this.fecha_sol_desde.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_sol_desde.DataPropertyName = "registros";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            this.fecha_sol_desde.DefaultCellStyle = dataGridViewCellStyle2;
            this.fecha_sol_desde.HeaderText = "Total registros";
            this.fecha_sol_desde.MinimumWidth = 6;
            this.fecha_sol_desde.Name = "fecha_sol_desde";
            this.fecha_sol_desde.ReadOnly = true;
            this.fecha_sol_desde.Width = 98;
            // 
            // btn_lanzar_proceso
            // 
            this.btn_lanzar_proceso.Location = new System.Drawing.Point(546, 75);
            this.btn_lanzar_proceso.Name = "btn_lanzar_proceso";
            this.btn_lanzar_proceso.Size = new System.Drawing.Size(159, 61);
            this.btn_lanzar_proceso.TabIndex = 76;
            this.btn_lanzar_proceso.Text = "Lanzar Proceso";
            this.btn_lanzar_proceso.UseVisualStyleBackColor = true;
            this.btn_lanzar_proceso.Click += new System.EventHandler(this.btn_lanzar_proceso_Click);
            // 
            // txtInfo
            // 
            this.txtInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtInfo.Location = new System.Drawing.Point(366, 283);
            this.txtInfo.Margin = new System.Windows.Forms.Padding(2);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.Size = new System.Drawing.Size(374, 238);
            this.txtInfo.TabIndex = 77;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(10, 27);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(458, 149);
            this.textBox1.TabIndex = 89;
            this.textBox1.Text = resources.GetString("textBox1.Text");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(367, 255);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 79;
            this.label1.Text = "Log proceso";
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // lbl_total_agrupadas
            // 
            this.lbl_total_agrupadas.AutoSize = true;
            this.lbl_total_agrupadas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_total_agrupadas.ForeColor = System.Drawing.Color.Black;
            this.lbl_total_agrupadas.Location = new System.Drawing.Point(294, 222);
            this.lbl_total_agrupadas.Name = "lbl_total_agrupadas";
            this.lbl_total_agrupadas.Size = new System.Drawing.Size(174, 13);
            this.lbl_total_agrupadas.TabIndex = 85;
            this.lbl_total_agrupadas.Text = "Nº Adjudicaciones agrupadas";
            // 
            // lbl_total_individuales
            // 
            this.lbl_total_individuales.AutoSize = true;
            this.lbl_total_individuales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_total_individuales.ForeColor = System.Drawing.Color.Black;
            this.lbl_total_individuales.Location = new System.Drawing.Point(14, 222);
            this.lbl_total_individuales.Name = "lbl_total_individuales";
            this.lbl_total_individuales.Size = new System.Drawing.Size(182, 13);
            this.lbl_total_individuales.TabIndex = 84;
            this.lbl_total_individuales.Text = "Nº Adjudicaciones individuales";
            // 
            // lbl_factoring_adjudicacion
            // 
            this.lbl_factoring_adjudicacion.AutoSize = true;
            this.lbl_factoring_adjudicacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_factoring_adjudicacion.ForeColor = System.Drawing.Color.Black;
            this.lbl_factoring_adjudicacion.Location = new System.Drawing.Point(17, 193);
            this.lbl_factoring_adjudicacion.Name = "lbl_factoring_adjudicacion";
            this.lbl_factoring_adjudicacion.Size = new System.Drawing.Size(136, 13);
            this.lbl_factoring_adjudicacion.TabIndex = 88;
            this.lbl_factoring_adjudicacion.Text = "Factoring adjudicacion";
            // 
            // FrmMes12BT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(909, 530);
            this.Controls.Add(this.lbl_factoring_adjudicacion);
            this.Controls.Add(this.lbl_total_agrupadas);
            this.Controls.Add(this.lbl_total_individuales);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.txtInfo);
            this.Controls.Add(this.btn_lanzar_proceso);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmMes12BT";
            this.Text = "MES 12 BT";
            this.Load += new System.EventHandler(this.FrmMes12BT_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarExcelMes12ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_sol_desde;
        private System.Windows.Forms.ToolStripMenuItem importarPrevisiónToolStripMenuItem;
        private System.Windows.Forms.Button btn_lanzar_proceso;
        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.ToolStripMenuItem importarAdjudicaciónToolStripMenuItem;
        private System.Windows.Forms.Label lbl_factoring_adjudicacion;
        private System.Windows.Forms.Label lbl_total_agrupadas;
        private System.Windows.Forms.Label lbl_total_individuales;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
    }
}