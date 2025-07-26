namespace GestionOperaciones.forms.facturacion
{
    partial class FrmMes13Conversion
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMes13Conversion));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarAdjudicaciónToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lbl_total_facturas_mes12 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_fd_agrupadas = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_fh_agrupadas = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_fh_consumo = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_fd_consumo = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFFACTURAHAS = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.txtFFACTURADES = new System.Windows.Forms.DateTimePicker();
            this.txt_factoring = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chk_solo_excel = new System.Windows.Forms.CheckBox();
            this.btn_genera_excel = new System.Windows.Forms.Button();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lbl_total_individuales = new System.Windows.Forms.Label();
            this.lbl_total_agrupadas = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(798, 24);
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
            this.importarAdjudicaciónToolStripMenuItem,
            this.toolStripSeparator1,
            this.opcionesToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarAdjudicaciónToolStripMenuItem
            // 
            this.importarAdjudicaciónToolStripMenuItem.Name = "importarAdjudicaciónToolStripMenuItem";
            this.importarAdjudicaciónToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.importarAdjudicaciónToolStripMenuItem.Text = "Importar adjudicación";
            this.importarAdjudicaciónToolStripMenuItem.Click += new System.EventHandler(this.importarAdjudicaciónToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbl_total_agrupadas);
            this.groupBox1.Controls.Add(this.lbl_total_individuales);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.lbl_total_facturas_mes12);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txt_fd_agrupadas);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txt_fh_agrupadas);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txt_fh_consumo);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txt_fd_consumo);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txtFFACTURAHAS);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtFFACTURADES);
            this.groupBox1.Controls.Add(this.txt_factoring);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(570, 268);
            this.groupBox1.TabIndex = 40;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(192, 116);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(336, 72);
            this.textBox1.TabIndex = 66;
            this.textBox1.Text = "Importante!!!\r\nÚnicamente rellenar con valores numéricos y en formato AAAAMM,\r\nya" +
    " que este valor se utilizará para guardar la importación de la adjudicación.";
            // 
            // lbl_total_facturas_mes12
            // 
            this.lbl_total_facturas_mes12.AutoSize = true;
            this.lbl_total_facturas_mes12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_total_facturas_mes12.ForeColor = System.Drawing.Color.Black;
            this.lbl_total_facturas_mes12.Location = new System.Drawing.Point(20, 234);
            this.lbl_total_facturas_mes12.Name = "lbl_total_facturas_mes12";
            this.lbl_total_facturas_mes12.Size = new System.Drawing.Size(130, 13);
            this.lbl_total_facturas_mes12.TabIndex = 50;
            this.lbl_total_facturas_mes12.Text = "Total Facturas Mes12";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.ForeColor = System.Drawing.Color.Black;
            this.label8.Location = new System.Drawing.Point(25, 80);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(164, 13);
            this.label8.TabIndex = 49;
            this.label8.Text = "Desde Fecha Factura Agrupadas";
            // 
            // txt_fd_agrupadas
            // 
            this.txt_fd_agrupadas.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fd_agrupadas.Location = new System.Drawing.Point(192, 78);
            this.txt_fd_agrupadas.Name = "txt_fd_agrupadas";
            this.txt_fd_agrupadas.Size = new System.Drawing.Size(80, 20);
            this.txt_fd_agrupadas.TabIndex = 48;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(300, 80);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(161, 13);
            this.label6.TabIndex = 47;
            this.label6.Text = "Hasta Fecha Factura Agrupadas";
            // 
            // txt_fh_agrupadas
            // 
            this.txt_fh_agrupadas.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fh_agrupadas.Location = new System.Drawing.Point(467, 78);
            this.txt_fh_agrupadas.Name = "txt_fh_agrupadas";
            this.txt_fh_agrupadas.Size = new System.Drawing.Size(80, 20);
            this.txt_fh_agrupadas.TabIndex = 46;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(235, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 13);
            this.label4.TabIndex = 45;
            this.label4.Text = "Hasta Fecha Consumo";
            // 
            // txt_fh_consumo
            // 
            this.txt_fh_consumo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fh_consumo.Location = new System.Drawing.Point(356, 53);
            this.txt_fh_consumo.Name = "txt_fh_consumo";
            this.txt_fh_consumo.Size = new System.Drawing.Size(80, 20);
            this.txt_fh_consumo.TabIndex = 44;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(20, 54);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(118, 13);
            this.label5.TabIndex = 43;
            this.label5.Text = "Desde Fecha Consumo";
            // 
            // txt_fd_consumo
            // 
            this.txt_fd_consumo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fd_consumo.Location = new System.Drawing.Point(144, 54);
            this.txt_fd_consumo.Name = "txt_fd_consumo";
            this.txt_fd_consumo.Size = new System.Drawing.Size(80, 20);
            this.txt_fd_consumo.TabIndex = 42;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(243, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 13);
            this.label2.TabIndex = 41;
            this.label2.Text = "Hasta Fecha Factura";
            // 
            // txtFFACTURAHAS
            // 
            this.txtFFACTURAHAS.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFFACTURAHAS.Location = new System.Drawing.Point(356, 28);
            this.txtFFACTURAHAS.Name = "txtFFACTURAHAS";
            this.txtFFACTURAHAS.Size = new System.Drawing.Size(80, 20);
            this.txtFFACTURAHAS.TabIndex = 40;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(28, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 13);
            this.label3.TabIndex = 39;
            this.label3.Text = "Desde Fecha Factura";
            // 
            // txtFFACTURADES
            // 
            this.txtFFACTURADES.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFFACTURADES.Location = new System.Drawing.Point(144, 28);
            this.txtFFACTURADES.Name = "txtFFACTURADES";
            this.txtFFACTURADES.Size = new System.Drawing.Size(80, 20);
            this.txtFFACTURADES.TabIndex = 38;
            // 
            // txt_factoring
            // 
            this.txt_factoring.Location = new System.Drawing.Point(94, 116);
            this.txt_factoring.Name = "txt_factoring";
            this.txt_factoring.Size = new System.Drawing.Size(80, 20);
            this.txt_factoring.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(28, 119);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Factoring";
            // 
            // chk_solo_excel
            // 
            this.chk_solo_excel.AutoSize = true;
            this.chk_solo_excel.Location = new System.Drawing.Point(605, 60);
            this.chk_solo_excel.Name = "chk_solo_excel";
            this.chk_solo_excel.Size = new System.Drawing.Size(181, 17);
            this.chk_solo_excel.TabIndex = 42;
            this.chk_solo_excel.Text = "Sólo generar Excel (Sin cálculos)";
            this.chk_solo_excel.UseVisualStyleBackColor = true;
            // 
            // btn_genera_excel
            // 
            this.btn_genera_excel.Location = new System.Drawing.Point(636, 99);
            this.btn_genera_excel.Name = "btn_genera_excel";
            this.btn_genera_excel.Size = new System.Drawing.Size(121, 74);
            this.btn_genera_excel.TabIndex = 41;
            this.btn_genera_excel.Text = "Generar Excel Seguimiento";
            this.btn_genera_excel.UseVisualStyleBackColor = true;
            this.btn_genera_excel.Click += new System.EventHandler(this.btn_genera_excel_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(188, 6);
            // 
            // opcionesToolStripMenuItem
            // 
            this.opcionesToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_settings_2199094;
            this.opcionesToolStripMenuItem.Name = "opcionesToolStripMenuItem";
            this.opcionesToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.opcionesToolStripMenuItem.Text = "Opciones";
            this.opcionesToolStripMenuItem.Click += new System.EventHandler(this.opcionesToolStripMenuItem_Click);
            // 
            // lbl_total_individuales
            // 
            this.lbl_total_individuales.AutoSize = true;
            this.lbl_total_individuales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_total_individuales.ForeColor = System.Drawing.Color.Black;
            this.lbl_total_individuales.Location = new System.Drawing.Point(20, 208);
            this.lbl_total_individuales.Name = "lbl_total_individuales";
            this.lbl_total_individuales.Size = new System.Drawing.Size(197, 13);
            this.lbl_total_individuales.TabIndex = 67;
            this.lbl_total_individuales.Text = "Total Adjudicaciones individuales";
            // 
            // lbl_total_agrupadas
            // 
            this.lbl_total_agrupadas.AutoSize = true;
            this.lbl_total_agrupadas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_total_agrupadas.ForeColor = System.Drawing.Color.Black;
            this.lbl_total_agrupadas.Location = new System.Drawing.Point(300, 208);
            this.lbl_total_agrupadas.Name = "lbl_total_agrupadas";
            this.lbl_total_agrupadas.Size = new System.Drawing.Size(189, 13);
            this.lbl_total_agrupadas.TabIndex = 68;
            this.lbl_total_agrupadas.Text = "Total Adjudicaciones agrupadas";
            // 
            // FrmMes13Conversion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(798, 330);
            this.Controls.Add(this.chk_solo_excel);
            this.Controls.Add(this.btn_genera_excel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmMes13Conversion";
            this.Text = "Factoring Mes13 Conversión";
            this.Load += new System.EventHandler(this.FrmMes13Conversion_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarAdjudicaciónToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lbl_total_facturas_mes12;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DateTimePicker txt_fd_agrupadas;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker txt_fh_agrupadas;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker txt_fh_consumo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker txt_fd_consumo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker txtFFACTURAHAS;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker txtFFACTURADES;
        private System.Windows.Forms.TextBox txt_factoring;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chk_solo_excel;
        private System.Windows.Forms.Button btn_genera_excel;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
        private System.Windows.Forms.Label lbl_total_agrupadas;
        private System.Windows.Forms.Label lbl_total_individuales;
    }
}