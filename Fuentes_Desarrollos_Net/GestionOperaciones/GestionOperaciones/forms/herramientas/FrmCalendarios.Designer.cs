namespace GestionOperaciones.forms.herramientas
{
    partial class FrmCalendarios
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCalendarios));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdb_cuartohoraria = new System.Windows.Forms.RadioButton();
            this.rdb_horaria = new System.Windows.Forms.RadioButton();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbTerritorios = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbTarifas = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_fh = new System.Windows.Forms.DateTimePicker();
            this.txt_fd = new System.Windows.Forms.DateTimePicker();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdb_cuartohoraria);
            this.groupBox1.Controls.Add(this.rdb_horaria);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cmbTerritorios);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.cmbTarifas);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txt_fh);
            this.groupBox1.Controls.Add(this.txt_fd);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(484, 187);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // rdb_cuartohoraria
            // 
            this.rdb_cuartohoraria.AutoSize = true;
            this.rdb_cuartohoraria.Location = new System.Drawing.Point(32, 122);
            this.rdb_cuartohoraria.Name = "rdb_cuartohoraria";
            this.rdb_cuartohoraria.Size = new System.Drawing.Size(90, 17);
            this.rdb_cuartohoraria.TabIndex = 15;
            this.rdb_cuartohoraria.TabStop = true;
            this.rdb_cuartohoraria.Text = "CuartoHoraria";
            this.rdb_cuartohoraria.UseVisualStyleBackColor = true;
            this.rdb_cuartohoraria.CheckedChanged += new System.EventHandler(this.rdb_cuartohoraria_CheckedChanged);
            // 
            // rdb_horaria
            // 
            this.rdb_horaria.AutoSize = true;
            this.rdb_horaria.Location = new System.Drawing.Point(32, 99);
            this.rdb_horaria.Name = "rdb_horaria";
            this.rdb_horaria.Size = new System.Drawing.Size(59, 17);
            this.rdb_horaria.TabIndex = 14;
            this.rdb_horaria.TabStop = true;
            this.rdb_horaria.Text = "Horaria";
            this.rdb_horaria.UseVisualStyleBackColor = true;
            this.rdb_horaria.CheckedChanged += new System.EventHandler(this.rdb_horaria_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(275, 59);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Territorio";
            // 
            // cmbTerritorios
            // 
            this.cmbTerritorios.FormattingEnabled = true;
            this.cmbTerritorios.Location = new System.Drawing.Point(329, 56);
            this.cmbTerritorios.Name = "cmbTerritorios";
            this.cmbTerritorios.Size = new System.Drawing.Size(96, 21);
            this.cmbTerritorios.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(94, 59);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Tarifa";
            // 
            // cmbTarifas
            // 
            this.cmbTarifas.FormattingEnabled = true;
            this.cmbTarifas.Location = new System.Drawing.Point(134, 56);
            this.cmbTarifas.Name = "cmbTarifas";
            this.cmbTarifas.Size = new System.Drawing.Size(96, 21);
            this.cmbTarifas.TabIndex = 10;
            this.cmbTarifas.SelectedIndexChanged += new System.EventHandler(this.cmbTarifas_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(257, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Fecha hasta";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(59, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Fecha desde";
            // 
            // txt_fh
            // 
            this.txt_fh.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fh.Location = new System.Drawing.Point(329, 27);
            this.txt_fh.Name = "txt_fh";
            this.txt_fh.Size = new System.Drawing.Size(96, 20);
            this.txt_fh.TabIndex = 7;
            // 
            // txt_fd
            // 
            this.txt_fd.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fd.Location = new System.Drawing.Point(134, 27);
            this.txt_fd.Name = "txt_fd";
            this.txt_fd.Size = new System.Drawing.Size(96, 20);
            this.txt_fd.TabIndex = 6;
            // 
            // cmdExcel
            // 
            this.cmdExcel.Image = global::GestionOperaciones.Properties.Resources.if_microsoft_office_excel_1784856;
            this.cmdExcel.Location = new System.Drawing.Point(550, 84);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(52, 48);
            this.cmdExcel.TabIndex = 35;
            this.cmdExcel.UseVisualStyleBackColor = true;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(534, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(84, 13);
            this.label5.TabIndex = 36;
            this.label5.Text = "Exportar a Excel";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(628, 24);
            this.menuStrip1.TabIndex = 37;
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
            // FrmCalendarios
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(628, 226);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cmdExcel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmCalendarios";
            this.Text = "Calendarios Tarifarios";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmCalendarios_FormClosing);
            this.Load += new System.EventHandler(this.FrmCalendarios_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker txt_fh;
        private System.Windows.Forms.DateTimePicker txt_fd;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbTerritorios;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbTarifas;
        private System.Windows.Forms.RadioButton rdb_cuartohoraria;
        private System.Windows.Forms.RadioButton rdb_horaria;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
    }
}