namespace GestionOperaciones.forms.medida
{
    partial class FrmAdif
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAdif));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.informesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.medidaADIFFueraDeInventarioToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtFD = new System.Windows.Forms.DateTimePicker();
            this.txtFH = new System.Windows.Forms.DateTimePicker();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.btnMail = new System.Windows.Forms.Button();
            this.btnImportar = new System.Windows.Forms.Button();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.informesToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(901, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.ShowShortcutKeys = false;
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // informesToolStripMenuItem
            // 
            this.informesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.medidaADIFFueraDeInventarioToolStripMenuItem,
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem,
            this.toolStripSeparator1,
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem});
            this.informesToolStripMenuItem.Name = "informesToolStripMenuItem";
            this.informesToolStripMenuItem.Size = new System.Drawing.Size(66, 20);
            this.informesToolStripMenuItem.Text = "Informes";
            // 
            // medidaADIFFueraDeInventarioToolStripMenuItem
            // 
            this.medidaADIFFueraDeInventarioToolStripMenuItem.Enabled = false;
            this.medidaADIFFueraDeInventarioToolStripMenuItem.Name = "medidaADIFFueraDeInventarioToolStripMenuItem";
            this.medidaADIFFueraDeInventarioToolStripMenuItem.Size = new System.Drawing.Size(299, 22);
            this.medidaADIFFueraDeInventarioToolStripMenuItem.Text = "Medida ADIF fuera de inventario";
            this.medidaADIFFueraDeInventarioToolStripMenuItem.Click += new System.EventHandler(this.medidaADIFFueraDeInventarioToolStripMenuItem_Click);
            // 
            // medidaHorariaADIFAgrupadaVerticalToolStripMenuItem
            // 
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem.Enabled = false;
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem.Name = "medidaHorariaADIFAgrupadaVerticalToolStripMenuItem";
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem.Size = new System.Drawing.Size(299, 22);
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem.Text = "Medida horaria ADIF agrupada vertical";
            this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem.Click += new System.EventHandler(this.medidaHorariaADIFAgrupadaVerticalToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(149, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(182, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "Importación de Datos";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(429, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(316, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "Enviar correo actualización datos SCE";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(553, 283);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 20);
            this.label3.TabIndex = 8;
            this.label3.Text = "Inventario";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(149, 283);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(182, 20);
            this.label4.TabIndex = 6;
            this.label4.Text = "Exportación de Datos";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(697, 136);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(82, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Desde Fecha";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(700, 166);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(79, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Hasta Fecha";
            // 
            // txtFD
            // 
            this.txtFD.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFD.Location = new System.Drawing.Point(785, 130);
            this.txtFD.Name = "txtFD";
            this.txtFD.Size = new System.Drawing.Size(84, 20);
            this.txtFD.TabIndex = 11;
            // 
            // txtFH
            // 
            this.txtFH.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFH.Location = new System.Drawing.Point(785, 160);
            this.txtFH.Name = "txtFH";
            this.txtFH.Size = new System.Drawing.Size(84, 20);
            this.txtFH.TabIndex = 12;
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(296, 6);
            // 
            // button3
            // 
            this.button3.Image = global::GestionOperaciones.Properties.Resources._1496067747_Checklist_clipboard_inventory_list_report_tasks_todo;
            this.button3.Location = new System.Drawing.Point(522, 329);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(153, 135);
            this.button3.TabIndex = 7;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Image = global::GestionOperaciones.Properties.Resources.Export;
            this.button4.Location = new System.Drawing.Point(164, 329);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(153, 135);
            this.button4.TabIndex = 5;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // btnMail
            // 
            this.btnMail.Image = global::GestionOperaciones.Properties.Resources._1495200266_Outlook;
            this.btnMail.Location = new System.Drawing.Point(522, 105);
            this.btnMail.Name = "btnMail";
            this.btnMail.Size = new System.Drawing.Size(153, 135);
            this.btnMail.TabIndex = 3;
            this.btnMail.UseVisualStyleBackColor = true;
            this.btnMail.Click += new System.EventHandler(this.btnMail_Click);
            // 
            // btnImportar
            // 
            this.btnImportar.Image = global::GestionOperaciones.Properties.Resources.Import;
            this.btnImportar.Location = new System.Drawing.Point(164, 105);
            this.btnImportar.Name = "btnImportar";
            this.btnImportar.Size = new System.Drawing.Size(153, 135);
            this.btnImportar.TabIndex = 1;
            this.btnImportar.UseVisualStyleBackColor = true;
            this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
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
            // medidaHorariaADIFExportaciónExcelToolStripMenuItem
            // 
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources._1493038278_excel;
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem.Name = "medidaHorariaADIFExportaciónExcelToolStripMenuItem";
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem.Size = new System.Drawing.Size(299, 22);
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem.Text = "Medida horaria ADIF --> Exportación Excel";
            this.medidaHorariaADIFExportaciónExcelToolStripMenuItem.Click += new System.EventHandler(this.medidaHorariaADIFExportaciónExcelToolStripMenuItem_Click);
            // 
            // FrmAdif
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 506);
            this.Controls.Add(this.txtFH);
            this.Controls.Add(this.txtFD);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnMail);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnImportar);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmAdif";
            this.Text = "ADIF";
            this.Load += new System.EventHandler(this.FrmAdif_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.Button btnImportar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnMail;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker txtFD;
        private System.Windows.Forms.DateTimePicker txtFH;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem informesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem medidaADIFFueraDeInventarioToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem medidaHorariaADIFAgrupadaVerticalToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem medidaHorariaADIFExportaciónExcelToolStripMenuItem;
    }
}