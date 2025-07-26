namespace GestionOperaciones.forms.medida
{
    partial class FrmAdif_Importar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAdif_Importar));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtYYYYMM = new System.Windows.Forms.MaskedTextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnDescargaFTP = new System.Windows.Forms.Button();
            this.lblResumen = new System.Windows.Forms.Label();
            this.lblCuartoHoraria = new System.Windows.Forms.Label();
            this.lblAdif = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnResumen = new System.Windows.Forms.Button();
            this.btnCuartoHoraria = new System.Windows.Forms.Button();
            this.btnFicherosDat = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cerrarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtYYYYMM);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.btnDescargaFTP);
            this.groupBox1.Controls.Add(this.lblResumen);
            this.groupBox1.Controls.Add(this.lblCuartoHoraria);
            this.groupBox1.Controls.Add(this.lblAdif);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnResumen);
            this.groupBox1.Controls.Add(this.btnCuartoHoraria);
            this.groupBox1.Controls.Add(this.btnFicherosDat);
            this.groupBox1.Location = new System.Drawing.Point(12, 56);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1070, 214);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tareas";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // txtYYYYMM
            // 
            this.txtYYYYMM.Location = new System.Drawing.Point(883, 41);
            this.txtYYYYMM.Mask = "999999";
            this.txtYYYYMM.Name = "txtYYYYMM";
            this.txtYYYYMM.Size = new System.Drawing.Size(43, 20);
            this.txtYYYYMM.TabIndex = 14;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(932, 44);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "YYYYMM";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(509, 44);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(368, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Descarga Fichero ADIF desde FTP para el AñoMes de consumo";
            // 
            // btnDescargaFTP
            // 
            this.btnDescargaFTP.Image = global::GestionOperaciones.Properties.Resources.if_icon_129_cloud_download_314711;
            this.btnDescargaFTP.Location = new System.Drawing.Point(448, 29);
            this.btnDescargaFTP.Name = "btnDescargaFTP";
            this.btnDescargaFTP.Size = new System.Drawing.Size(55, 42);
            this.btnDescargaFTP.TabIndex = 10;
            this.toolTip1.SetToolTip(this.btnDescargaFTP, resources.GetString("btnDescargaFTP.ToolTip"));
            this.btnDescargaFTP.UseVisualStyleBackColor = true;
            this.btnDescargaFTP.Click += new System.EventHandler(this.btnDescargaFTP_Click);
            // 
            // lblResumen
            // 
            this.lblResumen.AutoSize = true;
            this.lblResumen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblResumen.Location = new System.Drawing.Point(279, 140);
            this.lblResumen.Name = "lblResumen";
            this.lblResumen.Size = new System.Drawing.Size(41, 13);
            this.lblResumen.TabIndex = 9;
            this.lblResumen.Text = "label6";
            // 
            // lblCuartoHoraria
            // 
            this.lblCuartoHoraria.AutoSize = true;
            this.lblCuartoHoraria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCuartoHoraria.Location = new System.Drawing.Point(279, 92);
            this.lblCuartoHoraria.Name = "lblCuartoHoraria";
            this.lblCuartoHoraria.Size = new System.Drawing.Size(41, 13);
            this.lblCuartoHoraria.TabIndex = 8;
            this.lblCuartoHoraria.Text = "label5";
            // 
            // lblAdif
            // 
            this.lblAdif.AutoSize = true;
            this.lblAdif.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAdif.Location = new System.Drawing.Point(279, 44);
            this.lblAdif.Name = "lblAdif";
            this.lblAdif.Size = new System.Drawing.Size(41, 13);
            this.lblAdif.TabIndex = 7;
            this.lblAdif.Text = "label5";
            this.lblAdif.Click += new System.EventHandler(this.lblAdif_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(244, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(165, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Última fecha de importación";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(76, 140);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(193, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Importar Curva-Resumen SCE -->";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(76, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(186, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Importar Cuarto-Horaria SCE -->";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(76, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(188, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Carga Manual Ficheros ADIF -->";
            // 
            // btnResumen
            // 
            this.btnResumen.Image = global::GestionOperaciones.Properties.Resources.if_f_right_256_282463;
            this.btnResumen.Location = new System.Drawing.Point(15, 125);
            this.btnResumen.Name = "btnResumen";
            this.btnResumen.Size = new System.Drawing.Size(55, 42);
            this.btnResumen.TabIndex = 2;
            this.btnResumen.UseVisualStyleBackColor = true;
            this.btnResumen.Click += new System.EventHandler(this.btnResumen_Click);
            // 
            // btnCuartoHoraria
            // 
            this.btnCuartoHoraria.Image = global::GestionOperaciones.Properties.Resources.if_f_right_256_282463;
            this.btnCuartoHoraria.Location = new System.Drawing.Point(15, 77);
            this.btnCuartoHoraria.Name = "btnCuartoHoraria";
            this.btnCuartoHoraria.Size = new System.Drawing.Size(55, 42);
            this.btnCuartoHoraria.TabIndex = 1;
            this.btnCuartoHoraria.UseVisualStyleBackColor = true;
            this.btnCuartoHoraria.Click += new System.EventHandler(this.btnCuartoHoraria_Click);
            // 
            // btnFicherosDat
            // 
            this.btnFicherosDat.Image = global::GestionOperaciones.Properties.Resources.if_f_right_256_282463;
            this.btnFicherosDat.Location = new System.Drawing.Point(15, 29);
            this.btnFicherosDat.Name = "btnFicherosDat";
            this.btnFicherosDat.Size = new System.Drawing.Size(55, 42);
            this.btnFicherosDat.TabIndex = 0;
            this.btnFicherosDat.UseVisualStyleBackColor = true;
            this.btnFicherosDat.Click += new System.EventHandler(this.btnFicherosDat_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1094, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cerrarToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // cerrarToolStripMenuItem
            // 
            this.cerrarToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_dialog_close_29299;
            this.cerrarToolStripMenuItem.Name = "cerrarToolStripMenuItem";
            this.cerrarToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.cerrarToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.cerrarToolStripMenuItem.Text = "Cerrar";
            this.cerrarToolStripMenuItem.Click += new System.EventHandler(this.cerrarToolStripMenuItem_Click);
            // 
            // FrmAdif_Importar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1094, 293);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmAdif_Importar";
            this.Text = "Carga de Datos";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmAdif_Importar_FormClosing);
            this.Load += new System.EventHandler(this.FrmAdif_Importar_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnResumen;
        private System.Windows.Forms.Button btnCuartoHoraria;
        private System.Windows.Forms.Button btnFicherosDat;
        private System.Windows.Forms.Label lblAdif;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblCuartoHoraria;
        private System.Windows.Forms.Label lblResumen;
        private System.Windows.Forms.Button btnDescargaFTP;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cerrarToolStripMenuItem;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.MaskedTextBox txtYYYYMM;
    }
}