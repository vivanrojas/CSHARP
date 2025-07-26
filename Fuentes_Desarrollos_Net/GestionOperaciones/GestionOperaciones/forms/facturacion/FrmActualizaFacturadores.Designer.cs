namespace GestionOperaciones.forms.facturacion
{
    partial class FrmActualizaFacturadores
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmActualizaFacturadores));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.grp_programas_consumo = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_programas_consumo = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.grp_programas_consumo.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
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
            // grp_programas_consumo
            // 
            this.grp_programas_consumo.Controls.Add(this.label1);
            this.grp_programas_consumo.Controls.Add(this.btn_programas_consumo);
            this.grp_programas_consumo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grp_programas_consumo.Location = new System.Drawing.Point(45, 61);
            this.grp_programas_consumo.Name = "grp_programas_consumo";
            this.grp_programas_consumo.Size = new System.Drawing.Size(316, 61);
            this.grp_programas_consumo.TabIndex = 1;
            this.grp_programas_consumo.TabStop = false;
            this.grp_programas_consumo.Text = "Programas de consumo";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(40, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(212, 13);
            this.label1.TabIndex = 65;
            this.label1.Text = "Carga fichero Excel programas de consumo";
            // 
            // btn_programas_consumo
            // 
            this.btn_programas_consumo.Image = ((System.Drawing.Image)(resources.GetObject("btn_programas_consumo.Image")));
            this.btn_programas_consumo.Location = new System.Drawing.Point(6, 19);
            this.btn_programas_consumo.Name = "btn_programas_consumo";
            this.btn_programas_consumo.Size = new System.Drawing.Size(28, 28);
            this.btn_programas_consumo.TabIndex = 64;
            this.btn_programas_consumo.UseVisualStyleBackColor = true;
            this.btn_programas_consumo.Click += new System.EventHandler(this.btn_programas_consumo_Click);
            // 
            // FrmActualizaFacturadores
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.grp_programas_consumo);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmActualizaFacturadores";
            this.Text = "Actualiza Facturadores";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmActualizaFacturadores_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.grp_programas_consumo.ResumeLayout(false);
            this.grp_programas_consumo.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.GroupBox grp_programas_consumo;
        private System.Windows.Forms.Button btn_programas_consumo;
        private System.Windows.Forms.Label label1;
    }
}