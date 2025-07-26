namespace GestionOperaciones.forms.herramientas
{
    partial class FrmCarteraSalesForce
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCarteraSalesForce));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarCarteraSalesForceCSVToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem});
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
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarCarteraSalesForceCSVToolStripMenuItem,
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarCarteraSalesForceCSVToolStripMenuItem
            // 
            this.importarCarteraSalesForceCSVToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.Spain_flags_flag_8858_1_;
            this.importarCarteraSalesForceCSVToolStripMenuItem.Name = "importarCarteraSalesForceCSVToolStripMenuItem";
            this.importarCarteraSalesForceCSVToolStripMenuItem.Size = new System.Drawing.Size(299, 22);
            this.importarCarteraSalesForceCSVToolStripMenuItem.Text = "Importar Cartera SalesForce España (CSV)";
            this.importarCarteraSalesForceCSVToolStripMenuItem.Click += new System.EventHandler(this.importarCarteraSalesForceCSVToolStripMenuItem_Click);
            // 
            // importarCarteraSalesForcePortugalCSVToolStripMenuItem
            // 
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.portugal;
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem.Name = "importarCarteraSalesForcePortugalCSVToolStripMenuItem";
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem.Size = new System.Drawing.Size(299, 22);
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem.Text = "Importar Cartera SalesForce Portugal (CSV)";
            this.importarCarteraSalesForcePortugalCSVToolStripMenuItem.Click += new System.EventHandler(this.importarCarteraSalesForcePortugalCSVToolStripMenuItem_Click);
            // 
            // FrmCarteraSalesForce
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmCarteraSalesForce";
            this.Text = "Cartera SalesForce";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmCarteraSalesForce_FormClosing);
            this.Load += new System.EventHandler(this.FrmCarteraSalesForce_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarCarteraSalesForceCSVToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarCarteraSalesForcePortugalCSVToolStripMenuItem;
    }
}