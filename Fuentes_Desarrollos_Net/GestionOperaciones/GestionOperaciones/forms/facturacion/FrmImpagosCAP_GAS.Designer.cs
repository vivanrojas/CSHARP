namespace GestionOperaciones.forms.facturacion
{
    partial class FrmImpagosCAP_GAS
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmImpagosCAP_GAS));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.procesarExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.procesarExcelDeCUPSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.procesarExcelAgrupadasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
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
            this.procesarExcelToolStripMenuItem,
            this.procesarExcelDeCUPSToolStripMenuItem,
            this.procesarExcelAgrupadasToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // procesarExcelToolStripMenuItem
            // 
            this.procesarExcelToolStripMenuItem.Name = "procesarExcelToolStripMenuItem";
            this.procesarExcelToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            this.procesarExcelToolStripMenuItem.Text = "Procesar Excel de facturas";
            this.procesarExcelToolStripMenuItem.Click += new System.EventHandler(this.procesarExcelToolStripMenuItem_Click);
            // 
            // procesarExcelDeCUPSToolStripMenuItem
            // 
            this.procesarExcelDeCUPSToolStripMenuItem.Name = "procesarExcelDeCUPSToolStripMenuItem";
            this.procesarExcelDeCUPSToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            this.procesarExcelDeCUPSToolStripMenuItem.Text = "Procesar Excel de CUPS";
            this.procesarExcelDeCUPSToolStripMenuItem.Click += new System.EventHandler(this.procesarExcelDeCUPSToolStripMenuItem_Click);
            // 
            // procesarExcelAgrupadasToolStripMenuItem
            // 
            this.procesarExcelAgrupadasToolStripMenuItem.Name = "procesarExcelAgrupadasToolStripMenuItem";
            this.procesarExcelAgrupadasToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            this.procesarExcelAgrupadasToolStripMenuItem.Text = "Procesar Excel Agrupadas";
            this.procesarExcelAgrupadasToolStripMenuItem.Click += new System.EventHandler(this.procesarExcelAgrupadasToolStripMenuItem_Click);
            // 
            // FrmImpagosCAP_GAS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmImpagosCAP_GAS";
            this.Text = "RECLAMACIÓN JJAA FACTURAS IMPAGADAS CAP DE GAS";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmImpagosCAP_GAS_FormClosing);
            this.Load += new System.EventHandler(this.FrmImpagosCAP_GAS_Load);
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
        private System.Windows.Forms.ToolStripMenuItem procesarExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem procesarExcelDeCUPSToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem procesarExcelAgrupadasToolStripMenuItem;
    }
}