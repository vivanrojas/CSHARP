namespace GestionOperaciones.forms.contratacion
{
    partial class FrmPlantillaXML
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPlantillaXML));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btn_plantilla = new System.Windows.Forms.Button();
            this.btn_excel_datos = new System.Windows.Forms.Button();
            this.txt_datos_excel = new System.Windows.Forms.TextBox();
            this.txt_plantilla = new System.Windows.Forms.TextBox();
            this.btn_generar_xml = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
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
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // btn_plantilla
            // 
            this.btn_plantilla.Location = new System.Drawing.Point(48, 68);
            this.btn_plantilla.Name = "btn_plantilla";
            this.btn_plantilla.Size = new System.Drawing.Size(107, 39);
            this.btn_plantilla.TabIndex = 1;
            this.btn_plantilla.Text = "Importar plantilla XML";
            this.btn_plantilla.UseVisualStyleBackColor = true;
            this.btn_plantilla.Click += new System.EventHandler(this.btn_plantilla_Click);
            // 
            // btn_excel_datos
            // 
            this.btn_excel_datos.Location = new System.Drawing.Point(48, 113);
            this.btn_excel_datos.Name = "btn_excel_datos";
            this.btn_excel_datos.Size = new System.Drawing.Size(107, 39);
            this.btn_excel_datos.TabIndex = 2;
            this.btn_excel_datos.Text = "Importar Excel Datos";
            this.btn_excel_datos.UseVisualStyleBackColor = true;
            this.btn_excel_datos.Click += new System.EventHandler(this.btn_excel_datos_Click);
            // 
            // txt_datos_excel
            // 
            this.txt_datos_excel.Enabled = false;
            this.txt_datos_excel.Location = new System.Drawing.Point(161, 123);
            this.txt_datos_excel.Name = "txt_datos_excel";
            this.txt_datos_excel.Size = new System.Drawing.Size(596, 20);
            this.txt_datos_excel.TabIndex = 3;
            // 
            // txt_plantilla
            // 
            this.txt_plantilla.Enabled = false;
            this.txt_plantilla.Location = new System.Drawing.Point(161, 78);
            this.txt_plantilla.Name = "txt_plantilla";
            this.txt_plantilla.Size = new System.Drawing.Size(596, 20);
            this.txt_plantilla.TabIndex = 4;
            // 
            // btn_generar_xml
            // 
            this.btn_generar_xml.Location = new System.Drawing.Point(48, 204);
            this.btn_generar_xml.Name = "btn_generar_xml";
            this.btn_generar_xml.Size = new System.Drawing.Size(107, 39);
            this.btn_generar_xml.TabIndex = 5;
            this.btn_generar_xml.Text = "Generar XML";
            this.btn_generar_xml.UseVisualStyleBackColor = true;
            this.btn_generar_xml.Click += new System.EventHandler(this.btn_generar_xml_Click);
            // 
            // FrmPlantillaXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn_generar_xml);
            this.Controls.Add(this.txt_plantilla);
            this.Controls.Add(this.txt_datos_excel);
            this.Controls.Add(this.btn_excel_datos);
            this.Controls.Add(this.btn_plantilla);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmPlantillaXML";
            this.Text = "Creador XML desde plantilla";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmPlantillaXML_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.Button btn_plantilla;
        private System.Windows.Forms.Button btn_excel_datos;
        private System.Windows.Forms.TextBox txt_datos_excel;
        private System.Windows.Forms.TextBox txt_plantilla;
        private System.Windows.Forms.Button btn_generar_xml;
    }
}