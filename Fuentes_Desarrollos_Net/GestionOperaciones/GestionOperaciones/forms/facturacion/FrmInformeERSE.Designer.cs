namespace GestionOperaciones.forms.facturacion
{
    partial class FrmInformeERSE
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInformeERSE));
            this.label2 = new System.Windows.Forms.Label();
            this.txtFFACTURAHAS = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFFACTURADES = new System.Windows.Forms.DateTimePicker();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.parametrizaciónToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.parámetrosERSEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.acercaDeInformeERSEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.cmdSearch = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(19, 92);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(240, 30);
            this.label2.TabIndex = 39;
            this.label2.Text = "Hasta Fecha Factura";
            // 
            // txtFFACTURAHAS
            // 
            this.txtFFACTURAHAS.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFFACTURAHAS.Location = new System.Drawing.Point(191, 87);
            this.txtFFACTURAHAS.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtFFACTURAHAS.Name = "txtFFACTURAHAS";
            this.txtFFACTURAHAS.Size = new System.Drawing.Size(118, 26);
            this.txtFFACTURAHAS.TabIndex = 38;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(19, 38);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(164, 20);
            this.label1.TabIndex = 37;
            this.label1.Text = "Desde Fecha Factura";
            // 
            // txtFFACTURADES
            // 
            this.txtFFACTURADES.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtFFACTURADES.Location = new System.Drawing.Point(191, 33);
            this.txtFFACTURADES.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtFFACTURADES.Name = "txtFFACTURADES";
            this.txtFFACTURADES.Size = new System.Drawing.Size(118, 26);
            this.txtFFACTURADES.TabIndex = 36;
            // 
            // menuStrip1
            // 
            this.menuStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.parametrizaciónToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(728, 33);
            this.menuStrip1.TabIndex = 44;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(88, 29);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Image = global::GestionOperaciones.Properties.Resources.if_dialog_close_29299;
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(212, 34);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // parametrizaciónToolStripMenuItem
            // 
            this.parametrizaciónToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.parámetrosERSEToolStripMenuItem});
            this.parametrizaciónToolStripMenuItem.Name = "parametrizaciónToolStripMenuItem";
            this.parametrizaciónToolStripMenuItem.Size = new System.Drawing.Size(152, 29);
            this.parametrizaciónToolStripMenuItem.Text = "Parametrización";
            this.parametrizaciónToolStripMenuItem.Click += new System.EventHandler(this.parametrizaciónToolStripMenuItem_Click);
            // 
            // parámetrosERSEToolStripMenuItem
            // 
            this.parámetrosERSEToolStripMenuItem.Name = "parámetrosERSEToolStripMenuItem";
            this.parámetrosERSEToolStripMenuItem.Size = new System.Drawing.Size(259, 34);
            this.parámetrosERSEToolStripMenuItem.Text = "Parámetros E.R.S.E";
            this.parámetrosERSEToolStripMenuItem.Click += new System.EventHandler(this.parámetrosERSEToolStripMenuItem_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.acercaDeInformeERSEToolStripMenuItem});
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(79, 29);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // acercaDeInformeERSEToolStripMenuItem
            // 
            this.acercaDeInformeERSEToolStripMenuItem.Name = "acercaDeInformeERSEToolStripMenuItem";
            this.acercaDeInformeERSEToolStripMenuItem.Size = new System.Drawing.Size(319, 34);
            this.acercaDeInformeERSEToolStripMenuItem.Text = "Acerca de Informe E.R.S.E.";
            this.acercaDeInformeERSEToolStripMenuItem.Click += new System.EventHandler(this.acercaDeInformeERSEToolStripMenuItem_Click);
            // 
            // cmdSearch
            // 
            this.cmdSearch.Image = global::GestionOperaciones.Properties.Resources.if_microsoft_office_excel_1784856;
            this.cmdSearch.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdSearch.Location = new System.Drawing.Point(433, 29);
            this.cmdSearch.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmdSearch.Name = "cmdSearch";
            this.cmdSearch.Size = new System.Drawing.Size(160, 93);
            this.cmdSearch.TabIndex = 40;
            this.cmdSearch.Text = "Exportar a Excel";
            this.cmdSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdSearch.UseVisualStyleBackColor = true;
            this.cmdSearch.Click += new System.EventHandler(this.cmdSearch_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtFFACTURADES);
            this.groupBox1.Controls.Add(this.cmdSearch);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtFFACTURAHAS);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(28, 60);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(661, 144);
            this.groupBox1.TabIndex = 45;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtro";
            // 
            // FrmInformeERSE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(728, 268);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FrmInformeERSE";
            this.Text = "Informe E.R.S.E.";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmInformeERSE_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Button cmdSearch;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker txtFFACTURAHAS;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker txtFFACTURADES;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem acercaDeInformeERSEToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem parametrizaciónToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem parámetrosERSEToolStripMenuItem;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}