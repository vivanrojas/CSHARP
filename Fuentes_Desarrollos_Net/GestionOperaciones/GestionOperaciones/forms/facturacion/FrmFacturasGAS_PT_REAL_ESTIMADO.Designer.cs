namespace GestionOperaciones.forms.facturacion
{
    partial class FrmFacturasGAS_PT_REAL_ESTIMADO
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFacturasGAS_PT_REAL_ESTIMADO));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_fecha_consumo_hasta = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_fecha_consumo_desde = new System.Windows.Forms.DateTimePicker();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_excel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_fecha_factura_hasta = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_fecha_factura_desde = new System.Windows.Forms.DateTimePicker();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_fecha_consumo_hasta);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txt_fecha_consumo_desde);
            this.groupBox1.Location = new System.Drawing.Point(117, 157);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(567, 137);
            this.groupBox1.TabIndex = 25;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtros";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Image = global::GestionOperaciones.Properties.Resources.if_microsoft_office_excel_1784856;
            this.btnRefresh.Location = new System.Drawing.Point(465, 33);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 52);
            this.btnRefresh.TabIndex = 27;
            this.btnRefresh.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(239, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(134, 13);
            this.label3.TabIndex = 24;
            this.label3.Text = "Hasta Fecha Consumo";
            // 
            // txt_fecha_consumo_hasta
            // 
            this.txt_fecha_consumo_hasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_consumo_hasta.Location = new System.Drawing.Point(379, 47);
            this.txt_fecha_consumo_hasta.Name = "txt_fecha_consumo_hasta";
            this.txt_fecha_consumo_hasta.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_consumo_hasta.TabIndex = 23;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(7, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(137, 13);
            this.label4.TabIndex = 22;
            this.label4.Text = "Desde Fecha Consumo";
            // 
            // txt_fecha_consumo_desde
            // 
            this.txt_fecha_consumo_desde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_consumo_desde.Location = new System.Drawing.Point(150, 47);
            this.txt_fecha_consumo_desde.Name = "txt_fecha_consumo_desde";
            this.txt_fecha_consumo_desde.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_consumo_desde.TabIndex = 21;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(589, 24);
            this.menuStrip1.TabIndex = 26;
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
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_excel);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txt_fecha_factura_hasta);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txt_fecha_factura_desde);
            this.groupBox2.Location = new System.Drawing.Point(12, 34);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(567, 137);
            this.groupBox2.TabIndex = 27;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Filtros";
            // 
            // btn_excel
            // 
            this.btn_excel.Image = global::GestionOperaciones.Properties.Resources.if_microsoft_office_excel_1784856;
            this.btn_excel.Location = new System.Drawing.Point(465, 33);
            this.btn_excel.Name = "btn_excel";
            this.btn_excel.Size = new System.Drawing.Size(75, 52);
            this.btn_excel.TabIndex = 27;
            this.btn_excel.UseVisualStyleBackColor = true;
            this.btn_excel.Click += new System.EventHandler(this.btn_excel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(247, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 13);
            this.label1.TabIndex = 24;
            this.label1.Text = "Hasta Fecha Factura";
            // 
            // txt_fecha_factura_hasta
            // 
            this.txt_fecha_factura_hasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_factura_hasta.Location = new System.Drawing.Point(379, 47);
            this.txt_fecha_factura_hasta.Name = "txt_fecha_factura_hasta";
            this.txt_fecha_factura_hasta.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_factura_hasta.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(15, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(129, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Desde Fecha Factura";
            // 
            // txt_fecha_factura_desde
            // 
            this.txt_fecha_factura_desde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_factura_desde.Location = new System.Drawing.Point(150, 47);
            this.txt_fecha_factura_desde.Name = "txt_fecha_factura_desde";
            this.txt_fecha_factura_desde.Size = new System.Drawing.Size(80, 20);
            this.txt_fecha_factura_desde.TabIndex = 21;
            // 
            // FrmFacturasGAS_PT_REAL_ESTIMADO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(589, 186);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmFacturasGAS_PT_REAL_ESTIMADO";
            this.Text = "Facturas GAS Portugal (Real / Estimado)";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmFacturasGAS_PT_REAL_ESTIMADO_FormClosing);
            this.Load += new System.EventHandler(this.FrmFacturasGAS_PT_REAL_ESTIMADO_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker txt_fecha_consumo_hasta;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker txt_fecha_consumo_desde;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_excel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker txt_fecha_factura_hasta;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker txt_fecha_factura_desde;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
    }
}