namespace GestionOperaciones.forms.facturacion
{
    partial class FrmTunel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmTunel));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.acercaDeClausulaTúnelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btn_Carga_Excel = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.chk_Mensual = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.editarToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1339, 24);
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
            // editarToolStripMenuItem
            // 
            this.editarToolStripMenuItem.Name = "editarToolStripMenuItem";
            this.editarToolStripMenuItem.Size = new System.Drawing.Size(49, 20);
            this.editarToolStripMenuItem.Text = "Editar";
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.acercaDeClausulaTúnelToolStripMenuItem});
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // acercaDeClausulaTúnelToolStripMenuItem
            // 
            this.acercaDeClausulaTúnelToolStripMenuItem.Name = "acercaDeClausulaTúnelToolStripMenuItem";
            this.acercaDeClausulaTúnelToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.acercaDeClausulaTúnelToolStripMenuItem.Text = "Acerca de Clausula Túnel";
            this.acercaDeClausulaTúnelToolStripMenuItem.Click += new System.EventHandler(this.acercaDeClausulaTúnelToolStripMenuItem_Click);
            // 
            // btn_Carga_Excel
            // 
            this.btn_Carga_Excel.Location = new System.Drawing.Point(1195, 187);
            this.btn_Carga_Excel.Name = "btn_Carga_Excel";
            this.btn_Carga_Excel.Size = new System.Drawing.Size(132, 49);
            this.btn_Carga_Excel.TabIndex = 1;
            this.btn_Carga_Excel.Text = "Cargar Excels Tunel";
            this.btn_Carga_Excel.UseVisualStyleBackColor = true;
            this.btn_Carga_Excel.Click += new System.EventHandler(this.btn_Carga_Excel_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(49, 61);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(326, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "Ejemplo de la estructura del archivo Excel a importar";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::GestionOperaciones.Properties.Resources.EstructuraArchivoTunel1;
            this.pictureBox1.Location = new System.Drawing.Point(49, 87);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1278, 79);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // chk_Mensual
            // 
            this.chk_Mensual.AutoSize = true;
            this.chk_Mensual.Location = new System.Drawing.Point(1064, 204);
            this.chk_Mensual.Name = "chk_Mensual";
            this.chk_Mensual.Size = new System.Drawing.Size(115, 17);
            this.chk_Mensual.TabIndex = 4;
            this.chk_Mensual.Text = "Agrupado Mensual";
            this.chk_Mensual.UseVisualStyleBackColor = true;
            // 
            // FrmTunel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1339, 267);
            this.Controls.Add(this.chk_Mensual);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btn_Carga_Excel);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmTunel";
            this.Text = "Clausula Túnel";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmTunel_FormClosing);
            this.Load += new System.EventHandler(this.FrmTunel_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.Button btn_Carga_Excel;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ToolStripMenuItem acercaDeClausulaTúnelToolStripMenuItem;
        private System.Windows.Forms.CheckBox chk_Mensual;
    }
}