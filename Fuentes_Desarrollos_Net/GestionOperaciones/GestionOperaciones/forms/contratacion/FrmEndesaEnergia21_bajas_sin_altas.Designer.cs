namespace GestionOperaciones.forms.contratacion
{
    partial class FrmEndesaEnergia21_bajas_sin_altas
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmEndesaEnergia21_bajas_sin_altas));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cups = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.proceso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.paso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.solicitud = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_solicitud = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lbl_total_solicitudes = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1004, 24);
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
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 42);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(980, 488);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lbl_total_solicitudes);
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(972, 462);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Solicitudes";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cups,
            this.proceso,
            this.paso,
            this.solicitud,
            this.fecha_solicitud});
            this.dgv.Location = new System.Drawing.Point(3, 44);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(963, 412);
            this.dgv.TabIndex = 0;
            // 
            // cups
            // 
            this.cups.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups.DataPropertyName = "cups";
            this.cups.HeaderText = "CUPS";
            this.cups.Name = "cups";
            this.cups.ReadOnly = true;
            this.cups.Width = 61;
            // 
            // proceso
            // 
            this.proceso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.proceso.DataPropertyName = "proceso";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.proceso.DefaultCellStyle = dataGridViewCellStyle2;
            this.proceso.HeaderText = "Proceso";
            this.proceso.Name = "proceso";
            this.proceso.ReadOnly = true;
            this.proceso.Width = 71;
            // 
            // paso
            // 
            this.paso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.paso.DataPropertyName = "paso";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.paso.DefaultCellStyle = dataGridViewCellStyle3;
            this.paso.HeaderText = "Paso";
            this.paso.Name = "paso";
            this.paso.ReadOnly = true;
            this.paso.Width = 56;
            // 
            // solicitud
            // 
            this.solicitud.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.solicitud.DataPropertyName = "solicitud";
            this.solicitud.HeaderText = "Solicitud";
            this.solicitud.Name = "solicitud";
            this.solicitud.ReadOnly = true;
            this.solicitud.Width = 72;
            // 
            // fecha_solicitud
            // 
            this.fecha_solicitud.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_solicitud.DataPropertyName = "fecha_solicitud";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.fecha_solicitud.DefaultCellStyle = dataGridViewCellStyle4;
            this.fecha_solicitud.HeaderText = "Fecha Solicitud";
            this.fecha_solicitud.Name = "fecha_solicitud";
            this.fecha_solicitud.ReadOnly = true;
            this.fecha_solicitud.Width = 97;
            // 
            // lbl_total_solicitudes
            // 
            this.lbl_total_solicitudes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_total_solicitudes.AutoSize = true;
            this.lbl_total_solicitudes.Location = new System.Drawing.Point(817, 15);
            this.lbl_total_solicitudes.Name = "lbl_total_solicitudes";
            this.lbl_total_solicitudes.Size = new System.Drawing.Size(139, 13);
            this.lbl_total_solicitudes.TabIndex = 62;
            this.lbl_total_solicitudes.Text = "Total Solicitudes:                 ";
            this.lbl_total_solicitudes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // FrmEndesaEnergia21_bajas_sin_altas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1004, 542);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmEndesaEnergia21_bajas_sin_altas";
            this.Text = "Bajas sin Altas";
            this.Load += new System.EventHandler(this.FrmEndesaEnergia21_bajas_sin_altas_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups;
        private System.Windows.Forms.DataGridViewTextBoxColumn proceso;
        private System.Windows.Forms.DataGridViewTextBoxColumn paso;
        private System.Windows.Forms.DataGridViewTextBoxColumn solicitud;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_solicitud;
        private System.Windows.Forms.Label lbl_total_solicitudes;
    }
}