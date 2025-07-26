namespace GestionOperaciones.forms.herramientas
{
    partial class FrmSeguimiento_Procesos
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSeguimiento_Procesos));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.opcionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.listBoxProceso = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.listBoxArea = new System.Windows.Forms.ListBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cnifdnic = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dapersoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.distribuidora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cups20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentarios_descuadres = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentarios_contratacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tramitacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.hora_inicio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tarea = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contacto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgvd = new System.Windows.Forms.DataGridView();
            this.area = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.proceso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tipo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_inicio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_fin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentario = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.tabControl2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvd)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.herramientasToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(8, 3, 0, 3);
            this.menuStrip1.Size = new System.Drawing.Size(949, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 19);
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
            this.opcionesToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 19);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // opcionesToolStripMenuItem
            // 
            this.opcionesToolStripMenuItem.Name = "opcionesToolStripMenuItem";
            this.opcionesToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.opcionesToolStripMenuItem.Text = "Opciones";
            this.opcionesToolStripMenuItem.Click += new System.EventHandler(this.opcionesToolStripMenuItem_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Image = global::GestionOperaciones.Properties.Resources.if_view_refresh_118801;
            this.btnRefresh.Location = new System.Drawing.Point(755, 70);
            this.btnRefresh.Margin = new System.Windows.Forms.Padding(4);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(57, 54);
            this.btnRefresh.TabIndex = 78;
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(222, 48);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 77;
            this.label2.Text = "Proceso";
            // 
            // listBoxProceso
            // 
            this.listBoxProceso.FormattingEnabled = true;
            this.listBoxProceso.ItemHeight = 12;
            this.listBoxProceso.Location = new System.Drawing.Point(227, 70);
            this.listBoxProceso.Margin = new System.Windows.Forms.Padding(4);
            this.listBoxProceso.Name = "listBoxProceso";
            this.listBoxProceso.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxProceso.Size = new System.Drawing.Size(520, 76);
            this.listBoxProceso.TabIndex = 76;
            this.listBoxProceso.SelectedIndexChanged += new System.EventHandler(this.listBoxProceso_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 48);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 75;
            this.label1.Text = "Área";
            // 
            // listBoxArea
            // 
            this.listBoxArea.FormattingEnabled = true;
            this.listBoxArea.ItemHeight = 12;
            this.listBoxArea.Location = new System.Drawing.Point(12, 70);
            this.listBoxArea.Margin = new System.Windows.Forms.Padding(4);
            this.listBoxArea.Name = "listBoxArea";
            this.listBoxArea.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxArea.Size = new System.Drawing.Size(204, 76);
            this.listBoxArea.TabIndex = 74;
            this.listBoxArea.SelectedIndexChanged += new System.EventHandler(this.listBoxArea_SelectedIndexChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 156);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(926, 224);
            this.tabControl1.TabIndex = 79;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgv);
            this.tabPage2.Location = new System.Drawing.Point(4, 21);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage2.Size = new System.Drawing.Size(918, 199);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Procesos";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.White;
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cnifdnic,
            this.dapersoc,
            this.distribuidora,
            this.cups20,
            this.comentarios_descuadres,
            this.comentarios_contratacion,
            this.tramitacion,
            this.hora_inicio,
            this.tarea,
            this.contacto});
            this.dgv.Location = new System.Drawing.Point(7, 8);
            this.dgv.Margin = new System.Windows.Forms.Padding(4);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.RowHeadersWidth = 62;
            this.dgv.Size = new System.Drawing.Size(904, 179);
            this.dgv.TabIndex = 60;
            this.dgv.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellClick_1);
            this.dgv.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellContentClick);
            // 
            // cnifdnic
            // 
            this.cnifdnic.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cnifdnic.DataPropertyName = "area";
            this.cnifdnic.HeaderText = "Área";
            this.cnifdnic.MinimumWidth = 8;
            this.cnifdnic.Name = "cnifdnic";
            this.cnifdnic.ReadOnly = true;
            this.cnifdnic.Width = 54;
            // 
            // dapersoc
            // 
            this.dapersoc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dapersoc.DataPropertyName = "proceso";
            this.dapersoc.HeaderText = "Proceso";
            this.dapersoc.MinimumWidth = 8;
            this.dapersoc.Name = "dapersoc";
            this.dapersoc.ReadOnly = true;
            this.dapersoc.Width = 70;
            // 
            // distribuidora
            // 
            this.distribuidora.DataPropertyName = "descripcion";
            this.distribuidora.HeaderText = "Descripción";
            this.distribuidora.MinimumWidth = 8;
            this.distribuidora.Name = "distribuidora";
            this.distribuidora.ReadOnly = true;
            this.distribuidora.Width = 88;
            // 
            // cups20
            // 
            this.cups20.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cups20.DataPropertyName = "fecha_inicio";
            this.cups20.HeaderText = "Inicio";
            this.cups20.MinimumWidth = 8;
            this.cups20.Name = "cups20";
            this.cups20.ReadOnly = true;
            this.cups20.Width = 56;
            // 
            // comentarios_descuadres
            // 
            this.comentarios_descuadres.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios_descuadres.DataPropertyName = "fecha_fin";
            this.comentarios_descuadres.HeaderText = "Fin";
            this.comentarios_descuadres.MinimumWidth = 8;
            this.comentarios_descuadres.Name = "comentarios_descuadres";
            this.comentarios_descuadres.ReadOnly = true;
            this.comentarios_descuadres.Width = 46;
            // 
            // comentarios_contratacion
            // 
            this.comentarios_contratacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentarios_contratacion.DataPropertyName = "comentario";
            this.comentarios_contratacion.HeaderText = "Comentario";
            this.comentarios_contratacion.MinimumWidth = 8;
            this.comentarios_contratacion.Name = "comentarios_contratacion";
            this.comentarios_contratacion.ReadOnly = true;
            this.comentarios_contratacion.Width = 85;
            // 
            // tramitacion
            // 
            this.tramitacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tramitacion.DataPropertyName = "ejecucion";
            this.tramitacion.HeaderText = "Periodicidad";
            this.tramitacion.MinimumWidth = 8;
            this.tramitacion.Name = "tramitacion";
            this.tramitacion.ReadOnly = true;
            this.tramitacion.Width = 89;
            // 
            // hora_inicio
            // 
            this.hora_inicio.DataPropertyName = "hora_inicio";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.Format = "t";
            dataGridViewCellStyle2.NullValue = null;
            this.hora_inicio.DefaultCellStyle = dataGridViewCellStyle2;
            this.hora_inicio.HeaderText = "Hora Inicio";
            this.hora_inicio.Name = "hora_inicio";
            this.hora_inicio.ReadOnly = true;
            // 
            // tarea
            // 
            this.tarea.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tarea.DataPropertyName = "tarea";
            this.tarea.HeaderText = "Tarea";
            this.tarea.MinimumWidth = 8;
            this.tarea.Name = "tarea";
            this.tarea.ReadOnly = true;
            this.tarea.Width = 58;
            // 
            // contacto
            // 
            this.contacto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.contacto.DataPropertyName = "contacto";
            this.contacto.HeaderText = "Contacto";
            this.contacto.MinimumWidth = 8;
            this.contacto.Name = "contacto";
            this.contacto.ReadOnly = true;
            this.contacto.Width = 74;
            // 
            // tabControl2
            // 
            this.tabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl2.Controls.Add(this.tabPage1);
            this.tabControl2.Location = new System.Drawing.Point(12, 388);
            this.tabControl2.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(926, 150);
            this.tabControl2.TabIndex = 80;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgvd);
            this.tabPage1.Location = new System.Drawing.Point(4, 21);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(918, 125);
            this.tabPage1.TabIndex = 1;
            this.tabPage1.Text = "Pasos";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgvd
            // 
            this.dgvd.AllowUserToAddRows = false;
            this.dgvd.AllowUserToDeleteRows = false;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.White;
            this.dgvd.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgvd.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvd.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.area,
            this.proceso,
            this.tipo,
            this.fecha_inicio,
            this.fecha_fin,
            this.comentario});
            this.dgvd.Location = new System.Drawing.Point(4, 8);
            this.dgvd.Margin = new System.Windows.Forms.Padding(4);
            this.dgvd.Name = "dgvd";
            this.dgvd.ReadOnly = true;
            this.dgvd.RowHeadersWidth = 62;
            this.dgvd.Size = new System.Drawing.Size(908, 108);
            this.dgvd.TabIndex = 60;
            // 
            // area
            // 
            this.area.DataPropertyName = "area";
            this.area.HeaderText = "area";
            this.area.MinimumWidth = 8;
            this.area.Name = "area";
            this.area.ReadOnly = true;
            this.area.Visible = false;
            this.area.Width = 150;
            // 
            // proceso
            // 
            this.proceso.DataPropertyName = "proceso";
            this.proceso.HeaderText = "proceso";
            this.proceso.MinimumWidth = 8;
            this.proceso.Name = "proceso";
            this.proceso.ReadOnly = true;
            this.proceso.Visible = false;
            this.proceso.Width = 150;
            // 
            // tipo
            // 
            this.tipo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.tipo.DataPropertyName = "paso";
            this.tipo.HeaderText = "Paso";
            this.tipo.MinimumWidth = 8;
            this.tipo.Name = "tipo";
            this.tipo.ReadOnly = true;
            this.tipo.Width = 56;
            // 
            // fecha_inicio
            // 
            this.fecha_inicio.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_inicio.DataPropertyName = "fecha_inicio";
            this.fecha_inicio.HeaderText = "Fecha inicio";
            this.fecha_inicio.MinimumWidth = 8;
            this.fecha_inicio.Name = "fecha_inicio";
            this.fecha_inicio.ReadOnly = true;
            this.fecha_inicio.Width = 87;
            // 
            // fecha_fin
            // 
            this.fecha_fin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_fin.DataPropertyName = "fecha_fin";
            this.fecha_fin.HeaderText = "Fecha fin";
            this.fecha_fin.MinimumWidth = 8;
            this.fecha_fin.Name = "fecha_fin";
            this.fecha_fin.ReadOnly = true;
            this.fecha_fin.Width = 75;
            // 
            // comentario
            // 
            this.comentario.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentario.DataPropertyName = "comentario";
            this.comentario.HeaderText = "Comentario";
            this.comentario.MinimumWidth = 8;
            this.comentario.Name = "comentario";
            this.comentario.ReadOnly = true;
            this.comentario.Width = 85;
            // 
            // FrmSeguimiento_Procesos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(949, 550);
            this.Controls.Add(this.tabControl2);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listBoxProceso);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBoxArea);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FrmSeguimiento_Procesos";
            this.Text = "Seguimiento Procesos";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmSeguimiento_Procesos_FormClosing);
            this.Load += new System.EventHandler(this.FrmSeguimiento_Procesos_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.tabControl2.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvd)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox listBoxProceso;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBoxArea;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem opcionesToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dgvd;
        private System.Windows.Forms.DataGridViewTextBoxColumn cnifdnic;
        private System.Windows.Forms.DataGridViewTextBoxColumn dapersoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn distribuidora;
        private System.Windows.Forms.DataGridViewTextBoxColumn cups20;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios_descuadres;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentarios_contratacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn tramitacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn hora_inicio;
        private System.Windows.Forms.DataGridViewTextBoxColumn tarea;
        private System.Windows.Forms.DataGridViewTextBoxColumn contacto;
        private System.Windows.Forms.DataGridViewTextBoxColumn area;
        private System.Windows.Forms.DataGridViewTextBoxColumn proceso;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipo;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_inicio;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_fin;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentario;
    }
}