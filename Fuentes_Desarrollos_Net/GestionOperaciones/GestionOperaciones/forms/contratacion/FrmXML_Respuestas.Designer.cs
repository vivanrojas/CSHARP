namespace GestionOperaciones.forms.contratacion
{
    partial class FrmXML_Respuestas
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmXML_Respuestas));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cargaXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.area = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.proceso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.paso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.descripcion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_inicio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha_fin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comentario = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.listBoxProceso = new System.Windows.Forms.ListBox();
            this.listBoxArea = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_c205 = new System.Windows.Forms.Button();
            this.btn_c202 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_m105 = new System.Windows.Forms.Button();
            this.btn_m102 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_b105 = new System.Windows.Forms.Button();
            this.btn_b102 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_C105 = new System.Windows.Forms.Button();
            this.btn_c102 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_a305 = new System.Windows.Forms.Button();
            this.btn_a302 = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.groupBox1.SuspendLayout();
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
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 1, 0, 1);
            this.menuStrip1.Size = new System.Drawing.Size(748, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 22);
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
            this.cargaXMLToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 22);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // cargaXMLToolStripMenuItem
            // 
            this.cargaXMLToolStripMenuItem.Name = "cargaXMLToolStripMenuItem";
            this.cargaXMLToolStripMenuItem.Size = new System.Drawing.Size(132, 22);
            this.cargaXMLToolStripMenuItem.Text = "Carga XML";
            this.cargaXMLToolStripMenuItem.Click += new System.EventHandler(this.cargaXMLToolStripMenuItem_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 183);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(724, 308);
            this.tabControl1.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(716, 282);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Archivos XML";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.id,
            this.area,
            this.proceso,
            this.paso,
            this.descripcion,
            this.fecha_inicio,
            this.fecha_fin,
            this.comentario});
            this.dgv.Location = new System.Drawing.Point(6, 6);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.RowHeadersWidth = 51;
            this.dgv.Size = new System.Drawing.Size(704, 270);
            this.dgv.TabIndex = 0;
            // 
            // id
            // 
            this.id.DataPropertyName = "id";
            this.id.HeaderText = "id";
            this.id.MinimumWidth = 8;
            this.id.Name = "id";
            this.id.ReadOnly = true;
            this.id.Visible = false;
            this.id.Width = 150;
            // 
            // area
            // 
            this.area.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.area.DataPropertyName = "CodigoREEEmpresaEmisora";
            this.area.HeaderText = "Emisora";
            this.area.MinimumWidth = 6;
            this.area.Name = "area";
            this.area.ReadOnly = true;
            this.area.Width = 69;
            // 
            // proceso
            // 
            this.proceso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.proceso.DataPropertyName = "CodigoREEEmpresaDestino";
            this.proceso.HeaderText = "Destino";
            this.proceso.MinimumWidth = 6;
            this.proceso.Name = "proceso";
            this.proceso.ReadOnly = true;
            this.proceso.Width = 68;
            // 
            // paso
            // 
            this.paso.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.paso.DataPropertyName = "CodigoDelProceso";
            this.paso.HeaderText = "Proceso";
            this.paso.MinimumWidth = 6;
            this.paso.Name = "paso";
            this.paso.ReadOnly = true;
            this.paso.Width = 71;
            // 
            // descripcion
            // 
            this.descripcion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.descripcion.DataPropertyName = "CodigoDePaso";
            this.descripcion.HeaderText = "Paso";
            this.descripcion.MinimumWidth = 6;
            this.descripcion.Name = "descripcion";
            this.descripcion.ReadOnly = true;
            this.descripcion.Width = 56;
            // 
            // fecha_inicio
            // 
            this.fecha_inicio.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_inicio.DataPropertyName = "CodigoDeSolicitud";
            this.fecha_inicio.HeaderText = "Cod. Solicitud";
            this.fecha_inicio.MinimumWidth = 6;
            this.fecha_inicio.Name = "fecha_inicio";
            this.fecha_inicio.ReadOnly = true;
            this.fecha_inicio.Width = 97;
            // 
            // fecha_fin
            // 
            this.fecha_fin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.fecha_fin.DataPropertyName = "FechaSolicitud";
            this.fecha_fin.HeaderText = "F. Solicitud";
            this.fecha_fin.MinimumWidth = 6;
            this.fecha_fin.Name = "fecha_fin";
            this.fecha_fin.ReadOnly = true;
            this.fecha_fin.Width = 84;
            // 
            // comentario
            // 
            this.comentario.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comentario.DataPropertyName = "CUPS";
            this.comentario.HeaderText = "CUPS";
            this.comentario.MinimumWidth = 6;
            this.comentario.Name = "comentario";
            this.comentario.ReadOnly = true;
            this.comentario.Width = 61;
            // 
            // listBoxProceso
            // 
            this.listBoxProceso.FormattingEnabled = true;
            this.listBoxProceso.Location = new System.Drawing.Point(118, 55);
            this.listBoxProceso.Name = "listBoxProceso";
            this.listBoxProceso.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxProceso.Size = new System.Drawing.Size(107, 95);
            this.listBoxProceso.TabIndex = 73;
            // 
            // listBoxArea
            // 
            this.listBoxArea.FormattingEnabled = true;
            this.listBoxArea.Location = new System.Drawing.Point(12, 55);
            this.listBoxArea.Name = "listBoxArea";
            this.listBoxArea.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxArea.Size = new System.Drawing.Size(88, 95);
            this.listBoxArea.TabIndex = 72;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_c205);
            this.groupBox1.Controls.Add(this.btn_c202);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.btn_m105);
            this.groupBox1.Controls.Add(this.btn_m102);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btn_b105);
            this.groupBox1.Controls.Add(this.btn_b102);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btn_C105);
            this.groupBox1.Controls.Add(this.btn_c102);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btn_a305);
            this.groupBox1.Controls.Add(this.btn_a302);
            this.groupBox1.Location = new System.Drawing.Point(246, 55);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(486, 109);
            this.groupBox1.TabIndex = 74;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Generación de pasos";
            // 
            // btn_c205
            // 
            this.btn_c205.Location = new System.Drawing.Point(341, 72);
            this.btn_c205.Name = "btn_c205";
            this.btn_c205.Size = new System.Drawing.Size(75, 23);
            this.btn_c205.TabIndex = 14;
            this.btn_c205.Text = "C205";
            this.btn_c205.UseVisualStyleBackColor = true;
            this.btn_c205.Click += new System.EventHandler(this.btn_c205_Click);
            // 
            // btn_c202
            // 
            this.btn_c202.Location = new System.Drawing.Point(341, 36);
            this.btn_c202.Name = "btn_c202";
            this.btn_c202.Size = new System.Drawing.Size(75, 23);
            this.btn_c202.TabIndex = 13;
            this.btn_c202.Text = "C202";
            this.btn_c202.UseVisualStyleBackColor = true;
            this.btn_c202.Click += new System.EventHandler(this.btn_c202_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(348, 20);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "ACC + MOD";
            // 
            // btn_m105
            // 
            this.btn_m105.Location = new System.Drawing.Point(168, 72);
            this.btn_m105.Name = "btn_m105";
            this.btn_m105.Size = new System.Drawing.Size(74, 23);
            this.btn_m105.TabIndex = 11;
            this.btn_m105.Text = "M105";
            this.btn_m105.UseVisualStyleBackColor = true;
            // 
            // btn_m102
            // 
            this.btn_m102.Location = new System.Drawing.Point(168, 36);
            this.btn_m102.Name = "btn_m102";
            this.btn_m102.Size = new System.Drawing.Size(74, 23);
            this.btn_m102.TabIndex = 10;
            this.btn_m102.Text = "M102";
            this.btn_m102.UseVisualStyleBackColor = true;
            this.btn_m102.Click += new System.EventHandler(this.button5_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(196, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "MOD";
            // 
            // btn_b105
            // 
            this.btn_b105.Location = new System.Drawing.Point(259, 73);
            this.btn_b105.Name = "btn_b105";
            this.btn_b105.Size = new System.Drawing.Size(76, 22);
            this.btn_b105.TabIndex = 8;
            this.btn_b105.Text = "B105";
            this.btn_b105.UseVisualStyleBackColor = true;
            this.btn_b105.Click += new System.EventHandler(this.btn_b105_Click);
            // 
            // btn_b102
            // 
            this.btn_b102.Location = new System.Drawing.Point(259, 36);
            this.btn_b102.Name = "btn_b102";
            this.btn_b102.Size = new System.Drawing.Size(76, 23);
            this.btn_b102.TabIndex = 7;
            this.btn_b102.Text = "B102";
            this.btn_b102.UseVisualStyleBackColor = true;
            this.btn_b102.Click += new System.EventHandler(this.btn_b102_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(275, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(24, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "B A";
            // 
            // btn_C105
            // 
            this.btn_C105.Location = new System.Drawing.Point(87, 72);
            this.btn_C105.Name = "btn_C105";
            this.btn_C105.Size = new System.Drawing.Size(75, 23);
            this.btn_C105.TabIndex = 5;
            this.btn_C105.Text = "C105";
            this.btn_C105.UseVisualStyleBackColor = true;
            this.btn_C105.Click += new System.EventHandler(this.btn_C105_Click);
            // 
            // btn_c102
            // 
            this.btn_c102.Location = new System.Drawing.Point(87, 36);
            this.btn_c102.Name = "btn_c102";
            this.btn_c102.Size = new System.Drawing.Size(75, 23);
            this.btn_c102.TabIndex = 4;
            this.btn_c102.Text = "C102";
            this.btn_c102.UseVisualStyleBackColor = true;
            this.btn_c102.Click += new System.EventHandler(this.btn_c102_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(112, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(28, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "ACC";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(25, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "A D";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // btn_a305
            // 
            this.btn_a305.Location = new System.Drawing.Point(6, 72);
            this.btn_a305.Name = "btn_a305";
            this.btn_a305.Size = new System.Drawing.Size(75, 23);
            this.btn_a305.TabIndex = 1;
            this.btn_a305.Text = "A305";
            this.btn_a305.UseVisualStyleBackColor = true;
            this.btn_a305.Click += new System.EventHandler(this.btn_a305_Click);
            // 
            // btn_a302
            // 
            this.btn_a302.Location = new System.Drawing.Point(6, 36);
            this.btn_a302.Name = "btn_a302";
            this.btn_a302.Size = new System.Drawing.Size(75, 23);
            this.btn_a302.TabIndex = 0;
            this.btn_a302.Text = "A302";
            this.btn_a302.UseVisualStyleBackColor = true;
            this.btn_a302.Click += new System.EventHandler(this.btn_a302_Click);
            // 
            // FrmXML_Respuestas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 503);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.listBoxProceso);
            this.Controls.Add(this.listBoxArea);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmXML_Respuestas";
            this.Text = "Generador XML Respuestas";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmXML_Respuestas_FormClosing);
            this.Load += new System.EventHandler(this.FrmXML_Respuestas_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.ToolStripMenuItem cargaXMLToolStripMenuItem;
        private System.Windows.Forms.ListBox listBoxProceso;
        private System.Windows.Forms.ListBox listBoxArea;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_a305;
        private System.Windows.Forms.Button btn_a302;
        private System.Windows.Forms.DataGridViewTextBoxColumn id;
        private System.Windows.Forms.DataGridViewTextBoxColumn area;
        private System.Windows.Forms.DataGridViewTextBoxColumn proceso;
        private System.Windows.Forms.DataGridViewTextBoxColumn paso;
        private System.Windows.Forms.DataGridViewTextBoxColumn descripcion;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_inicio;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha_fin;
        private System.Windows.Forms.DataGridViewTextBoxColumn comentario;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_b102;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_C105;
        private System.Windows.Forms.Button btn_c102;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_m105;
        private System.Windows.Forms.Button btn_m102;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btn_b105;
        private System.Windows.Forms.Button btn_c205;
        private System.Windows.Forms.Button btn_c202;
        private System.Windows.Forms.Label label5;
    }
}