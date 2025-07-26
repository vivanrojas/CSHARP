namespace GestionOperaciones.forms.contratacion
{
    partial class FrmDirecciones
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmDirecciones));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btnSave = new System.Windows.Forms.Button();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.solicitud = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.usuario = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nif = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.razon_social = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ccups = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dir1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dir2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dir3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dir4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.calle = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.numero = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.poblacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.provincia = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.codpostal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.ItemSize = new System.Drawing.Size(10, 18);
            this.tabControl1.Location = new System.Drawing.Point(10, 87);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1146, 329);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btnSave);
            this.tabPage1.Controls.Add(this.dgv);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1138, 303);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Resumen";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Image = global::GestionOperaciones.Properties.Resources.iconfinder_floppy_285657_3_;
            this.btnSave.Location = new System.Drawing.Point(6, 11);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(28, 28);
            this.btnSave.TabIndex = 44;
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
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
            this.solicitud,
            this.usuario,
            this.nif,
            this.razon_social,
            this.ccups,
            this.dir1,
            this.dir2,
            this.dir3,
            this.dir4,
            this.calle,
            this.numero,
            this.poblacion,
            this.provincia,
            this.codpostal});
            this.dgv.Location = new System.Drawing.Point(6, 45);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(1126, 252);
            this.dgv.TabIndex = 0;
            this.dgv.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.Dgv_CellDoubleClick);
            this.dgv.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.Dgv_CellEnter);
            this.dgv.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.Dgv_CellMouseClick);
            this.dgv.CellMouseLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.Dgv_CellMouseLeave);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(634, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "AVISO: Las direcciones que se muestran, serán las utilizadas para la carátrula de" +
    "l burofax.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(17, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(733, 16);
            this.label2.TabIndex = 6;
            this.label2.Text = "Según el tipo de XML, se podrá utilizar las direcciones en las líneas o bien las " +
    "direcciones normalizadas.";
            // 
            // solicitud
            // 
            this.solicitud.DataPropertyName = "codigoDeSolicitud";
            this.solicitud.HeaderText = "Solicitud";
            this.solicitud.Name = "solicitud";
            this.solicitud.Visible = false;
            // 
            // usuario
            // 
            this.usuario.HeaderText = "Usuario";
            this.usuario.Name = "usuario";
            this.usuario.Visible = false;
            // 
            // nif
            // 
            this.nif.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.nif.DataPropertyName = "identificador";
            this.nif.HeaderText = "NIF";
            this.nif.Name = "nif";
            this.nif.ReadOnly = true;
            this.nif.Width = 49;
            // 
            // razon_social
            // 
            this.razon_social.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.razon_social.DataPropertyName = "razonSocial";
            this.razon_social.HeaderText = "RAZÓN SOCIAL";
            this.razon_social.Name = "razon_social";
            this.razon_social.ReadOnly = true;
            this.razon_social.Width = 102;
            // 
            // ccups
            // 
            this.ccups.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ccups.DataPropertyName = "cups";
            this.ccups.HeaderText = "CUPS";
            this.ccups.Name = "ccups";
            this.ccups.ReadOnly = true;
            this.ccups.Width = 61;
            // 
            // dir1
            // 
            this.dir1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dir1.DataPropertyName = "linea1DeLaDireccionExterna";
            this.dir1.HeaderText = "Dir. Ext. Línea 1";
            this.dir1.Name = "dir1";
            this.dir1.Width = 92;
            // 
            // dir2
            // 
            this.dir2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dir2.DataPropertyName = "linea2DeLaDireccionExterna";
            this.dir2.HeaderText = "Dir. Ext. Línea 2";
            this.dir2.Name = "dir2";
            this.dir2.Width = 92;
            // 
            // dir3
            // 
            this.dir3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dir3.DataPropertyName = "linea3DeLaDireccionExterna";
            this.dir3.HeaderText = "Dir. Ext. Línea 3";
            this.dir3.Name = "dir3";
            this.dir3.Width = 92;
            // 
            // dir4
            // 
            this.dir4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dir4.DataPropertyName = "linea4DeLaDireccionExterna";
            this.dir4.HeaderText = "Dir. Ext. Línea 4";
            this.dir4.Name = "dir4";
            this.dir4.Width = 92;
            // 
            // calle
            // 
            this.calle.DataPropertyName = "calleCliente";
            this.calle.HeaderText = "Calle";
            this.calle.Name = "calle";
            this.calle.Visible = false;
            // 
            // numero
            // 
            this.numero.HeaderText = "Número";
            this.numero.Name = "numero";
            this.numero.Visible = false;
            // 
            // poblacion
            // 
            this.poblacion.DataPropertyName = "descripcionPoblacionCliente";
            this.poblacion.HeaderText = "Población";
            this.poblacion.Name = "poblacion";
            this.poblacion.Visible = false;
            // 
            // provincia
            // 
            this.provincia.DataPropertyName = "provinciaCliente";
            this.provincia.HeaderText = "Provincia";
            this.provincia.Name = "provincia";
            this.provincia.Visible = false;
            // 
            // codpostal
            // 
            this.codpostal.DataPropertyName = "codPostalCliente";
            this.codpostal.HeaderText = "Código Postal";
            this.codpostal.Name = "codpostal";
            this.codpostal.Visible = false;
            // 
            // FrmDirecciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1159, 450);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmDirecciones";
            this.Text = "Edición de direcciones";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmDirecciones_FormClosed);
            this.Load += new System.EventHandler(this.FrmDirecciones_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn solicitud;
        private System.Windows.Forms.DataGridViewTextBoxColumn usuario;
        private System.Windows.Forms.DataGridViewTextBoxColumn nif;
        private System.Windows.Forms.DataGridViewTextBoxColumn razon_social;
        private System.Windows.Forms.DataGridViewTextBoxColumn ccups;
        private System.Windows.Forms.DataGridViewTextBoxColumn dir1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dir2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dir3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dir4;
        private System.Windows.Forms.DataGridViewTextBoxColumn calle;
        private System.Windows.Forms.DataGridViewTextBoxColumn numero;
        private System.Windows.Forms.DataGridViewTextBoxColumn poblacion;
        private System.Windows.Forms.DataGridViewTextBoxColumn provincia;
        private System.Windows.Forms.DataGridViewTextBoxColumn codpostal;
    }
}