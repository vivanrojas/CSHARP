namespace GestionOperaciones.forms.contratacion.gas
{
    partial class FrmExcepcion_Edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmExcepcion_Edit));
            this.groupBox_edit_excepcion = new System.Windows.Forms.GroupBox();
            this.dateTimePicker_fh = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_fd = new System.Windows.Forms.DateTimePicker();
            this.cmb_lista_distribuidoras = new System.Windows.Forms.ComboBox();
            this.label_fh = new System.Windows.Forms.Label();
            this.label_fd = new System.Windows.Forms.Label();
            this.label_nombre_distribuidora = new System.Windows.Forms.Label();
            this.btn_excepcion_Cancel = new System.Windows.Forms.Button();
            this.btn_excepcion_OK = new System.Windows.Forms.Button();
            this.errorProvider_frmExcepcion = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupBox_edit_excepcion.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider_frmExcepcion)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox_edit_excepcion
            // 
            this.groupBox_edit_excepcion.Controls.Add(this.dateTimePicker_fh);
            this.groupBox_edit_excepcion.Controls.Add(this.dateTimePicker_fd);
            this.groupBox_edit_excepcion.Controls.Add(this.cmb_lista_distribuidoras);
            this.groupBox_edit_excepcion.Controls.Add(this.label_fh);
            this.groupBox_edit_excepcion.Controls.Add(this.label_fd);
            this.groupBox_edit_excepcion.Controls.Add(this.label_nombre_distribuidora);
            this.groupBox_edit_excepcion.Location = new System.Drawing.Point(13, 14);
            this.groupBox_edit_excepcion.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox_edit_excepcion.Name = "groupBox_edit_excepcion";
            this.groupBox_edit_excepcion.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox_edit_excepcion.Size = new System.Drawing.Size(520, 165);
            this.groupBox_edit_excepcion.TabIndex = 13;
            this.groupBox_edit_excepcion.TabStop = false;
            this.groupBox_edit_excepcion.Text = "Datos excepción";
            // 
            // dateTimePicker_fh
            // 
            this.dateTimePicker_fh.CustomFormat = "dd/MM/yyyy HH:mm";
            this.dateTimePicker_fh.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker_fh.Location = new System.Drawing.Point(186, 122);
            this.dateTimePicker_fh.Name = "dateTimePicker_fh";
            this.dateTimePicker_fh.Size = new System.Drawing.Size(194, 26);
            this.dateTimePicker_fh.TabIndex = 20;
            this.dateTimePicker_fh.Value = new System.DateTime(2024, 3, 5, 12, 34, 15, 0);
            // 
            // dateTimePicker_fd
            // 
            this.dateTimePicker_fd.CustomFormat = "dd/MM/yyyy HH:mm";
            this.dateTimePicker_fd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker_fd.Location = new System.Drawing.Point(186, 82);
            this.dateTimePicker_fd.Name = "dateTimePicker_fd";
            this.dateTimePicker_fd.Size = new System.Drawing.Size(194, 26);
            this.dateTimePicker_fd.TabIndex = 19;
            this.dateTimePicker_fd.Value = new System.DateTime(2024, 3, 5, 12, 34, 15, 0);
            // 
            // cmb_lista_distribuidoras
            // 
            this.cmb_lista_distribuidoras.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_lista_distribuidoras.FormattingEnabled = true;
            this.cmb_lista_distribuidoras.Location = new System.Drawing.Point(186, 42);
            this.cmb_lista_distribuidoras.Name = "cmb_lista_distribuidoras";
            this.cmb_lista_distribuidoras.Size = new System.Drawing.Size(294, 28);
            this.cmb_lista_distribuidoras.TabIndex = 18;
            // 
            // label_fh
            // 
            this.label_fh.AutoSize = true;
            this.label_fh.Location = new System.Drawing.Point(73, 122);
            this.label_fh.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_fh.Name = "label_fh";
            this.label_fh.Size = new System.Drawing.Size(102, 20);
            this.label_fh.TabIndex = 17;
            this.label_fh.Text = "Fecha hasta:";
            // 
            // label_fd
            // 
            this.label_fd.AutoSize = true;
            this.label_fd.Location = new System.Drawing.Point(69, 82);
            this.label_fd.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_fd.Name = "label_fd";
            this.label_fd.Size = new System.Drawing.Size(106, 20);
            this.label_fd.TabIndex = 16;
            this.label_fd.Text = "Fecha desde:";
            // 
            // label_nombre_distribuidora
            // 
            this.label_nombre_distribuidora.AutoSize = true;
            this.label_nombre_distribuidora.Location = new System.Drawing.Point(16, 45);
            this.label_nombre_distribuidora.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_nombre_distribuidora.Name = "label_nombre_distribuidora";
            this.label_nombre_distribuidora.Size = new System.Drawing.Size(159, 20);
            this.label_nombre_distribuidora.TabIndex = 15;
            this.label_nombre_distribuidora.Text = "Nombre distribuidora:";
            // 
            // btn_excepcion_Cancel
            // 
            this.btn_excepcion_Cancel.Location = new System.Drawing.Point(421, 200);
            this.btn_excepcion_Cancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_excepcion_Cancel.Name = "btn_excepcion_Cancel";
            this.btn_excepcion_Cancel.Size = new System.Drawing.Size(112, 35);
            this.btn_excepcion_Cancel.TabIndex = 15;
            this.btn_excepcion_Cancel.Text = "Cancelar";
            this.btn_excepcion_Cancel.UseVisualStyleBackColor = true;
            this.btn_excepcion_Cancel.Click += new System.EventHandler(this.btn_excepcion_Cancel_Click);
            // 
            // btn_excepcion_OK
            // 
            this.btn_excepcion_OK.Location = new System.Drawing.Point(299, 200);
            this.btn_excepcion_OK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_excepcion_OK.Name = "btn_excepcion_OK";
            this.btn_excepcion_OK.Size = new System.Drawing.Size(112, 35);
            this.btn_excepcion_OK.TabIndex = 14;
            this.btn_excepcion_OK.Text = "Aceptar";
            this.btn_excepcion_OK.UseVisualStyleBackColor = true;
            this.btn_excepcion_OK.Click += new System.EventHandler(this.btn_excepcion_OK_Click);
            // 
            // errorProvider_frmExcepcion
            // 
            this.errorProvider_frmExcepcion.ContainerControl = this;
            // 
            // FrmExcepcion_Edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.ClientSize = new System.Drawing.Size(568, 270);
            this.Controls.Add(this.btn_excepcion_Cancel);
            this.Controls.Add(this.btn_excepcion_OK);
            this.Controls.Add(this.groupBox_edit_excepcion);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmExcepcion_Edit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FrmExcepcion_Edit";
            this.groupBox_edit_excepcion.ResumeLayout(false);
            this.groupBox_edit_excepcion.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider_frmExcepcion)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox_edit_excepcion;
        private System.Windows.Forms.Label label_fh;
        private System.Windows.Forms.Label label_fd;
        private System.Windows.Forms.Label label_nombre_distribuidora;
        private System.Windows.Forms.Button btn_excepcion_Cancel;
        private System.Windows.Forms.Button btn_excepcion_OK;
        public System.Windows.Forms.ComboBox cmb_lista_distribuidoras;
        public System.Windows.Forms.DateTimePicker dateTimePicker_fd;
        public System.Windows.Forms.DateTimePicker dateTimePicker_fh;
        private System.Windows.Forms.ErrorProvider errorProvider_frmExcepcion;
    }
}