namespace GestionOperaciones.forms.contratacion
{
    partial class FrmXML_Respuestas_Datos_M102
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmXML_Respuestas_Datos_M102));
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.cmb_tipo_activacion = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.chk_aceptacion = new System.Windows.Forms.CheckBox();
            this.chk_rechazo = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textComentario = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbMotivoRechazo = new System.Windows.Forms.ComboBox();
            this.maskedTextFechaRechazo = new System.Windows.Forms.MaskedTextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.gbDatosSolicitud = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtboxCodigoSolicitud = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtboxCUPS = new System.Windows.Forms.TextBox();
            this.chk_actuacion_campo = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_FechaAceptacion = new System.Windows.Forms.MaskedTextBox();
            this.txt_FechaActivacionPrevista = new System.Windows.Forms.MaskedTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.gbDatosSolicitud.SuspendLayout();
            this.SuspendLayout();
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(543, 497);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 41;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(625, 497);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 40;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // cmb_tipo_activacion
            // 
            this.cmb_tipo_activacion.FormattingEnabled = true;
            this.cmb_tipo_activacion.Location = new System.Drawing.Point(160, 54);
            this.cmb_tipo_activacion.Name = "cmb_tipo_activacion";
            this.cmb_tipo_activacion.Size = new System.Drawing.Size(285, 21);
            this.cmb_tipo_activacion.TabIndex = 42;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 58);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 13);
            this.label1.TabIndex = 43;
            this.label1.Text = "Tipo activación prevista:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt_FechaActivacionPrevista);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.chk_actuacion_campo);
            this.groupBox1.Controls.Add(this.txt_FechaAceptacion);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cmb_tipo_activacion);
            this.groupBox1.Location = new System.Drawing.Point(12, 125);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(697, 118);
            this.groupBox1.TabIndex = 44;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Datos aceptación";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 84);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 13);
            this.label2.TabIndex = 44;
            this.label2.Text = "FechaAceptacion";
            // 
            // chk_aceptacion
            // 
            this.chk_aceptacion.AutoSize = true;
            this.chk_aceptacion.Location = new System.Drawing.Point(12, 86);
            this.chk_aceptacion.Name = "chk_aceptacion";
            this.chk_aceptacion.Size = new System.Drawing.Size(176, 17);
            this.chk_aceptacion.TabIndex = 47;
            this.chk_aceptacion.Text = "AceptacionModificacionDeATR";
            this.chk_aceptacion.UseVisualStyleBackColor = true;
            this.chk_aceptacion.CheckedChanged += new System.EventHandler(this.chk_aceptacion_CheckedChanged);
            // 
            // chk_rechazo
            // 
            this.chk_rechazo.AutoSize = true;
            this.chk_rechazo.Location = new System.Drawing.Point(12, 263);
            this.chk_rechazo.Name = "chk_rechazo";
            this.chk_rechazo.Size = new System.Drawing.Size(69, 17);
            this.chk_rechazo.TabIndex = 48;
            this.chk_rechazo.Text = "Rechazo";
            this.chk_rechazo.UseVisualStyleBackColor = true;
            this.chk_rechazo.CheckedChanged += new System.EventHandler(this.chk_rechazo_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.textComentario);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.cmbMotivoRechazo);
            this.groupBox2.Controls.Add(this.maskedTextFechaRechazo);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Location = new System.Drawing.Point(12, 287);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(703, 183);
            this.groupBox2.TabIndex = 49;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Datos rechazo";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 82);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 13);
            this.label5.TabIndex = 51;
            this.label5.Text = "Comentarios:";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // textComentario
            // 
            this.textComentario.Location = new System.Drawing.Point(90, 79);
            this.textComentario.Multiline = true;
            this.textComentario.Name = "textComentario";
            this.textComentario.Size = new System.Drawing.Size(537, 89);
            this.textComentario.TabIndex = 50;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 13);
            this.label4.TabIndex = 49;
            this.label4.Text = "Motivo rechazo:";
            // 
            // cmbMotivoRechazo
            // 
            this.cmbMotivoRechazo.FormattingEnabled = true;
            this.cmbMotivoRechazo.Location = new System.Drawing.Point(90, 52);
            this.cmbMotivoRechazo.Name = "cmbMotivoRechazo";
            this.cmbMotivoRechazo.Size = new System.Drawing.Size(285, 21);
            this.cmbMotivoRechazo.TabIndex = 48;
            this.cmbMotivoRechazo.SelectedIndexChanged += new System.EventHandler(this.cmbMotivoRechazo_SelectedIndexChanged);
            // 
            // maskedTextFechaRechazo
            // 
            this.maskedTextFechaRechazo.Location = new System.Drawing.Point(95, 24);
            this.maskedTextFechaRechazo.Mask = "00/00/0000";
            this.maskedTextFechaRechazo.Name = "maskedTextFechaRechazo";
            this.maskedTextFechaRechazo.Size = new System.Drawing.Size(81, 20);
            this.maskedTextFechaRechazo.TabIndex = 47;
            this.maskedTextFechaRechazo.ValidatingType = typeof(System.DateTime);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 13);
            this.label3.TabIndex = 46;
            this.label3.Text = "Fecha rechazo:";
            // 
            // gbDatosSolicitud
            // 
            this.gbDatosSolicitud.Controls.Add(this.label7);
            this.gbDatosSolicitud.Controls.Add(this.txtboxCodigoSolicitud);
            this.gbDatosSolicitud.Controls.Add(this.label6);
            this.gbDatosSolicitud.Controls.Add(this.txtboxCUPS);
            this.gbDatosSolicitud.Location = new System.Drawing.Point(12, 19);
            this.gbDatosSolicitud.Margin = new System.Windows.Forms.Padding(2);
            this.gbDatosSolicitud.Name = "gbDatosSolicitud";
            this.gbDatosSolicitud.Padding = new System.Windows.Forms.Padding(2);
            this.gbDatosSolicitud.Size = new System.Drawing.Size(697, 47);
            this.gbDatosSolicitud.TabIndex = 50;
            this.gbDatosSolicitud.TabStop = false;
            this.gbDatosSolicitud.Text = "Datos Solicitud";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(59, 34);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(84, 13);
            this.label7.TabIndex = 3;
            this.label7.Text = "Código solicitud:";
            // 
            // txtboxCodigoSolicitud
            // 
            this.txtboxCodigoSolicitud.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtboxCodigoSolicitud.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxCodigoSolicitud.Location = new System.Drawing.Point(148, 34);
            this.txtboxCodigoSolicitud.Margin = new System.Windows.Forms.Padding(2);
            this.txtboxCodigoSolicitud.Name = "txtboxCodigoSolicitud";
            this.txtboxCodigoSolicitud.ReadOnly = true;
            this.txtboxCodigoSolicitud.Size = new System.Drawing.Size(173, 13);
            this.txtboxCodigoSolicitud.TabIndex = 2;
            this.txtboxCodigoSolicitud.Text = "XXXXXXXXXXX";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(369, 34);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "CUPS:";
            // 
            // txtboxCUPS
            // 
            this.txtboxCUPS.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtboxCUPS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxCUPS.Location = new System.Drawing.Point(414, 34);
            this.txtboxCUPS.Margin = new System.Windows.Forms.Padding(2);
            this.txtboxCUPS.Name = "txtboxCUPS";
            this.txtboxCUPS.ReadOnly = true;
            this.txtboxCUPS.Size = new System.Drawing.Size(209, 13);
            this.txtboxCUPS.TabIndex = 0;
            this.txtboxCUPS.Text = "ES00000000000000";
            // 
            // chk_actuacion_campo
            // 
            this.chk_actuacion_campo.AutoSize = true;
            this.chk_actuacion_campo.Location = new System.Drawing.Point(174, 19);
            this.chk_actuacion_campo.Name = "chk_actuacion_campo";
            this.chk_actuacion_campo.Size = new System.Drawing.Size(110, 17);
            this.chk_actuacion_campo.TabIndex = 47;
            this.chk_actuacion_campo.Text = "Actuación Campo";
            this.chk_actuacion_campo.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(303, 84);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(125, 13);
            this.label8.TabIndex = 48;
            this.label8.Text = "FechaActivacionPrevista";
            // 
            // txt_FechaAceptacion
            // 
            this.txt_FechaAceptacion.Location = new System.Drawing.Point(160, 84);
            this.txt_FechaAceptacion.Mask = "00/00/0000";
            this.txt_FechaAceptacion.Name = "txt_FechaAceptacion";
            this.txt_FechaAceptacion.Size = new System.Drawing.Size(89, 20);
            this.txt_FechaAceptacion.TabIndex = 45;
            this.txt_FechaAceptacion.ValidatingType = typeof(System.DateTime);
            // 
            // txt_FechaActivacionPrevista
            // 
            this.txt_FechaActivacionPrevista.Location = new System.Drawing.Point(452, 81);
            this.txt_FechaActivacionPrevista.Mask = "00/00/0000";
            this.txt_FechaActivacionPrevista.Name = "txt_FechaActivacionPrevista";
            this.txt_FechaActivacionPrevista.Size = new System.Drawing.Size(76, 20);
            this.txt_FechaActivacionPrevista.TabIndex = 49;
            this.txt_FechaActivacionPrevista.ValidatingType = typeof(System.DateTime);
            // 
            // FrmXML_Respuestas_Datos_M102
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 539);
            this.Controls.Add(this.gbDatosSolicitud);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.chk_rechazo);
            this.Controls.Add(this.chk_aceptacion);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmXML_Respuestas_Datos_M102";
            this.Text = "Datos M1 Aceptacion Modificacion De ATR";
            this.Load += new System.EventHandler(this.FrmXML_Respuestas_Fecha_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.gbDatosSolicitud.ResumeLayout(false);
            this.gbDatosSolicitud.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox cmb_tipo_activacion;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chk_rechazo;
        private System.Windows.Forms.CheckBox chk_aceptacion;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textComentario;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.ComboBox cmbMotivoRechazo;
        private System.Windows.Forms.MaskedTextBox maskedTextFechaRechazo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox gbDatosSolicitud;
        private System.Windows.Forms.TextBox txtboxCUPS;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtboxCodigoSolicitud;
        private System.Windows.Forms.MaskedTextBox txt_FechaActivacionPrevista;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckBox chk_actuacion_campo;
        private System.Windows.Forms.MaskedTextBox txt_FechaAceptacion;
    }
}