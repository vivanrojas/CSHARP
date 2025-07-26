namespace GestionOperaciones.forms.contratacion
{
            partial class FrmXML_Respuestas_Datos_C102R
        {
            /// <summary>
            /// Required designer variable.
            /// </summary>
            private System.ComponentModel.IContainer components = null;

            private System.Windows.Forms.ErrorProvider errorProvider;
            private System.Windows.Forms.Button btnCancel;
            private System.Windows.Forms.Button btnOK;
            private System.Windows.Forms.GroupBox groupBox2;
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

            private void InitializeComponent()
            {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmXML_Respuestas_Datos_C102R));
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txt_secuencial = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
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
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
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
            this.btnCancel.Location = new System.Drawing.Point(426, 376);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 41;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(575, 376);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 40;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txt_secuencial);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.textComentario);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.cmbMotivoRechazo);
            this.groupBox2.Controls.Add(this.maskedTextFechaRechazo);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Location = new System.Drawing.Point(12, 127);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(703, 233);
            this.groupBox2.TabIndex = 49;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Datos rechazo";
            // 
            // txt_secuencial
            // 
            this.txt_secuencial.Location = new System.Drawing.Point(95, 56);
            this.txt_secuencial.Name = "txt_secuencial";
            this.txt_secuencial.Size = new System.Drawing.Size(39, 20);
            this.txt_secuencial.TabIndex = 53;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 52;
            this.label1.Text = "Secuencial";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 127);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 13);
            this.label5.TabIndex = 51;
            this.label5.Text = "Comentarios:";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // textComentario
            // 
            this.textComentario.Location = new System.Drawing.Point(95, 127);
            this.textComentario.Multiline = true;
            this.textComentario.Name = "textComentario";
            this.textComentario.Size = new System.Drawing.Size(528, 89);
            this.textComentario.TabIndex = 50;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 97);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 13);
            this.label4.TabIndex = 49;
            this.label4.Text = "Motivo rechazo:";
            // 
            // cmbMotivoRechazo
            // 
            this.cmbMotivoRechazo.FormattingEnabled = true;
            this.cmbMotivoRechazo.Location = new System.Drawing.Point(95, 89);
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
            this.maskedTextFechaRechazo.Size = new System.Drawing.Size(74, 20);
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
            this.gbDatosSolicitud.Size = new System.Drawing.Size(697, 50);
            this.gbDatosSolicitud.TabIndex = 50;
            this.gbDatosSolicitud.TabStop = false;
            this.gbDatosSolicitud.Text = "Datos Solicitud  Rechazo";
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
            this.txtboxCodigoSolicitud.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
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
            this.txtboxCUPS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.txtboxCUPS.Location = new System.Drawing.Point(414, 34);
            this.txtboxCUPS.Margin = new System.Windows.Forms.Padding(2);
            this.txtboxCUPS.Name = "txtboxCUPS";
            this.txtboxCUPS.ReadOnly = true;
            this.txtboxCUPS.Size = new System.Drawing.Size(209, 13);
            this.txtboxCUPS.TabIndex = 0;
            this.txtboxCUPS.Text = "ES00000000000000";
            // 
            // FrmXML_Respuestas_Datos_C102R
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(727, 643);
            this.Controls.Add(this.gbDatosSolicitud);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmXML_Respuestas_Datos_C102R";
            this.Text = "Datos C102   Rechazo  CambiodeComercializadorSinCambios ";
            this.Load += new System.EventHandler(this.FrmXML_Respuestas_Fecha_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.gbDatosSolicitud.ResumeLayout(false);
            this.gbDatosSolicitud.PerformLayout();
            this.ResumeLayout(false);

            }

        #endregion

        private System.Windows.Forms.TextBox txt_secuencial;
        private System.Windows.Forms.Label label1;
    }
    }
