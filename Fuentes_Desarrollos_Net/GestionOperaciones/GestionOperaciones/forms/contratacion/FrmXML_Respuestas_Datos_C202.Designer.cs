namespace GestionOperaciones.forms.contratacion
{
            partial class FrmXML_Respuestas_Datos_C202
        {
            /// <summary>
            /// Required designer variable.
            /// </summary>
            private System.ComponentModel.IContainer components = null;

            private System.Windows.Forms.ErrorProvider errorProvider;
            private System.Windows.Forms.Button btnCancel;
            private System.Windows.Forms.Button btnOK;
            private System.Windows.Forms.GroupBox groupBox1;
            private System.Windows.Forms.Label label1;
            public System.Windows.Forms.ComboBox cmb_ActuacionCampo;
            private System.Windows.Forms.Label label2;
            private System.Windows.Forms.MaskedTextBox txt_fecha_Aceptacion;
            private System.Windows.Forms.CheckBox chk_rechazo;
            private System.Windows.Forms.CheckBox chk_aceptacion;
            private System.Windows.Forms.GroupBox gbDatosSolicitud;
            private System.Windows.Forms.TextBox txtboxCUPS;
            private System.Windows.Forms.Label label6;
            private System.Windows.Forms.Label label7;
            private System.Windows.Forms.TextBox txtboxCodigoSolicitud;
            private System.Windows.Forms.MaskedTextBox txt_FechaUltimaLecturaFirme;
            private System.Windows.Forms.Label label8;
            public System.Windows.Forms.ComboBox cmb_TipoContratoATR;
            private System.Windows.Forms.Label label9;
            private System.Windows.Forms.MaskedTextBox txt_FechaActivacionPrevista;
            private System.Windows.Forms.Label label11;
            public System.Windows.Forms.ComboBox cmb_TipoActivacionPrevista;
            private System.Windows.Forms.Label label10;
            private System.Windows.Forms.GroupBox groupBox3;
            public System.Windows.Forms.ComboBox cmb_TarifaATR;
            private System.Windows.Forms.Label label14;
            private System.Windows.Forms.DataGridView dgvPotencias;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmXML_Respuestas_Datos_C202));
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.cmb_ActuacionCampo = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txt_FechaActivacionPrevista = new System.Windows.Forms.MaskedTextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cmb_TipoActivacionPrevista = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cmb_TipoContratoATR = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_FechaUltimaLecturaFirme = new System.Windows.Forms.MaskedTextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_fecha_Aceptacion = new System.Windows.Forms.MaskedTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.chk_aceptacion = new System.Windows.Forms.CheckBox();
            this.chk_rechazo = new System.Windows.Forms.CheckBox();
            this.gbDatosSolicitud = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtboxCodigoSolicitud = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtboxCUPS = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.cmb_TarifaATR = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dgvPotencias = new System.Windows.Forms.DataGridView();
            this.Periodo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PotenciaW = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.gbDatosSolicitud.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPotencias)).BeginInit();
            this.SuspendLayout();
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(476, 485);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 41;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(606, 485);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 40;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // cmb_ActuacionCampo
            // 
            this.cmb_ActuacionCampo.FormattingEnabled = true;
            this.cmb_ActuacionCampo.Items.AddRange(new object[] {
            "S - Si",
            "N - No"});
            this.cmb_ActuacionCampo.Location = new System.Drawing.Point(393, 19);
            this.cmb_ActuacionCampo.Name = "cmb_ActuacionCampo";
            this.cmb_ActuacionCampo.Size = new System.Drawing.Size(146, 21);
            this.cmb_ActuacionCampo.TabIndex = 42;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(262, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 43;
            this.label1.Text = "Actuacion Campo:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt_FechaActivacionPrevista);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.cmb_TipoActivacionPrevista);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.cmb_TipoContratoATR);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.txt_FechaUltimaLecturaFirme);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txt_fecha_Aceptacion);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cmb_ActuacionCampo);
            this.groupBox1.Location = new System.Drawing.Point(12, 106);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(703, 147);
            this.groupBox1.TabIndex = 44;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Datos Aceptación Cambio de Comerciaizdor Con Cambios ";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // txt_FechaActivacionPrevista
            // 
            this.txt_FechaActivacionPrevista.Location = new System.Drawing.Point(393, 91);
            this.txt_FechaActivacionPrevista.Mask = "00/00/0000";
            this.txt_FechaActivacionPrevista.Name = "txt_FechaActivacionPrevista";
            this.txt_FechaActivacionPrevista.Size = new System.Drawing.Size(72, 20);
            this.txt_FechaActivacionPrevista.TabIndex = 54;
            this.txt_FechaActivacionPrevista.ValidatingType = typeof(System.DateTime);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(262, 98);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(125, 13);
            this.label11.TabIndex = 53;
            this.label11.Text = "FechaActivacionPrevista";
            // 
            // cmb_TipoActivacionPrevista
            // 
            this.cmb_TipoActivacionPrevista.FormattingEnabled = true;
            this.cmb_TipoActivacionPrevista.Location = new System.Drawing.Point(393, 59);
            this.cmb_TipoActivacionPrevista.Name = "cmb_TipoActivacionPrevista";
            this.cmb_TipoActivacionPrevista.Size = new System.Drawing.Size(146, 21);
            this.cmb_TipoActivacionPrevista.TabIndex = 52;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(259, 67);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(116, 13);
            this.label10.TabIndex = 51;
            this.label10.Text = "TipoActivacionPrevista";
            // 
            // cmb_TipoContratoATR
            // 
            this.cmb_TipoContratoATR.FormattingEnabled = true;
            this.cmb_TipoContratoATR.Location = new System.Drawing.Point(148, 59);
            this.cmb_TipoContratoATR.Name = "cmb_TipoContratoATR";
            this.cmb_TipoContratoATR.Size = new System.Drawing.Size(98, 21);
            this.cmb_TipoContratoATR.TabIndex = 50;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 72);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(90, 13);
            this.label9.TabIndex = 49;
            this.label9.Text = "TipoContratoATR";
            // 
            // txt_FechaUltimaLecturaFirme
            // 
            this.txt_FechaUltimaLecturaFirme.Location = new System.Drawing.Point(149, 95);
            this.txt_FechaUltimaLecturaFirme.Mask = "00/00/0000";
            this.txt_FechaUltimaLecturaFirme.Name = "txt_FechaUltimaLecturaFirme";
            this.txt_FechaUltimaLecturaFirme.Size = new System.Drawing.Size(78, 20);
            this.txt_FechaUltimaLecturaFirme.TabIndex = 48;
            this.txt_FechaUltimaLecturaFirme.ValidatingType = typeof(System.DateTime);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(16, 102);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(127, 13);
            this.label8.TabIndex = 47;
            this.label8.Text = "FechaUltimaLecturaFirme";
            // 
            // txt_fecha_Aceptacion
            // 
            this.txt_fecha_Aceptacion.Location = new System.Drawing.Point(149, 27);
            this.txt_fecha_Aceptacion.Mask = "00/00/0000";
            this.txt_fecha_Aceptacion.Name = "txt_fecha_Aceptacion";
            this.txt_fecha_Aceptacion.Size = new System.Drawing.Size(78, 20);
            this.txt_fecha_Aceptacion.TabIndex = 45;
            this.txt_fecha_Aceptacion.ValidatingType = typeof(System.DateTime);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(94, 13);
            this.label2.TabIndex = 44;
            this.label2.Text = "Fecha Aceptacion";
            // 
            // chk_aceptacion
            // 
            this.chk_aceptacion.AutoSize = true;
            this.chk_aceptacion.Location = new System.Drawing.Point(12, 83);
            this.chk_aceptacion.Name = "chk_aceptacion";
            this.chk_aceptacion.Size = new System.Drawing.Size(80, 17);
            this.chk_aceptacion.TabIndex = 47;
            this.chk_aceptacion.Text = "Aceptación";
            this.chk_aceptacion.UseVisualStyleBackColor = true;
            this.chk_aceptacion.CheckedChanged += new System.EventHandler(this.chk_aceptacion_CheckedChanged);
            // 
            // chk_rechazo
            // 
            this.chk_rechazo.AutoSize = true;
            this.chk_rechazo.Location = new System.Drawing.Point(161, 83);
            this.chk_rechazo.Name = "chk_rechazo";
            this.chk_rechazo.Size = new System.Drawing.Size(69, 17);
            this.chk_rechazo.TabIndex = 48;
            this.chk_rechazo.Text = "Rechazo";
            this.chk_rechazo.UseVisualStyleBackColor = true;
            this.chk_rechazo.CheckedChanged += new System.EventHandler(this.chk_rechazo_CheckedChanged);
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
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 27);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(56, 13);
            this.label14.TabIndex = 49;
            this.label14.Text = "TarifaATR";
            // 
            // cmb_TarifaATR
            // 
            this.cmb_TarifaATR.FormattingEnabled = true;
            this.cmb_TarifaATR.Location = new System.Drawing.Point(80, 19);
            this.cmb_TarifaATR.Name = "cmb_TarifaATR";
            this.cmb_TarifaATR.Size = new System.Drawing.Size(138, 21);
            this.cmb_TarifaATR.TabIndex = 50;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.cmb_TarifaATR);
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.dgvPotencias);
            this.groupBox3.Location = new System.Drawing.Point(12, 259);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(703, 183);
            this.groupBox3.TabIndex = 51;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Condiciones Contractuales";
            this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
            // 
            // dgvPotencias
            // 
            this.dgvPotencias.AllowUserToAddRows = false;
            this.dgvPotencias.AllowUserToDeleteRows = false;
            this.dgvPotencias.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPotencias.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Periodo,
            this.PotenciaW});
            this.dgvPotencias.Location = new System.Drawing.Point(233, 19);
            this.dgvPotencias.Name = "dgvPotencias";
            this.dgvPotencias.Size = new System.Drawing.Size(206, 157);
            this.dgvPotencias.TabIndex = 0;
            // 
            // Periodo
            // 
            this.Periodo.HeaderText = "Periodo";
            this.Periodo.Name = "Periodo";
            this.Periodo.ReadOnly = true;
            this.Periodo.Width = 60;
            // 
            // PotenciaW
            // 
            this.PotenciaW.HeaderText = "Potencia (W)";
            this.PotenciaW.Name = "PotenciaW";
            // 
            // FrmXML_Respuestas_Datos_C202
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(727, 643);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.gbDatosSolicitud);
            this.Controls.Add(this.chk_rechazo);
            this.Controls.Add(this.chk_aceptacion);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmXML_Respuestas_Datos_C202";
            this.Text = "Datos C202   Aceptacion  Cambio de Comercializador Con Cambios";
            this.Load += new System.EventHandler(this.FrmXML_Respuestas_Fecha_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.gbDatosSolicitud.ResumeLayout(false);
            this.gbDatosSolicitud.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPotencias)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

            }

        #endregion

        private System.Windows.Forms.DataGridViewTextBoxColumn Periodo;
        private System.Windows.Forms.DataGridViewTextBoxColumn PotenciaW;
    }
    }
