namespace GestionOperaciones.forms.contratacion
{
    partial class FrmXML_Respuestas_Datos_B105
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmXML_Respuestas_Datos_B105));
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.gbDatosSolicitud = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtboxCodigoSolicitud = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtboxCUPS = new System.Windows.Forms.TextBox();
            this.txt_fecha = new System.Windows.Forms.MaskedTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_codcontrato = new System.Windows.Forms.MaskedTextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Txt_CodPM = new System.Windows.Forms.MaskedTextBox();
            this.cmb_Funcion = new System.Windows.Forms.ComboBox();
            this.Txt_FechaAlta = new System.Windows.Forms.MaskedTextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.Txt_FehaVigor = new System.Windows.Forms.MaskedTextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cmd_Modolectura = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cmb_TipoMovimiento = new System.Windows.Forms.ComboBox();
            this.cmb_tipoPM = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.gbDatosSolicitud.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(454, 380);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 41;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(586, 380);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 40;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // gbDatosSolicitud
            // 
            this.gbDatosSolicitud.Controls.Add(this.label7);
            this.gbDatosSolicitud.Controls.Add(this.txtboxCodigoSolicitud);
            this.gbDatosSolicitud.Controls.Add(this.label6);
            this.gbDatosSolicitud.Controls.Add(this.txtboxCUPS);
            this.gbDatosSolicitud.Location = new System.Drawing.Point(25, 24);
            this.gbDatosSolicitud.Margin = new System.Windows.Forms.Padding(2);
            this.gbDatosSolicitud.Name = "gbDatosSolicitud";
            this.gbDatosSolicitud.Padding = new System.Windows.Forms.Padding(2);
            this.gbDatosSolicitud.Size = new System.Drawing.Size(686, 49);
            this.gbDatosSolicitud.TabIndex = 51;
            this.gbDatosSolicitud.TabStop = false;
            this.gbDatosSolicitud.Text = "Datos Activacion Baja";
            this.gbDatosSolicitud.Enter += new System.EventHandler(this.gbDatosSolicitud_Enter);
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
            // txt_fecha
            // 
            this.txt_fecha.Location = new System.Drawing.Point(173, 98);
            this.txt_fecha.Mask = "00/00/0000";
            this.txt_fecha.Name = "txt_fecha";
            this.txt_fecha.Size = new System.Drawing.Size(78, 20);
            this.txt_fecha.TabIndex = 52;
            this.txt_fecha.ValidatingType = typeof(System.DateTime);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(49, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 53;
            this.label2.Text = "Fecha  Activacion";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(305, 104);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 13);
            this.label4.TabIndex = 59;
            this.label4.Text = "CodContrato";
            // 
            // txt_codcontrato
            // 
            this.txt_codcontrato.Location = new System.Drawing.Point(383, 101);
            this.txt_codcontrato.Name = "txt_codcontrato";
            this.txt_codcontrato.Size = new System.Drawing.Size(121, 20);
            this.txt_codcontrato.TabIndex = 60;
            this.txt_codcontrato.ValidatingType = typeof(System.DateTime);
            this.txt_codcontrato.MaskInputRejected += new System.Windows.Forms.MaskInputRejectedEventHandler(this.txt_codcontrato_MaskInputRejected);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Txt_CodPM);
            this.groupBox1.Controls.Add(this.cmb_Funcion);
            this.groupBox1.Controls.Add(this.Txt_FechaAlta);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.Txt_FehaVigor);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.cmd_Modolectura);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.cmb_TipoMovimiento);
            this.groupBox1.Controls.Add(this.cmb_tipoPM);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Location = new System.Drawing.Point(25, 144);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(697, 199);
            this.groupBox1.TabIndex = 61;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Punto de Medida";
            // 
            // Txt_CodPM
            // 
            this.Txt_CodPM.Location = new System.Drawing.Point(88, 19);
            this.Txt_CodPM.Name = "Txt_CodPM";
            this.Txt_CodPM.Size = new System.Drawing.Size(163, 20);
            this.Txt_CodPM.TabIndex = 67;
            this.Txt_CodPM.ValidatingType = typeof(System.DateTime);
            // 
            // cmb_Funcion
            // 
            this.cmb_Funcion.ForeColor = System.Drawing.Color.Black;
            this.cmb_Funcion.FormattingEnabled = true;
            this.cmb_Funcion.Items.AddRange(new object[] {
            "C - Comprobante",
            "P - Principal",
            "R - Redundante"});
            this.cmb_Funcion.Location = new System.Drawing.Point(561, 56);
            this.cmb_Funcion.Name = "cmb_Funcion";
            this.cmb_Funcion.Size = new System.Drawing.Size(125, 21);
            this.cmb_Funcion.TabIndex = 66;
            // 
            // Txt_FechaAlta
            // 
            this.Txt_FechaAlta.Location = new System.Drawing.Point(372, 91);
            this.Txt_FechaAlta.Mask = "00/00/0000";
            this.Txt_FechaAlta.Name = "Txt_FechaAlta";
            this.Txt_FechaAlta.Size = new System.Drawing.Size(78, 20);
            this.Txt_FechaAlta.TabIndex = 65;
            this.Txt_FechaAlta.ValidatingType = typeof(System.DateTime);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(284, 98);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(58, 13);
            this.label13.TabIndex = 64;
            this.label13.Text = "Fecha Alta";
            // 
            // Txt_FehaVigor
            // 
            this.Txt_FehaVigor.Location = new System.Drawing.Point(88, 87);
            this.Txt_FehaVigor.Mask = "00/00/0000";
            this.Txt_FehaVigor.Name = "Txt_FehaVigor";
            this.Txt_FehaVigor.Size = new System.Drawing.Size(73, 20);
            this.Txt_FehaVigor.TabIndex = 63;
            this.Txt_FehaVigor.ValidatingType = typeof(System.DateTime);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(24, 94);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(61, 13);
            this.label12.TabIndex = 62;
            this.label12.Text = "FechaVigor";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(516, 64);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(45, 13);
            this.label11.TabIndex = 61;
            this.label11.Text = "Funcion";
            // 
            // cmd_Modolectura
            // 
            this.cmd_Modolectura.ForeColor = System.Drawing.Color.Black;
            this.cmd_Modolectura.FormattingEnabled = true;
            this.cmd_Modolectura.Items.AddRange(new object[] {
            "1 - Lectura local manual",
            "2 - Lectura ñocal optoacoplador",
            "3 - Lectura local puerto serie",
            "4 - Telemedida operativa",
            "5 - Telemedida no operativa"});
            this.cmd_Modolectura.Location = new System.Drawing.Point(351, 56);
            this.cmd_Modolectura.Name = "cmd_Modolectura";
            this.cmd_Modolectura.Size = new System.Drawing.Size(153, 21);
            this.cmd_Modolectura.TabIndex = 60;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(265, 64);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(70, 13);
            this.label10.TabIndex = 59;
            this.label10.Text = "ModoLectura";
            // 
            // cmb_TipoMovimiento
            // 
            this.cmb_TipoMovimiento.ForeColor = System.Drawing.Color.Black;
            this.cmb_TipoMovimiento.FormattingEnabled = true;
            this.cmb_TipoMovimiento.Items.AddRange(new object[] {
            "A - Alta",
            "B - Bajao",
            "M - Modificacion"});
            this.cmb_TipoMovimiento.Location = new System.Drawing.Point(351, 18);
            this.cmb_TipoMovimiento.Name = "cmb_TipoMovimiento";
            this.cmb_TipoMovimiento.Size = new System.Drawing.Size(189, 21);
            this.cmb_TipoMovimiento.TabIndex = 58;
            this.cmb_TipoMovimiento.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // cmb_tipoPM
            // 
            this.cmb_tipoPM.ForeColor = System.Drawing.Color.Black;
            this.cmb_tipoPM.FormattingEnabled = true;
            this.cmb_tipoPM.Items.AddRange(new object[] {
            "01 - Punto de medida tipo 1",
            "02 - Punto de medida tipo 2",
            "03 - Punto de medida tipo 3",
            "04 - Punto de medida tipo 4",
            "05 - Punto de medida tipo 5"});
            this.cmb_tipoPM.Location = new System.Drawing.Point(88, 56);
            this.cmb_tipoPM.Name = "cmb_tipoPM";
            this.cmb_tipoPM.Size = new System.Drawing.Size(163, 21);
            this.cmb_tipoPM.TabIndex = 57;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(24, 64);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 13);
            this.label9.TabIndex = 56;
            this.label9.Text = "TipoPM";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 26);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(45, 13);
            this.label5.TabIndex = 3;
            this.label5.Text = "Cod PM";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(264, 25);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(82, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "TipoMovimiento";
            // 
            // FrmXML_Respuestas_Datos_B105
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(756, 698);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txt_codcontrato);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txt_fecha);
            this.Controls.Add(this.gbDatosSolicitud);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmXML_Respuestas_Datos_B105";
            this.Text = "Activación Baja Suspension B105 ";
            this.Load += new System.EventHandler(this.FrmXML_Respuestas_Datos_A305_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.gbDatosSolicitud.ResumeLayout(false);
            this.gbDatosSolicitud.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.GroupBox gbDatosSolicitud;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtboxCodigoSolicitud;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtboxCUPS;
        private System.Windows.Forms.MaskedTextBox txt_fecha;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MaskedTextBox txt_codcontrato;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.ComboBox cmb_TipoMovimiento;
        public System.Windows.Forms.ComboBox cmb_tipoPM;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.ComboBox cmb_Funcion;
        private System.Windows.Forms.MaskedTextBox Txt_FechaAlta;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.MaskedTextBox Txt_FehaVigor;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        public System.Windows.Forms.ComboBox cmd_Modolectura;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.MaskedTextBox Txt_CodPM;
    }
}