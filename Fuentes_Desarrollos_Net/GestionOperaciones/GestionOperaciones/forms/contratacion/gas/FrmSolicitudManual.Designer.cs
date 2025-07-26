
namespace GestionOperaciones.forms.contratacion.gas
{
    partial class FrmSolicitudManual
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSolicitudManual));
            this.lbl_cliente = new System.Windows.Forms.Label();
            this.lbl_nif = new System.Windows.Forms.Label();
            this.lbl_distribuidora = new System.Windows.Forms.Label();
            this.txt_nif = new System.Windows.Forms.TextBox();
            this.txt_cliente = new System.Windows.Forms.TextBox();
            this.txt_distribuidora = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txt_qi = new System.Windows.Forms.TextBox();
            this.txt_horaInicio = new System.Windows.Forms.MaskedTextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtCodIndProd = new System.Windows.Forms.TextBox();
            this.cmbTipoProducto = new System.Windows.Forms.ComboBox();
            this.chkSoloXML = new System.Windows.Forms.CheckBox();
            this.txt_fecha_fin = new System.Windows.Forms.DateTimePicker();
            this.txt_fecha_inicio = new System.Windows.Forms.DateTimePicker();
            this.chkSolapado = new System.Windows.Forms.CheckBox();
            this.txt_tarifa = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_comentario = new System.Windows.Forms.TextBox();
            this.cmb_tipo = new System.Windows.Forms.ComboBox();
            this.txt_qd = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_cups20 = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txt_tipotratamitacion = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.toolTip_Qi = new System.Windows.Forms.ToolTip(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // lbl_cliente
            // 
            this.lbl_cliente.AutoSize = true;
            this.lbl_cliente.Location = new System.Drawing.Point(62, 51);
            this.lbl_cliente.Name = "lbl_cliente";
            this.lbl_cliente.Size = new System.Drawing.Size(42, 13);
            this.lbl_cliente.TabIndex = 65;
            this.lbl_cliente.Text = "Cliente:";
            // 
            // lbl_nif
            // 
            this.lbl_nif.AutoSize = true;
            this.lbl_nif.Location = new System.Drawing.Point(77, 25);
            this.lbl_nif.Name = "lbl_nif";
            this.lbl_nif.Size = new System.Drawing.Size(27, 13);
            this.lbl_nif.TabIndex = 66;
            this.lbl_nif.Text = "NIF:";
            // 
            // lbl_distribuidora
            // 
            this.lbl_distribuidora.AutoSize = true;
            this.lbl_distribuidora.Location = new System.Drawing.Point(36, 77);
            this.lbl_distribuidora.Name = "lbl_distribuidora";
            this.lbl_distribuidora.Size = new System.Drawing.Size(68, 13);
            this.lbl_distribuidora.TabIndex = 67;
            this.lbl_distribuidora.Text = "Distribuidora:";
            // 
            // txt_nif
            // 
            this.txt_nif.Enabled = false;
            this.txt_nif.Location = new System.Drawing.Point(110, 22);
            this.txt_nif.Name = "txt_nif";
            this.txt_nif.Size = new System.Drawing.Size(286, 20);
            this.txt_nif.TabIndex = 68;
            // 
            // txt_cliente
            // 
            this.txt_cliente.Enabled = false;
            this.txt_cliente.Location = new System.Drawing.Point(110, 48);
            this.txt_cliente.Name = "txt_cliente";
            this.txt_cliente.Size = new System.Drawing.Size(286, 20);
            this.txt_cliente.TabIndex = 69;
            // 
            // txt_distribuidora
            // 
            this.txt_distribuidora.Enabled = false;
            this.txt_distribuidora.Location = new System.Drawing.Point(110, 74);
            this.txt_distribuidora.Name = "txt_distribuidora";
            this.txt_distribuidora.Size = new System.Drawing.Size(286, 20);
            this.txt_distribuidora.TabIndex = 70;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.txt_qi);
            this.groupBox1.Controls.Add(this.txt_horaInicio);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.txtCodIndProd);
            this.groupBox1.Controls.Add(this.cmbTipoProducto);
            this.groupBox1.Controls.Add(this.chkSoloXML);
            this.groupBox1.Controls.Add(this.txt_fecha_fin);
            this.groupBox1.Controls.Add(this.txt_fecha_inicio);
            this.groupBox1.Controls.Add(this.chkSolapado);
            this.groupBox1.Controls.Add(this.txt_tarifa);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txt_comentario);
            this.groupBox1.Controls.Add(this.cmb_tipo);
            this.groupBox1.Controls.Add(this.txt_qd);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txt_cups20);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Location = new System.Drawing.Point(12, 140);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(447, 492);
            this.groupBox1.TabIndex = 71;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Datos Solicitud";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(234, 281);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(44, 13);
            this.label15.TabIndex = 96;
            this.label15.Text = "HH:MM";
            // 
            // txt_qi
            // 
            this.txt_qi.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.RecentlyUsedList;
            this.txt_qi.Enabled = false;
            this.txt_qi.Location = new System.Drawing.Point(140, 252);
            this.txt_qi.Name = "txt_qi";
            this.txt_qi.Size = new System.Drawing.Size(88, 20);
            this.txt_qi.TabIndex = 95;
            this.txt_qi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.toolTip_Qi.SetToolTip(this.txt_qi, "Caudal intradiario del producto. Indica la capacidad contratada para el total de " +
        "horas del producto intradiario, expresada en términos de kWh.");
            this.txt_qi.TextChanged += new System.EventHandler(this.txt_qi_TextChanged);
            this.txt_qi.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_qi_KeyPress);
            // 
            // txt_horaInicio
            // 
            this.txt_horaInicio.Enabled = false;
            this.txt_horaInicio.Location = new System.Drawing.Point(140, 278);
            this.txt_horaInicio.Mask = "00:00";
            this.txt_horaInicio.Name = "txt_horaInicio";
            this.txt_horaInicio.Size = new System.Drawing.Size(88, 20);
            this.txt_horaInicio.TabIndex = 94;
            this.txt_horaInicio.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_horaInicio.ValidatingType = typeof(System.DateTime);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(59, 280);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(75, 13);
            this.label14.TabIndex = 93;
            this.label14.Text = "Hora de inicio:";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(234, 256);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(82, 13);
            this.label12.TabIndex = 92;
            this.label12.Text = "kWh/día/horas";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(114, 254);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(20, 13);
            this.label13.TabIndex = 91;
            this.label13.Text = "Qi:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(25, 119);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(109, 13);
            this.label11.TabIndex = 89;
            this.label11.Text = "Código ind. Producto:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(14, 92);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(120, 13);
            this.label10.TabIndex = 88;
            this.label10.Text = "Tipo Solicitud Producto:";
            // 
            // txtCodIndProd
            // 
            this.txtCodIndProd.Enabled = false;
            this.txtCodIndProd.Location = new System.Drawing.Point(140, 116);
            this.txtCodIndProd.Name = "txtCodIndProd";
            this.txtCodIndProd.Size = new System.Drawing.Size(166, 20);
            this.txtCodIndProd.TabIndex = 87;
            this.txtCodIndProd.TextChanged += new System.EventHandler(this.txtCodIndProd_TextChanged);
            // 
            // cmbTipoProducto
            // 
            this.cmbTipoProducto.AutoCompleteCustomSource.AddRange(new string[] {
            "ANUAL",
            "DIARIO",
            "INDEFINIDO",
            "INTRADIARIO",
            "MENSUAL",
            "TRIMESTRAL"});
            this.cmbTipoProducto.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTipoProducto.Items.AddRange(new object[] {
            "01 - Nuevo producto consecutivo al existente manteniendo condiciones",
            "02 - Nuevo producto consecutivo al existente modificando condiciones",
            "03 - Nuevo producto "});
            this.cmbTipoProducto.Location = new System.Drawing.Point(140, 89);
            this.cmbTipoProducto.Name = "cmbTipoProducto";
            this.cmbTipoProducto.Size = new System.Drawing.Size(285, 21);
            this.cmbTipoProducto.TabIndex = 86;
            this.cmbTipoProducto.SelectedIndexChanged += new System.EventHandler(this.cmbTipoProducto_SelectedIndexChanged);
            // 
            // chkSoloXML
            // 
            this.chkSoloXML.AutoSize = true;
            this.chkSoloXML.Checked = true;
            this.chkSoloXML.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSoloXML.Location = new System.Drawing.Point(225, 19);
            this.chkSoloXML.Name = "chkSoloXML";
            this.chkSoloXML.Size = new System.Drawing.Size(209, 17);
            this.chkSoloXML.TabIndex = 85;
            this.chkSoloXML.Text = "No enviar al SCTD (Sólo generar XML)";
            this.chkSoloXML.UseVisualStyleBackColor = true;
            // 
            // txt_fecha_fin
            // 
            this.txt_fecha_fin.Enabled = false;
            this.txt_fecha_fin.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_fin.Location = new System.Drawing.Point(140, 199);
            this.txt_fecha_fin.Name = "txt_fecha_fin";
            this.txt_fecha_fin.Size = new System.Drawing.Size(99, 20);
            this.txt_fecha_fin.TabIndex = 4;
            // 
            // txt_fecha_inicio
            // 
            this.txt_fecha_inicio.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_inicio.Location = new System.Drawing.Point(140, 173);
            this.txt_fecha_inicio.Name = "txt_fecha_inicio";
            this.txt_fecha_inicio.Size = new System.Drawing.Size(99, 20);
            this.txt_fecha_inicio.TabIndex = 3;
            this.txt_fecha_inicio.ValueChanged += new System.EventHandler(this.txt_fecha_inicio_ValueChanged);
            // 
            // chkSolapado
            // 
            this.chkSolapado.AutoSize = true;
            this.chkSolapado.Location = new System.Drawing.Point(17, 19);
            this.chkSolapado.Name = "chkSolapado";
            this.chkSolapado.Size = new System.Drawing.Size(71, 17);
            this.chkSolapado.TabIndex = 82;
            this.chkSolapado.Text = "Solapado";
            this.chkSolapado.UseVisualStyleBackColor = true;
            // 
            // txt_tarifa
            // 
            this.txt_tarifa.Enabled = false;
            this.txt_tarifa.Location = new System.Drawing.Point(140, 304);
            this.txt_tarifa.Name = "txt_tarifa";
            this.txt_tarifa.Size = new System.Drawing.Size(88, 20);
            this.txt_tarifa.TabIndex = 81;
            this.txt_tarifa.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_tarifa.TextChanged += new System.EventHandler(this.txt_tarifa_TextChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(71, 333);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 13);
            this.label8.TabIndex = 80;
            this.label8.Text = "Comentario:";
            // 
            // txt_comentario
            // 
            this.txt_comentario.Location = new System.Drawing.Point(140, 330);
            this.txt_comentario.Multiline = true;
            this.txt_comentario.Name = "txt_comentario";
            this.txt_comentario.Size = new System.Drawing.Size(285, 127);
            this.txt_comentario.TabIndex = 6;
            // 
            // cmb_tipo
            // 
            this.cmb_tipo.AutoCompleteCustomSource.AddRange(new string[] {
            "ANUAL",
            "DIARIO",
            "INDEFINIDO",
            "INTRADIARIO",
            "MENSUAL",
            "TRIMESTRAL"});
            this.cmb_tipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_tipo.Items.AddRange(new object[] {
            "ANUAL",
            "DIARIO",
            "INDEFINIDO",
            "INTRADIARIO",
            "MENSUAL",
            "TRIMESTRAL"});
            this.cmb_tipo.Location = new System.Drawing.Point(140, 145);
            this.cmb_tipo.Name = "cmb_tipo";
            this.cmb_tipo.Size = new System.Drawing.Size(121, 21);
            this.cmb_tipo.TabIndex = 2;
            this.cmb_tipo.SelectedIndexChanged += new System.EventHandler(this.cmb_tipo_SelectedIndexChanged);
            // 
            // txt_qd
            // 
            this.txt_qd.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.RecentlyUsedList;
            this.txt_qd.Location = new System.Drawing.Point(140, 225);
            this.txt_qd.Name = "txt_qd";
            this.txt_qd.Size = new System.Drawing.Size(88, 20);
            this.txt_qd.TabIndex = 5;
            this.txt_qd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txt_qd.TextChanged += new System.EventHandler(this.txt_qd_TextChanged);
            this.txt_qd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_qd_KeyPress);
            this.txt_qd.Leave += new System.EventHandler(this.txt_qd_Leave);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(234, 228);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 78;
            this.label4.Text = "kWh/día";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(83, 148);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 13);
            this.label7.TabIndex = 77;
            this.label7.Text = "Producto:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(97, 307);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 13);
            this.label6.TabIndex = 76;
            this.label6.Text = "Tarifa:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(110, 228);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(24, 13);
            this.label5.TabIndex = 75;
            this.label5.Text = "Qd:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(80, 202);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 74;
            this.label3.Text = "Fecha fin:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(67, 176);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 73;
            this.label2.Text = "Fecha inicio:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(80, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 72;
            this.label1.Text = "CUPS20:";
            // 
            // txt_cups20
            // 
            this.txt_cups20.Location = new System.Drawing.Point(140, 63);
            this.txt_cups20.Name = "txt_cups20";
            this.txt_cups20.Size = new System.Drawing.Size(166, 20);
            this.txt_cups20.TabIndex = 1;
            this.txt_cups20.TextChanged += new System.EventHandler(this.txt_cups20_TextChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(218, 463);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 66;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(302, 463);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(123, 23);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "Enviar solicitud";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txt_tipotratamitacion
            // 
            this.txt_tipotratamitacion.Enabled = false;
            this.txt_tipotratamitacion.Location = new System.Drawing.Point(110, 101);
            this.txt_tipotratamitacion.Name = "txt_tipotratamitacion";
            this.txt_tipotratamitacion.Size = new System.Drawing.Size(286, 20);
            this.txt_tipotratamitacion.TabIndex = 73;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 104);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(89, 13);
            this.label9.TabIndex = 72;
            this.label9.Text = "Tipo Tramitación:";
            this.label9.Click += new System.EventHandler(this.label9_Click);
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // toolTip_Qi
            // 
            this.toolTip_Qi.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTip1_Popup);
            // 
            // toolTip1
            // 
            this.toolTip1.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTip1_Popup_1);
            // 
            // FrmSolicitudManual
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 644);
            this.Controls.Add(this.txt_tipotratamitacion);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txt_distribuidora);
            this.Controls.Add(this.txt_cliente);
            this.Controls.Add(this.txt_nif);
            this.Controls.Add(this.lbl_distribuidora);
            this.Controls.Add(this.lbl_nif);
            this.Controls.Add(this.lbl_cliente);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmSolicitudManual";
            this.Text = "Solicitud Manual";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmSolicitudManual_FormClosing);
            this.Load += new System.EventHandler(this.FrmSolicitudManual_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lbl_cliente;
        private System.Windows.Forms.Label lbl_nif;
        private System.Windows.Forms.Label lbl_distribuidora;
        private System.Windows.Forms.TextBox txt_nif;
        private System.Windows.Forms.TextBox txt_cliente;
        private System.Windows.Forms.TextBox txt_distribuidora;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkSolapado;
        public System.Windows.Forms.TextBox txt_tarifa;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.TextBox txt_comentario;
        public System.Windows.Forms.ComboBox cmb_tipo;
        public System.Windows.Forms.TextBox txt_qd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_cups20;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txt_tipotratamitacion;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.DateTimePicker txt_fecha_fin;
        private System.Windows.Forms.DateTimePicker txt_fecha_inicio;
        private System.Windows.Forms.CheckBox chkSoloXML;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        public System.Windows.Forms.TextBox txtCodIndProd;
        public System.Windows.Forms.ComboBox cmbTipoProducto;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ToolTip toolTip_Qi;
        public System.Windows.Forms.TextBox txt_qi;
        private System.Windows.Forms.MaskedTextBox txt_horaInicio;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label15;
    }
}