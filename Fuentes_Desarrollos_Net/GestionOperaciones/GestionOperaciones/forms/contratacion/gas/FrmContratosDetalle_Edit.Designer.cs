namespace GestionOperaciones.forms.contratacion.gas
{
    partial class FrmContratosDetalle_Edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmContratosDetalle_Edit));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.cmb_tarifa = new System.Windows.Forms.ComboBox();
            this.cmb_tipo = new System.Windows.Forms.ComboBox();
            this.txt_qd = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_fecha_fin = new System.Windows.Forms.MaskedTextBox();
            this.txt_fecha_inicio = new System.Windows.Forms.MaskedTextBox();
            this.txt_cups20 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_comentario = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(198, 345);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(279, 345);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // cmb_tarifa
            // 
            this.cmb_tarifa.FormattingEnabled = true;
            this.cmb_tarifa.Items.AddRange(new object[] {
            "RL.1",
            "RL.10",
            "RL.11",
            "RL.2",
            "RL.3",
            "RL.4",
            "RL.8",
            "RL.9",
            "RLPS1",
            "RLPS2",
            "RLPS3",
            "RLPS4",
            "RLPS5",
            "RLPS6",
            "RLPS7",
            "RLPS8",
            "RLTA5",
            "RLTA6",
            "RLTA7",
            "RLTB5",
            "RLTB6",
            "RLTB7"});
            this.cmb_tarifa.Location = new System.Drawing.Point(101, 185);
            this.cmb_tarifa.Name = "cmb_tarifa";
            this.cmb_tarifa.Size = new System.Drawing.Size(121, 21);
            this.cmb_tarifa.TabIndex = 35;
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
            this.cmb_tipo.Location = new System.Drawing.Point(101, 80);
            this.cmb_tipo.Name = "cmb_tipo";
            this.cmb_tipo.Size = new System.Drawing.Size(121, 21);
            this.cmb_tipo.TabIndex = 36;
            this.cmb_tipo.SelectedIndexChanged += new System.EventHandler(this.cmb_tipo_SelectedIndexChanged);
            // 
            // txt_qd
            // 
            this.txt_qd.Location = new System.Drawing.Point(101, 160);
            this.txt_qd.Name = "txt_qd";
            this.txt_qd.Size = new System.Drawing.Size(88, 20);
            this.txt_qd.TabIndex = 34;
            this.txt_qd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txt_qd.TextChanged += new System.EventHandler(this.txt_qd_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(195, 163);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 43;
            this.label4.Text = "kWh/día";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(65, 83);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(31, 13);
            this.label7.TabIndex = 42;
            this.label7.Text = "Tipo:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(58, 189);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 13);
            this.label6.TabIndex = 41;
            this.label6.Text = "Tarifa:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(71, 163);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(24, 13);
            this.label5.TabIndex = 40;
            this.label5.Text = "Qd:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(41, 137);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 39;
            this.label3.Text = "Fecha fin:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 38;
            this.label2.Text = "Fecha inicio:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(44, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 37;
            this.label1.Text = "CUPS20:";
            // 
            // txt_fecha_fin
            // 
            this.txt_fecha_fin.Location = new System.Drawing.Point(101, 134);
            this.txt_fecha_fin.Mask = "00/00/0000";
            this.txt_fecha_fin.Name = "txt_fecha_fin";
            this.txt_fecha_fin.Size = new System.Drawing.Size(88, 20);
            this.txt_fecha_fin.TabIndex = 33;
            this.txt_fecha_fin.ValidatingType = typeof(System.DateTime);
            // 
            // txt_fecha_inicio
            // 
            this.txt_fecha_inicio.Location = new System.Drawing.Point(101, 108);
            this.txt_fecha_inicio.Mask = "00/00/0000";
            this.txt_fecha_inicio.Name = "txt_fecha_inicio";
            this.txt_fecha_inicio.Size = new System.Drawing.Size(88, 20);
            this.txt_fecha_inicio.TabIndex = 32;
            this.txt_fecha_inicio.ValidatingType = typeof(System.DateTime);
            // 
            // txt_cups20
            // 
            this.txt_cups20.Enabled = false;
            this.txt_cups20.Location = new System.Drawing.Point(101, 51);
            this.txt_cups20.Name = "txt_cups20";
            this.txt_cups20.Size = new System.Drawing.Size(166, 20);
            this.txt_cups20.TabIndex = 30;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(32, 215);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 13);
            this.label8.TabIndex = 45;
            this.label8.Text = "Comentario:";
            // 
            // txt_comentario
            // 
            this.txt_comentario.Location = new System.Drawing.Point(101, 212);
            this.txt_comentario.Multiline = true;
            this.txt_comentario.Name = "txt_comentario";
            this.txt_comentario.Size = new System.Drawing.Size(253, 127);
            this.txt_comentario.TabIndex = 44;
            // 
            // FrmContratosDetalle_Edit
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(372, 380);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txt_comentario);
            this.Controls.Add(this.cmb_tarifa);
            this.Controls.Add(this.cmb_tipo);
            this.Controls.Add(this.txt_qd);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txt_fecha_fin);
            this.Controls.Add(this.txt_fecha_inicio);
            this.Controls.Add(this.txt_cups20);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmContratosDetalle_Edit";
            this.Text = "FrmDistribuidoras_Edit";
            this.Load += new System.EventHandler(this.FrmContratosDetalle_Edit_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ErrorProvider errorProvider;
        public System.Windows.Forms.ComboBox cmb_tipo;
        public System.Windows.Forms.TextBox txt_qd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.MaskedTextBox txt_fecha_fin;
        public System.Windows.Forms.MaskedTextBox txt_fecha_inicio;
        public System.Windows.Forms.TextBox txt_cups20;
        public System.Windows.Forms.ComboBox cmb_tarifa;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.TextBox txt_comentario;
    }
}