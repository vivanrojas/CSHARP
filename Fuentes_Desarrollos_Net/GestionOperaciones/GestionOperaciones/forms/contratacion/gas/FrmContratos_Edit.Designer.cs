namespace GestionOperaciones.forms.contratacion.gas
{
    partial class FrmContratos_Edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmContratos_Edit));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cmb_tramitacion = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cmb_distribuidora = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_comentarios_contratacion = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_comentarios_descuadres = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txt_cups20 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_dapersoc = new System.Windows.Forms.TextBox();
            this.txt_cnifdnic = new System.Windows.Forms.TextBox();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(384, 638);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 35);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(506, 638);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(112, 35);
            this.btnOK.TabIndex = 10;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cmb_tramitacion);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.cmb_distribuidora);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_comentarios_contratacion);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txt_comentarios_descuadres);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txt_cups20);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txt_dapersoc);
            this.groupBox1.Controls.Add(this.txt_cnifdnic);
            this.groupBox1.Location = new System.Drawing.Point(18, 18);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(602, 611);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Datos contrato";
            // 
            // cmb_tramitacion
            // 
            this.cmb_tramitacion.AutoCompleteCustomSource.AddRange(new string[] {
            "ANUAL",
            "DIARIO",
            "INDEFINIDO",
            "INTRADIARIO",
            "MENSUAL",
            "TRIMESTRAL"});
            this.cmb_tramitacion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_tramitacion.Items.AddRange(new object[] {
            "Distribuidora",
            "Mail"});
            this.cmb_tramitacion.Location = new System.Drawing.Point(135, 203);
            this.cmb_tramitacion.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmb_tramitacion.Name = "cmb_tramitacion";
            this.cmb_tramitacion.Size = new System.Drawing.Size(217, 28);
            this.cmb_tramitacion.TabIndex = 78;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(28, 208);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(95, 20);
            this.label5.TabIndex = 79;
            this.label5.Text = "Tramitacion:";
            // 
            // cmb_distribuidora
            // 
            this.cmb_distribuidora.FormattingEnabled = true;
            this.cmb_distribuidora.Location = new System.Drawing.Point(135, 122);
            this.cmb_distribuidora.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmb_distribuidora.Name = "cmb_distribuidora";
            this.cmb_distribuidora.Size = new System.Drawing.Size(415, 28);
            this.cmb_distribuidora.TabIndex = 28;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 428);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(195, 20);
            this.label3.TabIndex = 27;
            this.label3.Text = "Comentarios contratación:";
            // 
            // txt_comentarios_contratacion
            // 
            this.txt_comentarios_contratacion.Location = new System.Drawing.Point(14, 452);
            this.txt_comentarios_contratacion.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_comentarios_contratacion.Multiline = true;
            this.txt_comentarios_contratacion.Name = "txt_comentarios_contratacion";
            this.txt_comentarios_contratacion.Size = new System.Drawing.Size(577, 113);
            this.txt_comentarios_contratacion.TabIndex = 26;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 272);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(194, 20);
            this.label2.TabIndex = 25;
            this.label2.Text = " Comentarios descuadres:";
            // 
            // txt_comentarios_descuadres
            // 
            this.txt_comentarios_descuadres.Location = new System.Drawing.Point(14, 297);
            this.txt_comentarios_descuadres.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_comentarios_descuadres.Multiline = true;
            this.txt_comentarios_descuadres.Name = "txt_comentarios_descuadres";
            this.txt_comentarios_descuadres.Size = new System.Drawing.Size(577, 113);
            this.txt_comentarios_descuadres.TabIndex = 24;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(50, 168);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(75, 20);
            this.label7.TabIndex = 23;
            this.label7.Text = "CUPS20:";
            // 
            // txt_cups20
            // 
            this.txt_cups20.Location = new System.Drawing.Point(135, 163);
            this.txt_cups20.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_cups20.Name = "txt_cups20";
            this.txt_cups20.Size = new System.Drawing.Size(217, 26);
            this.txt_cups20.TabIndex = 22;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(24, 126);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(102, 20);
            this.label6.TabIndex = 21;
            this.label6.Text = "Distribuidora:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(58, 86);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(62, 20);
            this.label4.TabIndex = 18;
            this.label4.Text = "Cliente:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(86, 46);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 20);
            this.label1.TabIndex = 15;
            this.label1.Text = "NIF:";
            // 
            // txt_dapersoc
            // 
            this.txt_dapersoc.Location = new System.Drawing.Point(135, 82);
            this.txt_dapersoc.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_dapersoc.Name = "txt_dapersoc";
            this.txt_dapersoc.Size = new System.Drawing.Size(415, 26);
            this.txt_dapersoc.TabIndex = 11;
            // 
            // txt_cnifdnic
            // 
            this.txt_cnifdnic.Location = new System.Drawing.Point(135, 42);
            this.txt_cnifdnic.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_cnifdnic.Name = "txt_cnifdnic";
            this.txt_cnifdnic.Size = new System.Drawing.Size(180, 26);
            this.txt_cnifdnic.TabIndex = 10;
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // FrmContratos_Edit
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(638, 700);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmContratos_Edit";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "FrmDistribuidoras_Edit";
            this.Load += new System.EventHandler(this.FrmContratos_Edit_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox txt_cnifdnic;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.Label label7;
        public System.Windows.Forms.TextBox txt_cups20;
        public System.Windows.Forms.ComboBox cmb_distribuidora;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox txt_comentarios_contratacion;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox txt_comentarios_descuadres;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox txt_dapersoc;
        public System.Windows.Forms.ComboBox cmb_tramitacion;
        private System.Windows.Forms.Label label5;
    }
}