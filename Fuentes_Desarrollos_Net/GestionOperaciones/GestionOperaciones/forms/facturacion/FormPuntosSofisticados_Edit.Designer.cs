namespace GestionOperaciones.forms.facturacion
{
    partial class FormPuntosSofisticados_Edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormPuntosSofisticados_Edit));
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtCCOUNIPS = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDAPERSOC = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtGRUPO = new System.Windows.Forms.TextBox();
            this.chkFACTURAS_A_CUENTA = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtPrecios = new System.Windows.Forms.TextBox();
            this.txtFH = new System.Windows.Forms.MaskedTextBox();
            this.txtFD = new System.Windows.Forms.MaskedTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtcusp20 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(195, 240);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 9;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(276, 240);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtCCOUNIPS
            // 
            this.txtCCOUNIPS.Location = new System.Drawing.Point(94, 12);
            this.txtCCOUNIPS.Name = "txtCCOUNIPS";
            this.txtCCOUNIPS.Size = new System.Drawing.Size(100, 20);
            this.txtCCOUNIPS.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(40, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "CUPS13";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "DAPERSOC";
            // 
            // txtDAPERSOC
            // 
            this.txtDAPERSOC.Location = new System.Drawing.Point(94, 66);
            this.txtDAPERSOC.Name = "txtDAPERSOC";
            this.txtDAPERSOC.Size = new System.Drawing.Size(203, 20);
            this.txtDAPERSOC.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(42, 95);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "GRUPO";
            // 
            // txtGRUPO
            // 
            this.txtGRUPO.Location = new System.Drawing.Point(94, 92);
            this.txtGRUPO.Name = "txtGRUPO";
            this.txtGRUPO.Size = new System.Drawing.Size(203, 20);
            this.txtGRUPO.TabIndex = 4;
            // 
            // chkFACTURAS_A_CUENTA
            // 
            this.chkFACTURAS_A_CUENTA.AutoSize = true;
            this.chkFACTURAS_A_CUENTA.Location = new System.Drawing.Point(94, 196);
            this.chkFACTURAS_A_CUENTA.Name = "chkFACTURAS_A_CUENTA";
            this.chkFACTURAS_A_CUENTA.Size = new System.Drawing.Size(112, 17);
            this.chkFACTURAS_A_CUENTA.TabIndex = 8;
            this.chkFACTURAS_A_CUENTA.Text = "Facturas a cuenta";
            this.chkFACTURAS_A_CUENTA.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 124);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Fecha Desde";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 150);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Fecha Hasta";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(34, 173);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(54, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "PRECIOS";
            // 
            // txtPrecios
            // 
            this.txtPrecios.Location = new System.Drawing.Point(94, 170);
            this.txtPrecios.Name = "txtPrecios";
            this.txtPrecios.Size = new System.Drawing.Size(112, 20);
            this.txtPrecios.TabIndex = 7;
            // 
            // txtFH
            // 
            this.txtFH.Location = new System.Drawing.Point(94, 143);
            this.txtFH.Mask = "00/00/0000";
            this.txtFH.Name = "txtFH";
            this.txtFH.Size = new System.Drawing.Size(100, 20);
            this.txtFH.TabIndex = 6;
            this.txtFH.ValidatingType = typeof(System.DateTime);
            // 
            // txtFD
            // 
            this.txtFD.AsciiOnly = true;
            this.txtFD.Location = new System.Drawing.Point(94, 117);
            this.txtFD.Mask = "00/00/0000";
            this.txtFD.Name = "txtFD";
            this.txtFD.Size = new System.Drawing.Size(100, 20);
            this.txtFD.TabIndex = 5;
            this.txtFD.ValidatingType = typeof(System.DateTime);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(40, 41);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(48, 13);
            this.label7.TabIndex = 17;
            this.label7.Text = "CUPS20";
            // 
            // txtcusp20
            // 
            this.txtcusp20.Location = new System.Drawing.Point(94, 38);
            this.txtcusp20.Name = "txtcusp20";
            this.txtcusp20.Size = new System.Drawing.Size(203, 20);
            this.txtcusp20.TabIndex = 2;
            // 
            // FormPuntosSofisticados_Edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 275);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtcusp20);
            this.Controls.Add(this.txtFD);
            this.Controls.Add(this.txtFH);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtPrecios);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.chkFACTURAS_A_CUENTA);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtGRUPO);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtDAPERSOC);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtCCOUNIPS);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormPuntosSofisticados_Edit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sofisticados";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox txtCCOUNIPS;
        public System.Windows.Forms.TextBox txtDAPERSOC;
        public System.Windows.Forms.TextBox txtGRUPO;
        public System.Windows.Forms.CheckBox chkFACTURAS_A_CUENTA;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        public System.Windows.Forms.TextBox txtPrecios;
        public System.Windows.Forms.MaskedTextBox txtFH;
        public System.Windows.Forms.MaskedTextBox txtFD;
        private System.Windows.Forms.Label label7;
        public System.Windows.Forms.TextBox txtcusp20;
    }
}