namespace GestionOperaciones.forms.facturacion
{
    partial class FrmTunelCuadroMando_Edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmTunelCuadroMando_Edit));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chk_formula_antigua = new System.Windows.Forms.CheckBox();
            this.txt_fecha_fin_tunel = new System.Windows.Forms.DateTimePicker();
            this.txt_fecha_inicio_tunel = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txt_cliente = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chk_formula_antigua);
            this.groupBox1.Controls.Add(this.txt_fecha_fin_tunel);
            this.groupBox1.Controls.Add(this.txt_fecha_inicio_tunel);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.txt_cliente);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Location = new System.Drawing.Point(16, 15);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Size = new System.Drawing.Size(716, 428);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Datos inventario";
            // 
            // chk_formula_antigua
            // 
            this.chk_formula_antigua.AutoSize = true;
            this.chk_formula_antigua.Location = new System.Drawing.Point(51, 97);
            this.chk_formula_antigua.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_formula_antigua.Name = "chk_formula_antigua";
            this.chk_formula_antigua.Size = new System.Drawing.Size(119, 19);
            this.chk_formula_antigua.TabIndex = 49;
            this.chk_formula_antigua.Text = "Fórmula antigua";
            this.chk_formula_antigua.UseVisualStyleBackColor = true;
            // 
            // txt_fecha_fin_tunel
            // 
            this.txt_fecha_fin_tunel.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_fin_tunel.Location = new System.Drawing.Point(429, 55);
            this.txt_fecha_fin_tunel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txt_fecha_fin_tunel.Name = "txt_fecha_fin_tunel";
            this.txt_fecha_fin_tunel.Size = new System.Drawing.Size(116, 20);
            this.txt_fecha_fin_tunel.TabIndex = 48;
            // 
            // txt_fecha_inicio_tunel
            // 
            this.txt_fecha_inicio_tunel.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txt_fecha_inicio_tunel.Location = new System.Drawing.Point(165, 55);
            this.txt_fecha_inicio_tunel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txt_fecha_inicio_tunel.Name = "txt_fecha_inicio_tunel";
            this.txt_fecha_inicio_tunel.Size = new System.Drawing.Size(116, 20);
            this.txt_fecha_inicio_tunel.TabIndex = 47;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(304, 57);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(100, 15);
            this.label9.TabIndex = 46;
            this.label9.Text = "Fecha final tunel:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(37, 57);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(106, 15);
            this.label10.TabIndex = 45;
            this.label10.Text = "Fecha inicio tunel:";
            // 
            // txt_cliente
            // 
            this.txt_cliente.Location = new System.Drawing.Point(165, 23);
            this.txt_cliente.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txt_cliente.Name = "txt_cliente";
            this.txt_cliente.Size = new System.Drawing.Size(480, 20);
            this.txt_cliente.TabIndex = 44;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(101, 27);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 15);
            this.label6.TabIndex = 43;
            this.label6.Text = "Cliente:";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(524, 450);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 28);
            this.btnCancel.TabIndex = 18;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(632, 450);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(100, 28);
            this.btnOK.TabIndex = 17;
            this.btnOK.Text = "Aceptar";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // FrmTunelCuadroMando_Edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 496);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "FrmTunelCuadroMando_Edit";
            this.Text = "Datos contrato";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.DateTimePicker txt_fecha_fin_tunel;
        public System.Windows.Forms.DateTimePicker txt_fecha_inicio_tunel;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        public System.Windows.Forms.TextBox txt_cliente;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        public System.Windows.Forms.CheckBox chk_formula_antigua;
    }
}