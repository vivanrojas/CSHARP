namespace GestionOperaciones.forms.medida
{
    partial class FrmCurvaResumen
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCurvaResumen));
            this.dgv = new System.Windows.Forms.DataGridView();
            this.CPUNTMED = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FFACTDES = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FFACTHAS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VSECRCUR = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TESTRCUR = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VEACONTO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VERCONTO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FUENTE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToDeleteRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CPUNTMED,
            this.FFACTDES,
            this.FFACTHAS,
            this.VSECRCUR,
            this.TESTRCUR,
            this.VEACONTO,
            this.VERCONTO,
            this.FUENTE});
            this.dgv.Location = new System.Drawing.Point(12, 12);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.Size = new System.Drawing.Size(710, 135);
            this.dgv.TabIndex = 0;
            this.dgv.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellContentClick);
            // 
            // CPUNTMED
            // 
            this.CPUNTMED.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.CPUNTMED.DataPropertyName = "CPUNTMED";
            this.CPUNTMED.HeaderText = "CPUNTMED";
            this.CPUNTMED.Name = "CPUNTMED";
            this.CPUNTMED.ReadOnly = true;
            this.CPUNTMED.Width = 84;
            // 
            // FFACTDES
            // 
            this.FFACTDES.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FFACTDES.DataPropertyName = "FFACTDES";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FFACTDES.DefaultCellStyle = dataGridViewCellStyle2;
            this.FFACTDES.HeaderText = "FFACTDES";
            this.FFACTDES.Name = "FFACTDES";
            this.FFACTDES.ReadOnly = true;
            this.FFACTDES.Width = 80;
            // 
            // FFACTHAS
            // 
            this.FFACTHAS.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FFACTHAS.DataPropertyName = "FFACTHAS";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.FFACTHAS.DefaultCellStyle = dataGridViewCellStyle3;
            this.FFACTHAS.HeaderText = "FFACTHAS";
            this.FFACTHAS.Name = "FFACTHAS";
            this.FFACTHAS.ReadOnly = true;
            this.FFACTHAS.Width = 81;
            // 
            // VSECRCUR
            // 
            this.VSECRCUR.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.VSECRCUR.DataPropertyName = "VSECRCUR";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.VSECRCUR.DefaultCellStyle = dataGridViewCellStyle4;
            this.VSECRCUR.HeaderText = "VERSIÓN";
            this.VSECRCUR.Name = "VSECRCUR";
            this.VSECRCUR.ReadOnly = true;
            this.VSECRCUR.Width = 73;
            // 
            // TESTRCUR
            // 
            this.TESTRCUR.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.TESTRCUR.DataPropertyName = "TESTRCUR";
            this.TESTRCUR.HeaderText = "TESTRCUR";
            this.TESTRCUR.Name = "TESTRCUR";
            this.TESTRCUR.ReadOnly = true;
            this.TESTRCUR.Width = 80;
            // 
            // VEACONTO
            // 
            this.VEACONTO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.VEACONTO.DataPropertyName = "VEACONTO";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N0";
            dataGridViewCellStyle5.NullValue = null;
            this.VEACONTO.DefaultCellStyle = dataGridViewCellStyle5;
            this.VEACONTO.HeaderText = "ACTIVA";
            this.VEACONTO.Name = "VEACONTO";
            this.VEACONTO.ReadOnly = true;
            this.VEACONTO.Width = 66;
            // 
            // VERCONTO
            // 
            this.VERCONTO.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.VERCONTO.DataPropertyName = "VERCONTO";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            this.VERCONTO.DefaultCellStyle = dataGridViewCellStyle6;
            this.VERCONTO.HeaderText = "REACTIVA";
            this.VERCONTO.Name = "VERCONTO";
            this.VERCONTO.ReadOnly = true;
            this.VERCONTO.Width = 79;
            // 
            // FUENTE
            // 
            this.FUENTE.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.FUENTE.DataPropertyName = "FUENTE";
            this.FUENTE.HeaderText = "FUENTE";
            this.FUENTE.Name = "FUENTE";
            this.FUENTE.ReadOnly = true;
            this.FUENTE.Width = 67;
            // 
            // FrmCurvaResumen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(729, 159);
            this.Controls.Add(this.dgv);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmCurvaResumen";
            this.Text = "Curva Resumen";
            this.Load += new System.EventHandler(this.FrmCurvaResumen_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn CPUNTMED;
        private System.Windows.Forms.DataGridViewTextBoxColumn FFACTDES;
        private System.Windows.Forms.DataGridViewTextBoxColumn FFACTHAS;
        private System.Windows.Forms.DataGridViewTextBoxColumn VSECRCUR;
        private System.Windows.Forms.DataGridViewTextBoxColumn TESTRCUR;
        private System.Windows.Forms.DataGridViewTextBoxColumn VEACONTO;
        private System.Windows.Forms.DataGridViewTextBoxColumn VERCONTO;
        private System.Windows.Forms.DataGridViewTextBoxColumn FUENTE;
    }
}