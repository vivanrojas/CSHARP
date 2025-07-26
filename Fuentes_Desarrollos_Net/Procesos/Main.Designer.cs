
namespace Procesos
{
    partial class Main
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_reglas = new System.Windows.Forms.Button();
            this.btn_eliminar = new System.Windows.Forms.Button();
            this.btn_modificar = new System.Windows.Forms.Button();
            this.btn_agregar = new System.Windows.Forms.Button();
            this.treeProcesos = new System.Windows.Forms.TreeView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.herramientasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarColaDeProcesosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txt_parametros = new System.Windows.Forms.TextBox();
            this.txt_mensaje = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.btn_dependencias = new System.Windows.Forms.Button();
            this.label14 = new System.Windows.Forms.Label();
            this.txt_ultimo_ok = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txt_ultimo_lanzamiento = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btn_aceptar = new System.Windows.Forms.Button();
            this.btn_cancelar = new System.Windows.Forms.Button();
            this.txt_grupo = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chk_obligatorio = new System.Windows.Forms.CheckBox();
            this.chk_activo = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dgv_dependencias = new System.Windows.Forms.DataGridView();
            this.cmb_tipoParametro = new System.Windows.Forms.ComboBox();
            this.cmb_periodicidad = new System.Windows.Forms.ComboBox();
            this.txt_descripcion = new System.Windows.Forms.TextBox();
            this.cmb_tipoProceso = new System.Windows.Forms.ComboBox();
            this.txt_nombre = new System.Windows.Forms.TextBox();
            this.txt_bbdd = new System.Windows.Forms.TextBox();
            this.txt_ruta = new System.Windows.Forms.TextBox();
            this.txt_proceso = new System.Windows.Forms.TextBox();
            this.dgv_reglas = new System.Windows.Forms.DataGridView();
            this.btn_ejecuta_todo = new System.Windows.Forms.Button();
            this.btn_solo_actual = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.ejecutarSóloActualToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_dependencias)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_reglas)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_reglas);
            this.groupBox1.Controls.Add(this.btn_eliminar);
            this.groupBox1.Controls.Add(this.btn_modificar);
            this.groupBox1.Controls.Add(this.btn_agregar);
            this.groupBox1.Controls.Add(this.treeProcesos);
            this.groupBox1.Location = new System.Drawing.Point(12, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(366, 760);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btn_reglas
            // 
            this.btn_reglas.Location = new System.Drawing.Point(281, 731);
            this.btn_reglas.Name = "btn_reglas";
            this.btn_reglas.Size = new System.Drawing.Size(75, 23);
            this.btn_reglas.TabIndex = 4;
            this.btn_reglas.Text = "Reglas";
            this.btn_reglas.UseVisualStyleBackColor = true;
            // 
            // btn_eliminar
            // 
            this.btn_eliminar.Location = new System.Drawing.Point(188, 731);
            this.btn_eliminar.Name = "btn_eliminar";
            this.btn_eliminar.Size = new System.Drawing.Size(75, 23);
            this.btn_eliminar.TabIndex = 3;
            this.btn_eliminar.Text = "Eliminar";
            this.btn_eliminar.UseVisualStyleBackColor = true;
            this.btn_eliminar.Click += new System.EventHandler(this.btn_eliminar_Click);
            // 
            // btn_modificar
            // 
            this.btn_modificar.Location = new System.Drawing.Point(96, 731);
            this.btn_modificar.Name = "btn_modificar";
            this.btn_modificar.Size = new System.Drawing.Size(75, 23);
            this.btn_modificar.TabIndex = 2;
            this.btn_modificar.Text = "Modificar";
            this.btn_modificar.UseVisualStyleBackColor = true;
            this.btn_modificar.Click += new System.EventHandler(this.btn_modificar_Click);
            // 
            // btn_agregar
            // 
            this.btn_agregar.Location = new System.Drawing.Point(6, 731);
            this.btn_agregar.Name = "btn_agregar";
            this.btn_agregar.Size = new System.Drawing.Size(75, 23);
            this.btn_agregar.TabIndex = 1;
            this.btn_agregar.Text = "Agregar";
            this.btn_agregar.UseVisualStyleBackColor = true;
            this.btn_agregar.Click += new System.EventHandler(this.btn_agregar_Click);
            // 
            // treeProcesos
            // 
            this.treeProcesos.Location = new System.Drawing.Point(6, 17);
            this.treeProcesos.Name = "treeProcesos";
            this.treeProcesos.Size = new System.Drawing.Size(350, 708);
            this.treeProcesos.TabIndex = 0;
            this.treeProcesos.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeProcesos_AfterCheck);
            this.treeProcesos.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeProcesos_AfterSelect);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.editarToolStripMenuItem,
            this.herramientasToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(919, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.ejecutarSóloActualToolStripMenuItem,
            this.toolStripSeparator1,
            this.salirToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(207, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // editarToolStripMenuItem
            // 
            this.editarToolStripMenuItem.Name = "editarToolStripMenuItem";
            this.editarToolStripMenuItem.Size = new System.Drawing.Size(49, 20);
            this.editarToolStripMenuItem.Text = "Editar";
            // 
            // herramientasToolStripMenuItem
            // 
            this.herramientasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarColaDeProcesosToolStripMenuItem});
            this.herramientasToolStripMenuItem.Name = "herramientasToolStripMenuItem";
            this.herramientasToolStripMenuItem.Size = new System.Drawing.Size(90, 20);
            this.herramientasToolStripMenuItem.Text = "Herramientas";
            // 
            // importarColaDeProcesosToolStripMenuItem
            // 
            this.importarColaDeProcesosToolStripMenuItem.Name = "importarColaDeProcesosToolStripMenuItem";
            this.importarColaDeProcesosToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.importarColaDeProcesosToolStripMenuItem.Text = "Importar Cola de Procesos";
            this.importarColaDeProcesosToolStripMenuItem.Click += new System.EventHandler(this.importarColaDeProcesosToolStripMenuItem_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txt_parametros);
            this.groupBox2.Controls.Add(this.txt_mensaje);
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.btn_dependencias);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.txt_ultimo_ok);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.txt_ultimo_lanzamiento);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.btn_aceptar);
            this.groupBox2.Controls.Add(this.btn_cancelar);
            this.groupBox2.Controls.Add(this.txt_grupo);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.chk_obligatorio);
            this.groupBox2.Controls.Add(this.chk_activo);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.dgv_dependencias);
            this.groupBox2.Controls.Add(this.cmb_tipoParametro);
            this.groupBox2.Controls.Add(this.cmb_periodicidad);
            this.groupBox2.Controls.Add(this.txt_descripcion);
            this.groupBox2.Controls.Add(this.cmb_tipoProceso);
            this.groupBox2.Controls.Add(this.txt_nombre);
            this.groupBox2.Controls.Add(this.txt_bbdd);
            this.groupBox2.Controls.Add(this.txt_ruta);
            this.groupBox2.Controls.Add(this.txt_proceso);
            this.groupBox2.Controls.Add(this.dgv_reglas);
            this.groupBox2.Location = new System.Drawing.Point(384, 27);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(523, 760);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            // 
            // txt_parametros
            // 
            this.txt_parametros.Location = new System.Drawing.Point(173, 230);
            this.txt_parametros.Name = "txt_parametros";
            this.txt_parametros.Size = new System.Drawing.Size(340, 20);
            this.txt_parametros.TabIndex = 35;
            // 
            // txt_mensaje
            // 
            this.txt_mensaje.Enabled = false;
            this.txt_mensaje.Location = new System.Drawing.Point(9, 515);
            this.txt_mensaje.Name = "txt_mensaje";
            this.txt_mensaje.Size = new System.Drawing.Size(504, 20);
            this.txt_mensaje.TabIndex = 34;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(113, 22);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(57, 13);
            this.label15.TabIndex = 33;
            this.label15.Text = "Proceso:";
            // 
            // btn_dependencias
            // 
            this.btn_dependencias.Location = new System.Drawing.Point(397, 256);
            this.btn_dependencias.Name = "btn_dependencias";
            this.btn_dependencias.Size = new System.Drawing.Size(120, 23);
            this.btn_dependencias.TabIndex = 32;
            this.btn_dependencias.Text = "Dependencias";
            this.btn_dependencias.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(319, 443);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(126, 13);
            this.label14.TabIndex = 31;
            this.label14.Text = "Última ejecución OK:";
            // 
            // txt_ultimo_ok
            // 
            this.txt_ultimo_ok.Enabled = false;
            this.txt_ultimo_ok.Location = new System.Drawing.Point(322, 459);
            this.txt_ultimo_ok.Name = "txt_ultimo_ok";
            this.txt_ultimo_ok.Size = new System.Drawing.Size(199, 20);
            this.txt_ultimo_ok.TabIndex = 30;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(3, 443);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(117, 13);
            this.label13.TabIndex = 29;
            this.label13.Text = "Último lanzamiento:";
            // 
            // txt_ultimo_lanzamiento
            // 
            this.txt_ultimo_lanzamiento.Enabled = false;
            this.txt_ultimo_lanzamiento.Location = new System.Drawing.Point(6, 459);
            this.txt_ultimo_lanzamiento.Name = "txt_ultimo_lanzamiento";
            this.txt_ultimo_lanzamiento.Size = new System.Drawing.Size(199, 20);
            this.txt_ultimo_lanzamiento.TabIndex = 28;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(6, 499);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(215, 13);
            this.label12.TabIndex = 27;
            this.label12.Text = "Mensaje informativo de la ejecución:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(19, 266);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(92, 13);
            this.label11.TabIndex = 26;
            this.label11.Text = "Dependencias:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(99, 230);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(68, 13);
            this.label10.TabIndex = 25;
            this.label10.Text = "Parámetro:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(10, 552);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(50, 13);
            this.label9.TabIndex = 24;
            this.label9.Text = "Reglas:";
            // 
            // btn_aceptar
            // 
            this.btn_aceptar.Enabled = false;
            this.btn_aceptar.Location = new System.Drawing.Point(361, 731);
            this.btn_aceptar.Name = "btn_aceptar";
            this.btn_aceptar.Size = new System.Drawing.Size(75, 23);
            this.btn_aceptar.TabIndex = 5;
            this.btn_aceptar.Text = "Aceptar";
            this.btn_aceptar.UseVisualStyleBackColor = true;
            this.btn_aceptar.Click += new System.EventHandler(this.btn_aceptar_Click);
            // 
            // btn_cancelar
            // 
            this.btn_cancelar.Enabled = false;
            this.btn_cancelar.Location = new System.Drawing.Point(442, 731);
            this.btn_cancelar.Name = "btn_cancelar";
            this.btn_cancelar.Size = new System.Drawing.Size(75, 23);
            this.btn_cancelar.TabIndex = 5;
            this.btn_cancelar.Text = "Cancelar";
            this.btn_cancelar.UseVisualStyleBackColor = true;
            this.btn_cancelar.Click += new System.EventHandler(this.btn_cancelar_Click);
            // 
            // txt_grupo
            // 
            this.txt_grupo.Location = new System.Drawing.Point(62, 18);
            this.txt_grupo.Name = "txt_grupo";
            this.txt_grupo.Size = new System.Drawing.Size(32, 20);
            this.txt_grupo.TabIndex = 23;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(19, 154);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(148, 13);
            this.label8.TabIndex = 22;
            this.label8.Text = "Descripción del proceso:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(53, 206);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(114, 13);
            this.label7.TabIndex = 21;
            this.label7.Text = "Tipo de parámetro:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(10, 180);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(158, 13);
            this.label6.TabIndex = 20;
            this.label6.Text = "Periodicidad de ejecución:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(43, 128);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 13);
            this.label5.TabIndex = 19;
            this.label5.Text = "Nombre del proceso:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(65, 102);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 13);
            this.label4.TabIndex = 18;
            this.label4.Text = "Tipo de proceso:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(37, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "B. D.:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(11, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Ruta:";
            // 
            // chk_obligatorio
            // 
            this.chk_obligatorio.AutoSize = true;
            this.chk_obligatorio.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chk_obligatorio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chk_obligatorio.Location = new System.Drawing.Point(356, 21);
            this.chk_obligatorio.Name = "chk_obligatorio";
            this.chk_obligatorio.Size = new System.Drawing.Size(91, 17);
            this.chk_obligatorio.TabIndex = 15;
            this.chk_obligatorio.Text = "Obligatorio:";
            this.chk_obligatorio.UseVisualStyleBackColor = true;
            // 
            // chk_activo
            // 
            this.chk_activo.AutoSize = true;
            this.chk_activo.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chk_activo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chk_activo.Location = new System.Drawing.Point(243, 20);
            this.chk_activo.Name = "chk_activo";
            this.chk_activo.Size = new System.Drawing.Size(66, 17);
            this.chk_activo.TabIndex = 14;
            this.chk_activo.Text = "Activo:";
            this.chk_activo.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(11, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Grupo:";
            // 
            // dgv_dependencias
            // 
            this.dgv_dependencias.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_dependencias.Enabled = false;
            this.dgv_dependencias.Location = new System.Drawing.Point(9, 285);
            this.dgv_dependencias.Name = "dgv_dependencias";
            this.dgv_dependencias.Size = new System.Drawing.Size(508, 141);
            this.dgv_dependencias.TabIndex = 10;
            // 
            // cmb_tipoParametro
            // 
            this.cmb_tipoParametro.FormattingEnabled = true;
            this.cmb_tipoParametro.Location = new System.Drawing.Point(173, 203);
            this.cmb_tipoParametro.Name = "cmb_tipoParametro";
            this.cmb_tipoParametro.Size = new System.Drawing.Size(176, 21);
            this.cmb_tipoParametro.TabIndex = 9;
            // 
            // cmb_periodicidad
            // 
            this.cmb_periodicidad.FormattingEnabled = true;
            this.cmb_periodicidad.Location = new System.Drawing.Point(173, 176);
            this.cmb_periodicidad.Name = "cmb_periodicidad";
            this.cmb_periodicidad.Size = new System.Drawing.Size(176, 21);
            this.cmb_periodicidad.TabIndex = 8;
            // 
            // txt_descripcion
            // 
            this.txt_descripcion.Location = new System.Drawing.Point(173, 151);
            this.txt_descripcion.Name = "txt_descripcion";
            this.txt_descripcion.Size = new System.Drawing.Size(340, 20);
            this.txt_descripcion.TabIndex = 7;
            // 
            // cmb_tipoProceso
            // 
            this.cmb_tipoProceso.FormattingEnabled = true;
            this.cmb_tipoProceso.Location = new System.Drawing.Point(173, 99);
            this.cmb_tipoProceso.Name = "cmb_tipoProceso";
            this.cmb_tipoProceso.Size = new System.Drawing.Size(176, 21);
            this.cmb_tipoProceso.TabIndex = 6;
            // 
            // txt_nombre
            // 
            this.txt_nombre.Location = new System.Drawing.Point(173, 125);
            this.txt_nombre.Name = "txt_nombre";
            this.txt_nombre.Size = new System.Drawing.Size(340, 20);
            this.txt_nombre.TabIndex = 5;
            // 
            // txt_bbdd
            // 
            this.txt_bbdd.Location = new System.Drawing.Point(83, 70);
            this.txt_bbdd.Name = "txt_bbdd";
            this.txt_bbdd.Size = new System.Drawing.Size(435, 20);
            this.txt_bbdd.TabIndex = 4;
            // 
            // txt_ruta
            // 
            this.txt_ruta.Location = new System.Drawing.Point(56, 44);
            this.txt_ruta.Name = "txt_ruta";
            this.txt_ruta.Size = new System.Drawing.Size(462, 20);
            this.txt_ruta.TabIndex = 3;
            // 
            // txt_proceso
            // 
            this.txt_proceso.Location = new System.Drawing.Point(173, 17);
            this.txt_proceso.Name = "txt_proceso";
            this.txt_proceso.Size = new System.Drawing.Size(32, 20);
            this.txt_proceso.TabIndex = 2;
            // 
            // dgv_reglas
            // 
            this.dgv_reglas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_reglas.Enabled = false;
            this.dgv_reglas.Location = new System.Drawing.Point(6, 568);
            this.dgv_reglas.Name = "dgv_reglas";
            this.dgv_reglas.Size = new System.Drawing.Size(511, 157);
            this.dgv_reglas.TabIndex = 0;
            // 
            // btn_ejecuta_todo
            // 
            this.btn_ejecuta_todo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_ejecuta_todo.Location = new System.Drawing.Point(599, 793);
            this.btn_ejecuta_todo.Name = "btn_ejecuta_todo";
            this.btn_ejecuta_todo.Size = new System.Drawing.Size(164, 23);
            this.btn_ejecuta_todo.TabIndex = 33;
            this.btn_ejecuta_todo.Text = "Ejecuta todos los activos";
            this.btn_ejecuta_todo.UseVisualStyleBackColor = true;
            this.btn_ejecuta_todo.Click += new System.EventHandler(this.btn_ejecuta_todo_Click);
            // 
            // btn_solo_actual
            // 
            this.btn_solo_actual.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_solo_actual.Location = new System.Drawing.Point(769, 793);
            this.btn_solo_actual.Name = "btn_solo_actual";
            this.btn_solo_actual.Size = new System.Drawing.Size(138, 23);
            this.btn_solo_actual.TabIndex = 34;
            this.btn_solo_actual.Text = "Ejecutar Sólo Actual";
            this.btn_solo_actual.UseVisualStyleBackColor = true;
            this.btn_solo_actual.Click += new System.EventHandler(this.btn_solo_actual_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(204, 6);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(207, 22);
            this.toolStripMenuItem1.Text = "Ejecutar todos los activos";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // ejecutarSóloActualToolStripMenuItem
            // 
            this.ejecutarSóloActualToolStripMenuItem.Name = "ejecutarSóloActualToolStripMenuItem";
            this.ejecutarSóloActualToolStripMenuItem.Size = new System.Drawing.Size(207, 22);
            this.ejecutarSóloActualToolStripMenuItem.Text = "Ejecutar Sólo Actual";
            this.ejecutarSóloActualToolStripMenuItem.Click += new System.EventHandler(this.ejecutarSóloActualToolStripMenuItem_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(919, 825);
            this.Controls.Add(this.btn_solo_actual);
            this.Controls.Add(this.btn_ejecuta_todo);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Main";
            this.Text = "Cola de Procesos";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Main_FormClosing);
            this.Load += new System.EventHandler(this.Main_Load);
            this.groupBox1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_dependencias)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_reglas)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_reglas;
        private System.Windows.Forms.Button btn_eliminar;
        private System.Windows.Forms.Button btn_modificar;
        private System.Windows.Forms.Button btn_agregar;
        private System.Windows.Forms.TreeView treeProcesos;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editarToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chk_obligatorio;
        private System.Windows.Forms.CheckBox chk_activo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgv_dependencias;
        private System.Windows.Forms.ComboBox cmb_tipoParametro;
        private System.Windows.Forms.ComboBox cmb_periodicidad;
        private System.Windows.Forms.TextBox txt_descripcion;
        private System.Windows.Forms.ComboBox cmb_tipoProceso;
        private System.Windows.Forms.TextBox txt_nombre;
        private System.Windows.Forms.TextBox txt_bbdd;
        private System.Windows.Forms.TextBox txt_ruta;
        private System.Windows.Forms.TextBox txt_proceso;
        private System.Windows.Forms.DataGridView dgv_reglas;
        private System.Windows.Forms.Button btn_dependencias;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txt_ultimo_ok;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txt_ultimo_lanzamiento;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btn_aceptar;
        private System.Windows.Forms.Button btn_cancelar;
        private System.Windows.Forms.TextBox txt_grupo;
        private System.Windows.Forms.Button btn_ejecuta_todo;
        private System.Windows.Forms.Button btn_solo_actual;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txt_mensaje;
        private System.Windows.Forms.ToolStripMenuItem herramientasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarColaDeProcesosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
        private System.Windows.Forms.TextBox txt_parametros;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem ejecutarSóloActualToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
    }
}

