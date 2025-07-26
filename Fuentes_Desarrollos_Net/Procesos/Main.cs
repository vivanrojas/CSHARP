using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Procesos
{
    public partial class Main : Form
    {
        business.Parametricas param;
        business.Procesos procesos;
        business.MSAccess access;
        int tiempo = 5;
        bool ejecuta_ahora;
        public string cola { get; set; }

        
        public Main(string _cola, bool ejecuta_now)
        {
            

            ejecuta_ahora = ejecuta_now;
            cola = _cola;
            InitializeComponent();            
        }

        private void Main_Load(object sender, EventArgs e)
        {
            
            param = new business.Parametricas();
            access = new business.MSAccess();        
            
            CargaCombos();
            LeeArbol();

            EjecutaCola(ejecuta_ahora);

            
        }

        private void EjecutaCola(bool ejecuta_now)
        {

            if (cola != "" && cola != null && ejecuta_now)
            {
                timer1.Enabled = true;              

            }
            else
            {                
                
                business.Cola cola = new business.Cola();
                if (cola.HayNuevaCola())
                    this.Text = "Cola de procesos de " + cola.UltimaColaCreada();
                


            }
        }


        private void RecorreArbol()
        {
            foreach(TreeNode tn in treeProcesos.Nodes)
            {
                BusquedaRecursiva(tn);
            }
        }

        private void BusquedaRecursiva(TreeNode treeNode)
        {
            foreach (TreeNode tn in treeNode.Nodes)
            {

                if (tn.Name.Contains(";"))
                {

                    string[] posicion = tn.Name.Split(';');
                    int grupo = Convert.ToInt32(posicion[0]);
                    int proceso = Convert.ToInt32(posicion[1]);


                    EndesaEntity.cola.ProcesoCola c = procesos.GetProceso(grupo, proceso);
                    c.activo = tn.Checked;
                    procesos.GuardaDatos(c);

                    //c.cola = cola;
                    //c.grupo = tn.
                    //c.proceso = Convert.ToInt32(txt_proceso.Text);
                    //c.id_p_proceso = param.GetIDProceso(cmb_tipoProceso.Text);
                    //c.ruta = txt_ruta.Text;
                    //c.bbdd = txt_bbdd.Text;
                    //c.nombre_proceso = txt_nombre.Text;
                    //c.activo = chk_activo.Checked;
                    //c.obligatorio = chk_obligatorio.Checked;
                    //c.una_vez = false;
                    //c.descripcion = txt_descripcion.Text;

                    //if (txt_ultimo_ok.Text != "")
                    //    c.fecha_ultima_ejec_ok = Convert.ToDateTime(txt_ultimo_ok.Text);

                    //if (txt_ultimo_lanzamiento.Text != "")
                    //    c.fecha_ultimo_lanzamiento = Convert.ToDateTime(txt_ultimo_lanzamiento.Text);

                    //c.id_p_periodicidad = param.GetIDPeriodicidad(cmb_periodicidad.Text);
                    //c.id_p_parametro = param.GetIDParametro(cmb_tipoParametro.Text);
                    //c.parametro = txt_parametros.Text;
                    //c.mensaje_error = txt_mensaje.Text;

                }
                BusquedaRecursiva(tn);
            }
        }

        private void LeeArbol()
        {
            TreeNode nodo;
            TreeNode nodoHijo;

            string nGrupo = "";
            string proceso = "";
            string descripcion = "";
            int inodo = 0;

            Controles(false);
            treeProcesos.Nodes.Clear();
            treeProcesos.CheckBoxes = true;

            procesos = new business.Procesos(cola);

            foreach (KeyValuePair<int, List<EndesaEntity.cola.ProcesoCola>> p in procesos.dic)
            {

                nGrupo = p.Key.ToString();
                nodo = new TreeNode();
                nodo.Name = nGrupo;
                nodo.Text = "Grupo: [" + nGrupo + "]";
                treeProcesos.Nodes.Add(nodo);
                for (int i = 0; i < p.Value.Count; i++)
                {
                    proceso = p.Value[i].proceso.ToString();
                    descripcion = p.Value[i].descripcion;

                    nodoHijo = new TreeNode();
                    nodoHijo.ForeColor = procesos.ColorNodo(p.Value[i]);
                    nodoHijo.Checked = p.Value[i].activo;
                    nodoHijo.Name = nGrupo + ";" + proceso;
                    nodoHijo.Text = proceso + "-" + p.Value[i].descripcion;
                    treeProcesos.Nodes[inodo].Nodes.Add(nodoHijo);
                }
                inodo++;
            }

            treeProcesos.ExpandAll();
        }

        
        private void treeProcesos_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeProcesos.SelectedNode.Name.Contains(";"))
            {

                //treeProcesos.SelectedNode.BackColor = Color.LightGoldenrodYellow;
                string[] posicion = treeProcesos.SelectedNode.Name.Split(';');

                int grupo = Convert.ToInt32(posicion[0]);
                int proceso = Convert.ToInt32(posicion[1]);

                EndesaEntity.cola.ProcesoCola nodo = procesos.GetProceso(grupo, proceso);
                
                if (nodo != null)
                {
                    txt_grupo.Text = nodo.grupo.ToString();
                    txt_proceso.Text = nodo.proceso.ToString();
                    chk_activo.Checked = nodo.activo;
                    chk_obligatorio.Checked = nodo.obligatorio;
                    txt_ruta.Text = nodo.ruta;
                    txt_bbdd.Text = nodo.bbdd;
                    txt_nombre.Text = nodo.nombre_proceso;
                    txt_descripcion.Text = nodo.descripcion;

                    cmb_tipoProceso.SelectedItem = param.GetProceso(nodo.id_p_proceso);
                    cmb_periodicidad.SelectedItem = param.GetPeriodicidad(nodo.id_p_periodicidad);
                    cmb_tipoParametro.SelectedItem = param.GetParametro(nodo.id_p_parametro);
                    txt_parametros.Text = nodo.parametro;


                    if(nodo.fecha_ultimo_lanzamiento > DateTime.MinValue)
                        txt_ultimo_lanzamiento.Text = nodo.fecha_ultimo_lanzamiento.ToString("dd/MM/yyyy HH:mm:ss");

                    if (nodo.fecha_ultima_ejec_ok > DateTime.MinValue)
                        txt_ultimo_ok.Text = nodo.fecha_ultima_ejec_ok.ToString("dd/MM/yyyy HH:mm:ss");

                    txt_mensaje.Text = nodo.mensaje_error;
                }
            }
        }

        private void btn_solo_actual_Click(object sender, EventArgs e)
        {
            if (treeProcesos.SelectedNode.Name.Contains(";"))
            {
                DialogResult result = MessageBox.Show("¿Quiere ejecutar el proceso actual?", "Procesos",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string[] posicion = treeProcesos.SelectedNode.Name.Split(';');
                    int grupo = Convert.ToInt32(posicion[0]);
                    int proceso = Convert.ToInt32(posicion[1]);
                    EndesaEntity.cola.ProcesoCola nodo = procesos.GetProceso(grupo, proceso);
                    procesos.EjecutaNodo(nodo, true);
                    procesos = new business.Procesos("Contratación");
                    LeeArbol();
                }
                
            }
        }

        private void CargaCombos()
        {
            List<string> lista = param.dic_periodicidad.Select(z => z.Value).ToList();
            for (int i = 0; i < lista.Count; i++)
                cmb_periodicidad.Items.Add(lista[i]);

            lista = param.dic_tipoParametro.Select(z => z.Value).ToList();
            for (int i = 0; i < lista.Count; i++)
                cmb_tipoParametro.Items.Add(lista[i]);

            lista = param.dic_procesos.Select(z => z.Value).ToList();
            for (int i = 0; i < lista.Count; i++)
                cmb_tipoProceso.Items.Add(lista[i]);

        }

        private void treeProcesos_AfterCheck(object sender, TreeViewEventArgs e)
        {
            chk_activo.Checked = treeProcesos.SelectedNode.Checked;
        }

        private void btn_ejecuta_todo_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("¿Quiere ejecutar todos los procesos marcados?", "Procesos",
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                RecorreArbol();
                procesos.EjecutaCola();
                LeeArbol();
            }
                
        }

        private void btn_modificar_Click(object sender, EventArgs e)
        {
            Controles(true);

        }

       

        private void Controles(bool habilitar)
        {
            txt_grupo.Enabled = habilitar;
            txt_proceso.Enabled = habilitar;
            chk_activo.Enabled = habilitar;
            chk_obligatorio.Enabled = habilitar;
            txt_ruta.Enabled = habilitar;
            txt_bbdd.Enabled = habilitar;
            cmb_tipoProceso.Enabled = habilitar;
            txt_nombre.Enabled = habilitar;
            txt_descripcion.Enabled = habilitar;
            cmb_periodicidad.Enabled = habilitar;
            cmb_tipoParametro.Enabled = habilitar;

            btn_dependencias.Enabled = habilitar;
            txt_parametros.Enabled = habilitar;

            btn_aceptar.Enabled = habilitar;
            btn_cancelar.Enabled = habilitar;

        }


        private EndesaEntity.cola.ProcesoCola CapturaDatos()
        {
            EndesaEntity.cola.ProcesoCola c = new EndesaEntity.cola.ProcesoCola();
            c.cola = cola;
            c.grupo = Convert.ToInt32(txt_grupo.Text);
            c.proceso = Convert.ToInt32(txt_proceso.Text);
            c.id_p_proceso = param.GetIDProceso(cmb_tipoProceso.Text);
            c.ruta = txt_ruta.Text;
            c.bbdd = txt_bbdd.Text;
            c.nombre_proceso = txt_nombre.Text;
            c.activo = chk_activo.Checked;
            c.obligatorio = chk_obligatorio.Checked;
            c.una_vez = false;
            c.descripcion = txt_descripcion.Text;

            if (txt_ultimo_ok.Text != "")
                c.fecha_ultima_ejec_ok = Convert.ToDateTime(txt_ultimo_ok.Text);

            if (txt_ultimo_lanzamiento.Text != "")
                c.fecha_ultimo_lanzamiento = Convert.ToDateTime(txt_ultimo_lanzamiento.Text);

            c.id_p_periodicidad = param.GetIDPeriodicidad(cmb_periodicidad.Text);
            c.id_p_parametro = param.GetIDParametro(cmb_tipoParametro.Text);
            c.parametro = txt_parametros.Text;
            c.mensaje_error = txt_mensaje.Text;

            return c;
        }
        private void btn_aceptar_Click(object sender, EventArgs e)
        {
            Controles(false);
            
            procesos.GuardaDatos(CapturaDatos());
            Controles(false);
            LeeArbol();
        }

        private void LimpiaDatos()
        {
            txt_grupo.Text = "";
            txt_proceso.Text = "";
            txt_ruta.Text = "";
            txt_bbdd.Text = "";
        }

        private void btn_agregar_Click(object sender, EventArgs e)
        {
            Controles(true);
            LimpiaDatos();           

            
        }

        private void btn_eliminar_Click(object sender, EventArgs e)
        {
            if (treeProcesos.SelectedNode.Name.Contains(";"))
            {
                DialogResult result = MessageBox.Show("¿Quiere borrar el proceso actual?", "Procesos",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    procesos.EliminaNodo(CapturaDatos());
                    LeeArbol();
                }
            }
        }

        private void importarColaDeProcesosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Este proceso importará una cola de procesos Access"
                + System.Environment.NewLine
                + "¿Desea continurar?",
                "Importar Cola de Procesos",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                procesos.ImportarCola();
            }
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_cancelar_Click(object sender, EventArgs e)
        {
            Controles(false);
            LeeArbol();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            tiempo--;
            if(tiempo == 0)
            {
                this.Text = "Cola de procesos de " + cola;
                procesos = new business.Procesos(cola);
                if (ejecuta_ahora)
                {
                    procesos.EjecutaCola();
                    this.Close();
                }                
                    
            }

        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            RecorreArbol();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("¿Quiere ejecutar todos los procesos marcados?", "Procesos",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                RecorreArbol();
                procesos.EjecutaCola();
            }
        }

        private void ejecutarSóloActualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeProcesos.SelectedNode.Name.Contains(";"))
            {
                DialogResult result = MessageBox.Show("¿Quiere ejecutar el proceso actual?", "Procesos",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string[] posicion = treeProcesos.SelectedNode.Name.Split(';');
                    int grupo = Convert.ToInt32(posicion[0]);
                    int proceso = Convert.ToInt32(posicion[1]);
                    EndesaEntity.cola.ProcesoCola nodo = procesos.GetProceso(grupo, proceso);
                    procesos.EjecutaNodo(nodo, true);
                    procesos = new business.Procesos("Contratación");
                    LeeArbol();
                }
            }
        }
    }
}
