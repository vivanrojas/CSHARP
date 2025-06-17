using EndesaBusiness.servidores;
using EndesaBusiness.sigame;
using EndesaEntity.contratacion.gas;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class ContratosGasDetalle : EndesaEntity.Table_atrgas_contratos_detalle
    {
        public Dictionary<string, List<EndesaEntity.Table_atrgas_contratos_detalle>> dic { get; set; }
        public EndesaEntity.Table_atrgas_contratos_detalle last_data { get; set; }

        logs.Log ficheroLog;

        public ContratosGasDetalle()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ATRGas");
            last_data = new EndesaEntity.Table_atrgas_contratos_detalle();
            dic = new Dictionary<string, List<EndesaEntity.Table_atrgas_contratos_detalle>>();            
            this.GetAll();
        }

        public void Save(bool nuevo_registro)
        {
            try
            {

                if (nuevo_registro)
                    this.New();
                else
                    this.Update();              

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "ContratosGasDetalle - Save",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        public bool ExisteRegistro(string nif, string cups20, DateTime f, string tipo, double qd, string comentario)
        {
            bool existe = false;
            List<EndesaEntity.Table_atrgas_contratos_detalle> lista;
            if (dic.TryGetValue(nif, out lista))
            {
                existe = lista.Exists(z => z.cups20 == cups20
                    && z.tipo == tipo
                    && z.fecha_inicio == f
                    && z.qd == qd
                    && z.comentario == comentario);
            }
            return existe;
        }

        public bool ExisteRegistro(string nif, string cups20, DateTime f, string tipo)
        {
            bool existe = false;
            List<EndesaEntity.Table_atrgas_contratos_detalle> lista;
            if (dic.TryGetValue(nif, out lista))
            {
                existe = lista.Exists(z => z.cups20 == cups20
                    && z.tipo == tipo
                    && z.fecha_inicio == f);
            }
            return existe;
        }

        public void GetAll()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select nif, cups20, fecha_inicio, fecha_fin, qd, tarifa, tipo, comentario,"
                    + " id_solicitud, linea_solicitud, qi, hora_inicio,"
                    + " created_by, creation_date, last_update_by, last_update_date from atrgas_contratos_detalle"
                    + " order by fecha_inicio desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Table_atrgas_contratos_detalle c = new EndesaEntity.Table_atrgas_contratos_detalle();

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["qd"] != System.DBNull.Value)
                        c.qd = Convert.ToDouble(r["qd"]);

                    if (r["qi"] != System.DBNull.Value)
                        c.qi = Convert.ToDouble(r["qi"]);

                    if (r["hora_inicio"] != System.DBNull.Value)
                        c.hora_inicio = Convert.ToDateTime(r["hora_inicio"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["tipo"] != System.DBNull.Value)
                        c.tipo = r["tipo"].ToString();

                    if (r["created_by"] != System.DBNull.Value)
                        c.created_by = r["created_by"].ToString();

                    if (r["creation_date"] != System.DBNull.Value)
                        c.creation_date = Convert.ToDateTime(r["creation_date"]);

                    if (r["last_update_by"] != System.DBNull.Value)
                        c.last_update_by = r["last_update_by"].ToString();

                    if (r["comentario"] != System.DBNull.Value)
                        c.comentario = r["comentario"].ToString();

                    if (r["id_solicitud"] != System.DBNull.Value)
                        c.id_solicitud = Convert.ToInt64(r["id_solicitud"]);

                    if (r["linea_solicitud"] != System.DBNull.Value)
                        c.linea_solicitud = Convert.ToInt32(r["linea_solicitud"]);

                    List<EndesaEntity.Table_atrgas_contratos_detalle> cc = new List<EndesaEntity.Table_atrgas_contratos_detalle>();
                    if (!dic.TryGetValue(c.nif, out cc))
                    {
                        List<EndesaEntity.Table_atrgas_contratos_detalle> lista = new List<EndesaEntity.Table_atrgas_contratos_detalle>();
                        lista.Add(c);
                        dic.Add(c.nif, lista);
                    }
                    else
                    {
                        cc.Add(c);
                    }

                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("ContratosGasDetalle - GetAll " + e.Message);
            }

        }

        public void GetAll(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select nif, cups20, fecha_inicio, fecha_fin, qd, tarifa, tipo, comentario,"
                    + " id_solicitud, linea_solicitud, qi, hora_inicio,"
                    + " created_by, creation_date, last_update_by, last_update_date from atrgas_contratos_detalle"
                    + " where fecha_inicio <= '"
                    + " order by fecha_inicio desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Table_atrgas_contratos_detalle c = new EndesaEntity.Table_atrgas_contratos_detalle();

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["qd"] != System.DBNull.Value)
                        c.qd = Convert.ToDouble(r["qd"]);

                    if (r["qi"] != System.DBNull.Value)
                        c.qi = Convert.ToDouble(r["qi"]);

                    if (r["hora_inicio"] != System.DBNull.Value)
                        c.hora_inicio = Convert.ToDateTime(r["hora_inicio"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["tipo"] != System.DBNull.Value)
                        c.tipo = r["tipo"].ToString();

                    if (r["created_by"] != System.DBNull.Value)
                        c.created_by = r["created_by"].ToString();

                    if (r["creation_date"] != System.DBNull.Value)
                        c.creation_date = Convert.ToDateTime(r["creation_date"]);

                    if (r["last_update_by"] != System.DBNull.Value)
                        c.last_update_by = r["last_update_by"].ToString();

                    if (r["comentario"] != System.DBNull.Value)
                        c.comentario = r["comentario"].ToString();

                    if (r["id_solicitud"] != System.DBNull.Value)
                        c.id_solicitud = Convert.ToInt64(r["id_solicitud"]);

                    if (r["linea_solicitud"] != System.DBNull.Value)
                        c.linea_solicitud = Convert.ToInt32(r["linea_solicitud"]);

                    List<EndesaEntity.Table_atrgas_contratos_detalle> cc = new List<EndesaEntity.Table_atrgas_contratos_detalle>();
                    if (!dic.TryGetValue(c.nif, out cc))
                    {
                        List<EndesaEntity.Table_atrgas_contratos_detalle> lista = new List<EndesaEntity.Table_atrgas_contratos_detalle>();
                        lista.Add(c);
                        dic.Add(c.nif, lista);
                    }
                    else
                    {
                        cc.Add(c);
                    }

                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("ContratosGasDetalle - GetAll " + e.Message);
            }

        }

        public void Replace()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "replace into atrgas_contratos_detalle (nif, cups20,  fecha_inicio,"
                   + " fecha_fin, qd, tarifa, qi, hora_inicio, tipo, comentario, id_solicitud, linea_solicitud, created_by, creation_date) values ";

                if (this.nif != null)
                    strSql += "('" + this.nif + "'";
                else
                    strSql += "(null";

                if (this.cups20 != null)
                    strSql += ",'" + this.cups20 + "'";
                else
                    strSql += ",null";

                if (this.fecha_inicio != null)
                    strSql += ",'" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.fecha_fin > DateTime.MinValue)
                    strSql += ",'" + this.fecha_fin.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.qd != 0)
                    strSql += "," + this.qd.ToString().Replace(",", ".");
                else
                    strSql += ",0";

                if (this.tarifa != null)
                    strSql += ",'" + this.tarifa + "'";
                else
                    strSql += ",null";

                if (this.qi != 0)
                    strSql += "," + this.qi.ToString().Replace(",", ".");
                else
                    strSql += ",null";

                if (this.hora_inicio > DateTime.MinValue)
                    strSql += ",'" + this.hora_inicio.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                else
                    strSql += ",null";

                if (this.tipo != null)
                    strSql += ",'" + this.tipo + "'";
                else
                    strSql += ",null";

                if (this.comentario != null)
                    strSql += ",'" + this.comentario + "'";
                else
                    strSql += ",''";

                if (this.id_solicitud != 0)
                    strSql += "," + this.id_solicitud;
                else
                    strSql += ",null";

                if (this.linea_solicitud != 0)
                    strSql += "," + this.linea_solicitud;
                else
                    strSql += ",null";


                strSql += ",'" + System.Environment.UserName + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message,
                  "ContratosGasDetalle - Replace",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
                }
        }

        public void New()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "insert into atrgas_contratos_detalle (nif, cups20,  fecha_inicio,"
                    + " fecha_fin, qd, tarifa, qi, hora_inicio, tipo, comentario, id_solicitud, linea_solicitud, created_by, creation_date) values ";

                if (this.nif != null)
                    strSql += "('" + this.nif + "'";
                else
                    strSql += "(null";

                if (this.cups20 != null)
                    strSql += ",'" + this.cups20 + "'";
                else
                    strSql += ",null";
                

                if (this.fecha_inicio != null)
                    strSql += ",'" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.fecha_fin > DateTime.MinValue)
                    strSql += ",'" + this.fecha_fin.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.qd != 0)
                    strSql += "," + this.qd.ToString().Replace(",", ".");
                else
                    strSql += ",0";

                if (this.tarifa != null)
                    strSql += ",'" + this.tarifa + "'";
                else
                    strSql += ",null";

                if (this.qi != 0)
                    strSql += "," + this.qi.ToString().Replace(",", ".");
                else
                    strSql += ",null";

                if (this.hora_inicio > DateTime.MinValue)
                    strSql += ",'" + this.hora_inicio.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                else
                    strSql += ",null";

                if (this.tipo != null)
                    strSql += ",'" + this.tipo + "'";
                else
                    strSql += ",null";

                if (this.comentario != null)
                    strSql += ",'" + this.comentario + "'";
                else
                    strSql += ",''";

                if (this.id_solicitud != 0)
                    strSql += "," + this.id_solicitud;
                else
                    strSql += ",null";

                if (this.linea_solicitud != 0)
                    strSql += "," + this.linea_solicitud;
                else
                    strSql += ",null";

                strSql += ",'" + System.Environment.UserName + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "ContratosGasDetalle - New",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

        public void Update()
        {


            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update atrgas_contratos_detalle set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";


            if (this.cups20 != null)
                strSql += " ,cups20 = '" + this.cups20 + "'";

            if (this.fecha_inicio != null)
                strSql += " ,fecha_inicio = '" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";

            if (this.fecha_fin > DateTime.MinValue)
                strSql += " ,fecha_fin = '" + this.fecha_fin.ToString("yyyy-MM-dd") + "'";
            else
                strSql += " ,fecha_fin = null";

            if (this.qd != 0)
                strSql += " ,qd = " + this.qd.ToString().Replace(",", ".");
            else
                strSql += " ,qd = 0";

            if (this.qi != 0)
                strSql += " ,qi = " + this.qd.ToString().Replace(",", ".");
            else
                strSql += " ,qi = null";

            if (this.hora_inicio > DateTime.MinValue)
                strSql += " ,hora_inicio = '" + this.fecha_fin.ToString("yyyy-MM-dd HH:mm:ss") + "'";
            else
                strSql += " ,hora_inicio = null";

            if (this.tarifa != null)
                strSql += " ,tarifa = '" + this.tarifa + "'";
                        

            if (this.tipo != null)
                strSql += " ,tipo = '" + this.tipo.ToUpper() + "'";

            if (this.comentario != null)
                strSql += " ,comentario = '" + this.comentario + "'";

            if (this.id_solicitud != 0)
                strSql += ",id_solicitud = " + this.id_solicitud;
            
            if (this.linea_solicitud != 0)
                strSql += ",linea_solicitud = " + this.linea_solicitud;
            

            strSql += " where nif = '" + last_data.nif + "'"
                + " and cups20 = '" + last_data.cups20 + "'"
                + " and fecha_inicio = '" + last_data.fecha_inicio.ToString("yyyy-MM-dd") + "'"
                + " and tipo = '" + last_data.tipo + "'";

            if (this.comentario != null)
                strSql += " and comentario = '" + last_data.comentario + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        

        public void Del()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                DialogResult result = MessageBox.Show("Warning:" + System.Environment.NewLine
                    + "La línea del contrato será borrada." + System.Environment.NewLine
                    + "¿Desea borrar la línea del contrato?"
                    , "Borrado de linea de contrato " + this.cups20 + " " + this.tipo + " - "
                    + this.fecha_inicio.ToString("dd/MM/yyyy") + " - " + this.fecha_fin.ToString("dd/MM/yyyy")
                    + this.tipo + " - " + this.comentario,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    strSql = "delete from atrgas_contratos_detalle where"
                         + " nif = '" + this.nif + "'"
                         + " and cups20 = '" + this.cups20 + "'"
                         + " and fecha_inicio = '" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'"
                         + " and tipo = '" + this.tipo + "'";

                    if(this.qd != 0)
                        strSql += " and qd = " + this.qd;

                    if (this.qi != 0)
                        strSql += " and qi = " + this.qd;

                    if (this.comentario != null)
                        strSql += " and comentario = '" + this.comentario + "'";


                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("Línea de contrato borrada.",
                  "Contratos Detalle GAS",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "ContratosGasDetalle - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }


        }

        public void Del(string nif, string cups20, string tipo, DateTime fecha_inicio, DateTime fecha_fin)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
               
                
                strSql = "delete from atrgas_contratos_detalle where"
                        + " nif = '" + nif + "'"
                        + " and cups20 = '" + cups20 + "'"
                        + " and fecha_inicio = '" + fecha_inicio.ToString("yyyy-MM-dd") + "'"
                        + " and fecha_fin = '" + fecha_fin.ToString("yyyy-MM-dd") + "'"
                        + " and tipo = '" + tipo + "'";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                   
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "ContratosGasDetalle - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }


        }        
        

        public List<EndesaEntity.Informe_contratos_gas_vencimiento> Lista_Proximo_Vencimiento(int dias)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            List<EndesaEntity.Informe_contratos_gas_vencimiento> l = new List<EndesaEntity.Informe_contratos_gas_vencimiento>();
            List<string> lista_nifs = new List<string>();
            cartera.Cartera_SalesForce cartera;
            DateTime fecha = new DateTime();
            string cups = "";
            DateTime fecha_maxima = new DateTime();
            

            fecha = DateTime.Now.AddDays(dias);
            Dictionary<string, DateTime> dic_fechas_maximas = new Dictionary<string, DateTime>();

            // Primero sacamos todos los contratos máximos
            strSql = "SELECT d.cups20, fecha_fin FROM cont.atrgas_contratos_detalle d";
                //+ " where d.cups20 = 'ES0217901000011537CV'";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                cups = r["cups20"].ToString();
                if (r["fecha_fin"] != System.DBNull.Value)
                    fecha_maxima = Convert.ToDateTime(r["fecha_fin"]);
                else
                    fecha_maxima = new DateTime(4999, 12, 31);

                DateTime o;
                if(!dic_fechas_maximas.TryGetValue(cups, out o))
                {
                    dic_fechas_maximas.Add(cups, fecha_maxima);
                }
                else
                {
                    if (o < fecha_maxima)
                        dic_fechas_maximas[cups] = fecha_maxima;
                }
                
            }
            db.CloseConnection();


            strSql = "select d.cnifdnic"
                + " FROM atrgas_contratos d INNER JOIN"
                + " atrgas_contratos_detalle dd ON"
                + " dd.nif = d.cnifdnic"
                + " WHERE(DATEDIFF(dd.fecha_fin, NOW()) > 0 AND DATEDIFF(dd.fecha_fin, NOW()) < " + dias + ")";                
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
                if (r["cnifdnic"] != System.DBNull.Value)
                    lista_nifs.Add(r["cnifdnic"].ToString());

            db.CloseConnection();

            if (lista_nifs.Count > 0)
            {
                cartera = new cartera.Cartera_SalesForce(lista_nifs);

                strSql = "select d.cnifdnic, d.dapersoc, d.distribuidora,"
                    + " dd.cups20, dd.tipo, dd.fecha_inicio, dd.fecha_fin, dd.qd, dd.tarifa,"
                    + " dd.qi, dd.hora_inicio"
                    + " FROM atrgas_contratos d INNER JOIN"
                    + " atrgas_contratos_detalle dd ON"
                    + " dd.nif = d.cnifdnic and"
                    + " dd.cups20 = d.cups20"
                    + " WHERE(DATEDIFF(dd.fecha_fin, NOW()) > 0 AND DATEDIFF(dd.fecha_fin, NOW()) < " + dias + ")"
                    + " ORDER BY dd.fecha_fin, d.cnifdnic";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Informe_contratos_gas_vencimiento c = new EndesaEntity.Informe_contratos_gas_vencimiento();

                    if (r["cnifdnic"] != System.DBNull.Value)
                    {
                        c.nif = r["cnifdnic"].ToString();
                        cartera.GetCartera(c.nif);
                        c.responsable_territorial = cartera.responsable_territorial;
                        c.responsable_territorial = c.responsable_territorial.ToUpper();
                        c.gestor = cartera.gestor;
                        c.gestor = c.gestor.ToUpper();
                    }

                    if (r["dapersoc"] != System.DBNull.Value)
                        c.cliente = r["dapersoc"].ToString();

                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["qd"] != System.DBNull.Value)
                        c.qd = Convert.ToDouble(r["qd"]);

                    if (r["qi"] != System.DBNull.Value)
                        c.qi = Convert.ToDouble(r["qi"]);

                    if (r["hora_inicio"] != System.DBNull.Value)
                        c.hora_inicio = Convert.ToDateTime(r["hora_inicio"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["tipo"] != System.DBNull.Value)
                        c.tipo = r["tipo"].ToString();

                    if (r["distribuidora"] != System.DBNull.Value)
                        c.distribuidora = r["distribuidora"].ToString();

                    


                    if (FechaMaximaProductoContrato(dic_fechas_maximas, c.cups20) > DateTime.Now.AddDays(dias))
                    {
                        c.continua = "Sí";
                    }                     
                    else
                        c.continua = "No";


                    l.Add(c);

                }
                db.CloseConnection();
            }
            return l;
        }

        public List<EndesaEntity.Informe_contratos_gas_vencimiento> Lista_Proximo_Vencimiento_SIGAME(int dias)
        {
            string strSql;
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;

            List<EndesaEntity.Informe_contratos_gas_vencimiento> l = new List<EndesaEntity.Informe_contratos_gas_vencimiento>();
            List<string> lista_nifs = new List<string>();
            cartera.Cartera_SalesForce cartera;
            DateTime fecha = new DateTime();
            string cups = "";
            DateTime fecha_maxima = new DateTime();


            fecha = DateTime.Now.AddDays(dias);
            DateTime fecha_calculo = new DateTime();

            Dictionary<string, DateTime> dic_fechas_maximas = new Dictionary<string, DateTime>();

            // Primero sacamos todos los contratos máximos
            strSql = "SELECT DISTINCT"
                + " T_SGM_M_CLIENTES.CD_CIF,"
                + " T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
                + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES,"
                + " T_SGM_G_PS.CD_CUPS,"
                + " T_SGM_P_PEAJE_DURACION.DE_PEAJE_DURACION,"
                + " T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO,"
                + " T_SGM_M_ADDENDA_CTT_DIST.FH_FIN,"
                + " T_SGM_M_ADDENDA_CTT_DIST.NM_CDC as Qd,"
                + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA,"
                + " T_SGM_M_GESTORES.DE_GESTOR"
                + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
                + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                + " LEFT JOIN T_SGM_M_CLIENTES ON"
                + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                + " LEFT JOIN T_SGM_M_GESTORES ON"
                + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR"
                + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
                + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
                + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
                + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES"
                + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON"
                + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
                + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
                + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
                + " T_SGM_G_PS.CD_PAIS = T_SGM_P_PROVINCIAS.CD_PAIS"
                + " LEFT JOIN T_SGM_P_MUNICIPIOS ON"
                + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
                + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS"
                + " INNER JOIN T_SGM_M_ADDENDA_CTT_DIST ON"
                + " T_SGM_M_ADDENDA_CTT_DIST.ID_PS = T_SGM_G_PS.ID_PS"
                + " INNER JOIN T_SGM_P_PEAJE_DURACION ON"
                + " T_SGM_P_PEAJE_DURACION.ID_PEAJE_DURACION = T_SGM_M_ADDENDA_CTT_DIST.ID_PEAJE_DURACION"
                + " WHERE T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO in (3, 6, 7, 8, 9, 10, 15, 16)"
                + " AND T_SGM_G_PS.CD_PAIS = 'ESP' AND T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA <> 'Cisternas'"
                + " AND T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                + " AND(T_SGM_G_CONTRATOS_PS.FH_FIN_REAL >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR T_SGM_G_CONTRATOS_PS.FH_FIN_REAL IS NULL)"
                + " AND T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO <= '" + DateTime.Now.ToString("yyyy-MM-dd")  + "'"
                + " AND(T_SGM_M_ADDENDA_CTT_DIST.FH_FIN >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR T_SGM_M_ADDENDA_CTT_DIST.FH_FIN is null)"
                + " order by T_SGM_G_PS.CD_CUPS";
            //+ " where d.cups20 = 'ES0217901000011537CV'";
            db = new SQLServer();
            command = new SqlCommand(strSql, db.con);
            SqlDataAdapter da = new SqlDataAdapter(command);
            r = command.ExecuteReader();
            while (r.Read())
            {
                cups = r["CD_CUPS"].ToString();
                if (r["FH_FIN"] != System.DBNull.Value)
                    fecha_maxima = Convert.ToDateTime(r["FH_FIN"]);
                else
                {
                    if (r["DE_PEAJE_DURACION"].ToString().ToUpper() == "MENSUAL")
                    {
                        fecha_calculo = Convert.ToDateTime(r["FH_INICIO"]);
                        fecha_maxima = fecha_calculo.AddDays(DateTime.DaysInMonth(fecha_calculo.Year, fecha_calculo.Month) - 1);
                    }else                    
                        fecha_maxima = new DateTime(4999, 12, 31);
                }
                    

                DateTime o;
                if (!dic_fechas_maximas.TryGetValue(cups, out o))
                {
                    dic_fechas_maximas.Add(cups, fecha_maxima);
                }
                else
                {
                    if (o < fecha_maxima)
                        dic_fechas_maximas[cups] = fecha_maxima;
                }

            }
            db.CloseConnection();


            cartera = new cartera.Cartera_SalesForce(Lista_NIF());


            strSql = "SELECT DISTINCT"
            + " T_SGM_M_CLIENTES.CD_CIF,"
            + " T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
            + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES,"
            + " T_SGM_G_PS.CD_CUPS,"
            + " T_SGM_P_PEAJE_DURACION.DE_PEAJE_DURACION,"
            + " T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO,"
            + " T_SGM_M_ADDENDA_CTT_DIST.FH_FIN,"
            + " T_SGM_M_ADDENDA_CTT_DIST.NM_CDC as Qd,"
            + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA,"
            + " T_SGM_M_GESTORES.DE_GESTOR"
            + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
            + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
            + " LEFT JOIN T_SGM_M_CLIENTES ON"
            + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
            + " LEFT JOIN T_SGM_M_GESTORES ON"
            + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR"
            + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
            + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
            + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
            + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES"
            + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON"
            + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
            + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
            + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
            + " T_SGM_G_PS.CD_PAIS = T_SGM_P_PROVINCIAS.CD_PAIS"
            + " LEFT JOIN T_SGM_P_MUNICIPIOS ON"
            + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
            + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS"
            + " INNER JOIN T_SGM_M_ADDENDA_CTT_DIST ON"
            + " T_SGM_M_ADDENDA_CTT_DIST.ID_PS = T_SGM_G_PS.ID_PS"
            + " INNER JOIN T_SGM_P_PEAJE_DURACION ON"
            + " T_SGM_P_PEAJE_DURACION.ID_PEAJE_DURACION = T_SGM_M_ADDENDA_CTT_DIST.ID_PEAJE_DURACION"
            + " WHERE T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO in (3, 6, 7, 8, 9, 10, 15, 16)"
            + " AND T_SGM_G_PS.CD_PAIS = 'ESP' AND T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA <> 'Cisternas'"
            + " AND T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
            + " AND(T_SGM_G_CONTRATOS_PS.FH_FIN_REAL >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR T_SGM_G_CONTRATOS_PS.FH_FIN_REAL IS NULL)"
            + " AND T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
            + " AND(T_SGM_M_ADDENDA_CTT_DIST.FH_FIN >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR T_SGM_M_ADDENDA_CTT_DIST.FH_FIN is null)"
            + " order by T_SGM_G_PS.CD_CUPS";
            db = new SQLServer();
            command = new SqlCommand(strSql, db.con);
            da = new SqlDataAdapter(command);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.Informe_contratos_gas_vencimiento c = new EndesaEntity.Informe_contratos_gas_vencimiento();

                if (r["CD_CIF"] != System.DBNull.Value)
                {
                    c.nif = r["CD_CIF"].ToString();
                    cartera.GetCartera(c.nif);
                    c.responsable_territorial = cartera.responsable_territorial;
                    c.responsable_territorial = c.responsable_territorial.ToUpper();
                    c.gestor = cartera.gestor;
                    c.gestor = c.gestor.ToUpper();
                }

                if (r["DE_NOMBRE_CLIENTE"] != System.DBNull.Value)
                    c.cliente = r["DE_NOMBRE_CLIENTE"].ToString();

                if (r["CD_CUPS"] != System.DBNull.Value)
                    c.cups20 = r["CD_CUPS"].ToString();

                if (r["FH_INICIO"] != System.DBNull.Value)
                    c.fecha_inicio = Convert.ToDateTime(r["FH_INICIO"]);

                if (r["FH_FIN"] != System.DBNull.Value)
                    c.fecha_fin = Convert.ToDateTime(r["FH_FIN"]);
                else
                {
                    if (r["DE_PEAJE_DURACION"].ToString().ToUpper() == "MENSUAL")
                    {
                        fecha_calculo = Convert.ToDateTime(r["FH_INICIO"]);
                        c.fecha_fin = fecha_calculo.AddDays(DateTime.DaysInMonth(fecha_calculo.Year, fecha_calculo.Month) - 1);
                    }
                    
                }


                if (r["qd"] != System.DBNull.Value)
                    c.qd = Convert.ToDouble(r["qd"]);                
                               
                if (r["DE_TIPO_TARIFA"] != System.DBNull.Value)
                    c.tarifa = r["DE_TIPO_TARIFA"].ToString();

                if (r["DE_PEAJE_DURACION"] != System.DBNull.Value)
                    c.tipo = r["DE_PEAJE_DURACION"].ToString().ToUpper();

                if (r["DE_DISTRIBUIDORES"] != System.DBNull.Value)
                    c.distribuidora = r["DE_DISTRIBUIDORES"].ToString();



                if (FechaMaximaProductoContrato(dic_fechas_maximas, c.cups20) > DateTime.Now.AddDays(dias))
                {
                    c.continua = "Sí";
                }
                else
                    c.continua = "No";

                if(c.fecha_fin != DateTime.MinValue)
                {
                    TimeSpan diferencia = c.fecha_fin - DateTime.Now.AddDays(dias);
                    if (diferencia.Days <= dias)
                        l.Add(c);
                }

                

            }
            db.CloseConnection();
            
            return l;
        }

        public Int32 NextLine(string cups20)
        {
            List<EndesaEntity.Table_atrgas_contratos_detalle> lista;
            if (dic.TryGetValue(cups20, out lista))
                return (lista.Count() + 1);
            else
                return 1;
        }


        private List<string> Lista_NIF()
        {
            string strSql;
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;

            Dictionary<string, string> dic = new Dictionary<string, string>();

            strSql = "SELECT DISTINCT"
            + " T_SGM_M_CLIENTES.CD_CIF,"
            + " T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE,"
            + " T_SGM_P_DISTRIBUIDORES.DE_DISTRIBUIDORES,"
            + " T_SGM_G_PS.CD_CUPS,"
            + " T_SGM_P_PEAJE_DURACION.DE_PEAJE_DURACION,"
            + " T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO,"
            + " T_SGM_M_ADDENDA_CTT_DIST.FH_FIN,"
            + " T_SGM_M_ADDENDA_CTT_DIST.NM_CDC as Qd,"
            + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA,"
            + " T_SGM_M_GESTORES.DE_GESTOR"
            + " FROM T_SGM_G_PS LEFT JOIN T_SGM_G_CONTRATOS_PS ON"
            + " T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
            + " LEFT JOIN T_SGM_M_CLIENTES ON"
            + " T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
            + " LEFT JOIN T_SGM_M_GESTORES ON"
            + " T_SGM_M_CLIENTES.ID_GESTOR = T_SGM_M_GESTORES.ID_GESTOR"
            + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
            + " T_SGM_G_PS.ID_TIPO_TARIFA = T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA"
            + " LEFT JOIN T_SGM_P_DISTRIBUIDORES ON"
            + " T_SGM_G_PS.ID_DISTRIBUIDORES = T_SGM_P_DISTRIBUIDORES.ID_DISTRIBUIDORES"
            + " LEFT JOIN T_SGM_M_RED_DISTRIBUCION ON"
            + " T_SGM_G_PS.ID_RED_DISTRIBUCION = T_SGM_M_RED_DISTRIBUCION.ID_RED_DISTRIBUCION"
            + " LEFT JOIN T_SGM_P_PROVINCIAS ON"
            + " T_SGM_G_PS.ID_PROVINCIA = T_SGM_P_PROVINCIAS.ID_PROVINCIA AND"
            + " T_SGM_G_PS.CD_PAIS = T_SGM_P_PROVINCIAS.CD_PAIS"
            + " LEFT JOIN T_SGM_P_MUNICIPIOS ON"
            + " T_SGM_G_PS.CD_MUNICIPIO = T_SGM_P_MUNICIPIOS.CD_MUNICIPIO AND"
            + " T_SGM_G_PS.CD_PAIS = T_SGM_P_MUNICIPIOS.CD_PAIS"
            + " INNER JOIN T_SGM_M_ADDENDA_CTT_DIST ON"
            + " T_SGM_M_ADDENDA_CTT_DIST.ID_PS = T_SGM_G_PS.ID_PS"
            + " INNER JOIN T_SGM_P_PEAJE_DURACION ON"
            + " T_SGM_P_PEAJE_DURACION.ID_PEAJE_DURACION = T_SGM_M_ADDENDA_CTT_DIST.ID_PEAJE_DURACION"
            + " WHERE T_SGM_G_CONTRATOS_PS.ID_ESTADO_CTO in (3, 6, 7, 8, 9, 10, 15, 16)"
            + " AND T_SGM_G_PS.CD_PAIS = 'ESP' AND T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA <> 'Cisternas'"
            + " AND T_SGM_G_CONTRATOS_PS.FH_INICIO_REAL <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
            + " AND(T_SGM_G_CONTRATOS_PS.FH_FIN_REAL >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR T_SGM_G_CONTRATOS_PS.FH_FIN_REAL IS NULL)"
            + " AND T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
            + " AND(T_SGM_M_ADDENDA_CTT_DIST.FH_FIN >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR T_SGM_M_ADDENDA_CTT_DIST.FH_FIN is null)"
            + " order by T_SGM_G_PS.CD_CUPS";
            db = new SQLServer();
            command = new SqlCommand(strSql, db.con);
            SqlDataAdapter da = new SqlDataAdapter(command);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["CD_CIF"] != System.DBNull.Value)
                {
                    string o;
                    if (!dic.TryGetValue(r["CD_CIF"].ToString(), out o))
                        dic.Add(r["CD_CIF"].ToString(), r["CD_CIF"].ToString());
                }
                
            }
            db.CloseConnection();

            return dic.Keys.ToList();
        }

        private DateTime FechaMaximaProductoContrato(Dictionary<string, DateTime> dic, string cups)
        {            
            DateTime o;
            if (dic.TryGetValue(cups, out o))
                return o;
            else
                return DateTime.MinValue;
          
        }

        public bool Existe_ContratoDetalle_ConFinMundo(string nif, string cups, string tipo, DateTime fecha_inicio)
        {
            DateTime fecha_fin = new DateTime(4999, 12, 31);

            List<EndesaEntity.Table_atrgas_contratos_detalle> o;

            if (dic.TryGetValue(nif, out o))
            {
                for (int i = 0; i < o.Count; i++)
                {
                    if (o[i].cups20 == cups
                        && o[i].tipo == tipo
                        && o[i].fecha_inicio.Date == fecha_inicio.Date
                        && o[i].fecha_fin.Date == fecha_fin.Date)
                        return true;
                }
            }
            return false;
        }

    }
}
