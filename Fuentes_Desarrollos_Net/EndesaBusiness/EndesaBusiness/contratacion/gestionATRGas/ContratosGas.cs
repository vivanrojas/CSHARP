using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class ContratosGas : EndesaEntity.Table_atrgas_contratos
    {
        public Dictionary<string, List<EndesaEntity.Table_atrgas_contratos>> dic { get; set; }
        public ContratosGasDetalle cgd { get; set; }

        public ContratosGas()
        {
            dic = new Dictionary<string, List<EndesaEntity.Table_atrgas_contratos>>();
            cgd = new ContratosGasDetalle();
            this.GetAll();
        }

        public void Save()
        {
            try
            {
                List< EndesaEntity.Table_atrgas_contratos > o;
                if (!dic.TryGetValue(this.cnifdnic, out o))
                    this.New();
                else if (dic.TryGetValue(this.cnifdnic, out o))
                {
                    if (!o.Exists(z => z.cups20 == this.cups20))
                        this.New();
                    else
                        this.Update();
                }
                else
                    this.Update();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "ContratosGas - Save",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private void New()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {

                strSql = "insert into atrgas_contratos (cnifdnic, dapersoc, distribuidora,"
                    + " cups20, comentarios_descuadres, comentarios_contratacion, tramitacion, created_by, creation_date) values ";

                if (this.cnifdnic != null)
                    strSql += "('" + this.cnifdnic + "'";
                else
                    strSql += "(null";

                if (this.dapersoc != null)
                    strSql += ",'" + this.dapersoc + "'";
                else
                    strSql += ",null";

                if (this.distribuidora != null)
                    strSql += ",'" + this.distribuidora + "'";
                else
                    strSql += ",null";

                if (this.cups20 != null)
                    strSql += ",'" + this.cups20 + "'";
                else
                    strSql += ",null";

                if (this.comentarios_descuadres != null)
                    strSql += ",'" + this.comentarios_descuadres + "'";
                else
                    strSql += ",null";

                if (this.comentarios_contratacion != null)
                    strSql += ",'" + this.comentarios_contratacion + "'";
                else
                    strSql += ",null";

                if (this.tramitacion != null)
                    strSql += ",'" + this.tramitacion + "'";
                else
                    strSql += ",null";

                // strSql += ",'Distribuidora'";


                if (this.created_by != null)
                    strSql += ",'" + this.created_by + "'";
                else
                    strSql += ",'" + System.Environment.UserName + "'";

                strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                MessageBox.Show("Se ha añadido correctamente el contrato para el CUPS: " + this.cups20,
                 "Contratos Gas",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "ContratosGas - New",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

        private void Update()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update atrgas_contratos set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            if (this.cnifdnic != null)
                strSql += " ,cnifdnic = '" + this.cnifdnic + "'";

            if (this.dapersoc != null)
                strSql += " ,dapersoc = '" + this.dapersoc + "'";

            if (this.distribuidora != null)
                strSql += " ,distribuidora = '" + this.distribuidora + "'";

            if (this.comentarios_descuadres != null)
                strSql += " ,comentarios_descuadres = '" + this.comentarios_descuadres + "'";

            if (this.comentarios_contratacion != null)
                strSql += " ,comentarios_contratacion = '" + this.comentarios_contratacion + "'";

            if (this.tramitacion != null)
                strSql += " ,tramitacion = '" + this.tramitacion + "'";

            strSql += " where cnifdnic = '" + this.cnifdnic + "'"
                + " and cups20 = '" + this.cups20 + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el contrato para el CUPS: " + this.cups20,
            "Contratos Gas",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);

        }

        public void GetAll()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select cnifdnic, dapersoc, distribuidora, cups20, comentarios_descuadres, comentarios_contratacion,"
                    + "tramitacion,"
                    + "created_by, creation_date, last_update_by, last_update_date from atrgas_contratos order by last_update_date desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Table_atrgas_contratos c = new EndesaEntity.Table_atrgas_contratos();

                    if (r["cnifdnic"] != System.DBNull.Value)
                        c.cnifdnic = r["cnifdnic"].ToString().ToUpper();

                    if (r["dapersoc"] != System.DBNull.Value)
                        c.dapersoc = r["dapersoc"].ToString().ToUpper();

                    if (r["distribuidora"] != System.DBNull.Value)
                        c.distribuidora = r["distribuidora"].ToString().ToUpper();

                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString().ToUpper();

                    if (r["comentarios_descuadres"] != System.DBNull.Value)
                        c.comentarios_descuadres = r["comentarios_descuadres"].ToString();

                    if (r["comentarios_contratacion"] != System.DBNull.Value)
                        c.comentarios_contratacion = r["comentarios_contratacion"].ToString();

                    if (r["tramitacion"] != System.DBNull.Value)
                        c.tramitacion = r["tramitacion"].ToString();

                    if (r["created_by"] != System.DBNull.Value)
                        c.created_by = r["created_by"].ToString();

                    List<EndesaEntity.Table_atrgas_contratos> o;
                    if (!dic.TryGetValue(c.cnifdnic, out o))
                    {
                        o = new List<EndesaEntity.Table_atrgas_contratos>();
                        o.Add(c);
                        dic.Add(c.cnifdnic, o);
                    }
                    else
                        o.Add(c);

                }

                // sod = new SalesOrderDetail(dic_so[dic_so.Count - 1].sales_order_id);

                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "ContratosGas - GetAll",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }

        public void Del()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                DialogResult result = MessageBox.Show("Warning:" + System.Environment.NewLine
                    + "Todas las líneas del contrato serán borradas." + System.Environment.NewLine
                    + "¿Desea borrar el contrato?"
                    , "Borrado de contrato " + this.cups20 + " - " + this.dapersoc + " - " + this.cnifdnic,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    //Eliminamos por clave CUPS - NIF
                    strSql = "delete from atrgas_contratos_detalle where cups20 = '" + this.cups20 + "' and nif = '"+ this.cnifdnic +"'";
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    //Eliminamos por clave CUPS - NIF
                    strSql = "delete from atrgas_contratos where cups20 = '" + this.cups20 + "' and cnifdnic = '" + this.cnifdnic + "'";
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("Contrato borrado.",
                  "Contratos GAS",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "ContratosGas - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }


        }

        //public Int32 LastID()
        //{
        //    if (dic.Count > 0)
        //        return dic[0].element_id;
        //    else
        //        return 0;
        //}

        public void GetRecord(string nif, string cups20)
        {

            try
            {
                EndesaEntity.Table_atrgas_contratos r = new EndesaEntity.Table_atrgas_contratos();
                r = dic.FirstOrDefault(z => z.Key == nif).Value.Find(z => z.cups20 == cups20);
                this.cnifdnic = r.cnifdnic;
                this.dapersoc = r.dapersoc;
                this.distribuidora = r.distribuidora;
                this.cups20 = r.cups20;
                this.comentarios_descuadres = r.comentarios_descuadres;
                this.comentarios_contratacion = r.comentarios_contratacion;
                this.created_by = r.created_by;
                this.creation_date = r.creation_date;
                this.last_update_by = r.last_update_by;
                this.last_update_date = r.last_update_date;
                this.tramitacion = r.tramitacion;


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "ContratosGas - GetSalesOrder",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public bool ExisteContrato(string nif, string cups20)
        {
            
            List<EndesaEntity.Table_atrgas_contratos> o;
            if(dic.TryGetValue(nif, out o))            
                return o.Exists(z => z.cups20 == cups20);
            else
                return false;

        }
    }
}
