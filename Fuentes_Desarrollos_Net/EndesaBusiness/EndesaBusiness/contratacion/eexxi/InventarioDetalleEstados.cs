using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EndesaEntity.contratacion;

namespace EndesaBusiness.contratacion.eexxi
{
    class InventarioDetalleEstados : Inventario_Detalle_Estados_Tabla
    {
        public Dictionary<string, List<Inventario_Detalle_Estados_Tabla>> dic { get; set; }
        public Dictionary<string, List<Inventario_Detalle_Estados_Tabla>> dic_tmp { get; set; }

        public InventarioDetalleEstados()
        {
            dic = new Dictionary<string, List<Inventario_Detalle_Estados_Tabla>>();
            dic_tmp = new Dictionary<string, List<Inventario_Detalle_Estados_Tabla>>();
            Carga();
        }

        public void AnalizaSolicitud(string cups22, EndesaEntity.contratacion.xxi.XML_Datos xml)
        {
            Inventario_Detalle_Estados_Tabla c = new Inventario_Detalle_Estados_Tabla();

            this.cups22 = xml.cups;
            this.codigoproceso = xml.codigoDelProceso;
            this.codigopaso = xml.codigoDePaso;
            this.codigosolicitud = xml.codigoDeSolicitud;
            this.fechaactivacion = xml.fechaActivacion;

            List<Inventario_Detalle_Estados_Tabla> o;
            if (dic.TryGetValue(cups22, out o))
            {
                c = o.Find(z => z.codigosolicitud == xml.codigoDeSolicitud);
                if (c != null)
                {
                    this.linea = c.linea;
                    this.estado = c.estado;
                    this.subestado = c.subestado;
                }
                else
                {
                    c = new Inventario_Detalle_Estados_Tabla();

                    c.linea = o.Count + 1;
                    c.nif = xml.identificador;
                    c.razon_social = xml.razonSocial;
                    c.codigoproceso = xml.codigoDelProceso;
                    c.codigopaso = xml.codigoDePaso;
                    c.codigosolicitud = xml.codigoDeSolicitud;
                    c.fechaactivacion = xml.fechaActivacion;
                    c.nif = xml.identificador;
                    c.razon_social = xml.razonSocial;
                    c.nif = xml.identificador;
                    c.razon_social = xml.razonSocial;

                    o.Add(c);

                }

            }
            else
            {
                if (dic_tmp.TryGetValue(cups22, out o))
                {
                    c = new Inventario_Detalle_Estados_Tabla();

                    c.linea = 1;
                    c.codigoproceso = xml.codigoDelProceso;
                    c.codigopaso = xml.codigoDePaso;
                    c.codigosolicitud = xml.codigoDeSolicitud;
                    c.fechaactivacion = xml.fechaActivacion;
                    c.nif = xml.identificador;
                    c.razon_social = xml.razonSocial;
                    o.Add(c);
                }
                else
                {
                    c = new Inventario_Detalle_Estados_Tabla();
                    c.cups22 = cups22;
                    c.linea = 1;
                    c.codigoproceso = xml.codigoDelProceso;
                    c.codigopaso = xml.codigoDePaso;
                    c.codigosolicitud = xml.codigoDeSolicitud;
                    c.fechaactivacion = xml.fechaActivacion;
                    c.nif = xml.identificador;
                    c.razon_social = xml.razonSocial;

                    List<Inventario_Detalle_Estados_Tabla> lista = new List<Inventario_Detalle_Estados_Tabla>();
                    lista.Add(c);
                    dic_tmp.Add(cups22, lista);
                }

            }

        }



        private void Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {

                strSql = "select id_inventario, linea, CodigoDelProceso, CodigoDePaso, CodigoDeSolicitud,"
                    + " FechaActivacion, Estado, SubEstado from eexxi_inventario_detalle_estados;";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Inventario_Detalle_Estados_Tabla c = new Inventario_Detalle_Estados_Tabla();
                    //c.id_inventario = Convert.ToInt32(r["id_inventario"]);
                    c.linea = Convert.ToInt32(r["linea"]);
                    c.codigoproceso = r["CodigoDelProceso"].ToString();
                    c.codigopaso = r["CodigoDePaso"].ToString();
                    c.codigosolicitud = r["CodigoDeSolicitud"].ToString();
                    c.fechaactivacion = Convert.ToDateTime(r["FechaActivacion"]);
                    if (r["Estado"] != System.DBNull.Value)
                        c.estado = Convert.ToInt32(r["Estado"]);
                    if (r["SubEstado"] != System.DBNull.Value)
                        c.subestado = Convert.ToInt32(r["SubEstado"]);

                    List<Inventario_Detalle_Estados_Tabla> o;
                    //if (!dic.TryGetValue(c.id_inventario, out o))
                    //{
                    //    List<Inventario_Detalle_Estados_Tabla> lista = new List<Inventario_Detalle_Estados_Tabla>();
                    //    lista.Add(c);
                    //    dic.Add(c.id_inventario, lista);
                    //}
                    //else
                    //    o.Add(c);

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "InventarioDetalleEstados - Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void insert()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "insert into eexxi_inventario_detalle_estados_tmp set"
                    // + " id_inventario = " + this.id_inventario + ","
                    + " linea = " + this.linea + ","
                    + " CodigoDelProceso = '" + this.codigoproceso + "',"
                    + " CodigoDePaso = '" + this.codigopaso + "',"
                    + " CodigoDeSolicitud = '" + this.codigosolicitud + "',"
                    + " FechaActivacion = '" + this.fechaactivacion.ToString("yyyy-MM-dd HH:mm:ss") + "'";

                if (this.estado != 0)
                    strSql += " ,Estado = " + this.estado;
                if (this.estado != 0)
                    strSql += " ,SubEstado = " + this.subestado;

                strSql += " ,created_by = '" + Environment.UserName + "'"
                    + " ,created_date = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "EEXXI - InsertaInventario",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void VuelcaTemp()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "replace into eexxi_inventario_detalle_estados select * from eexxi_inventario_detalle_estados_tmp;";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "EEXXI - VuelcaTemp",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }
    }
}
