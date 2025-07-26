using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class InventarioFacturacionPortugal: EndesaEntity.facturacion.InventarioFacturacion
    {
        public Dictionary<string, EndesaEntity.facturacion.InventarioFacturacion> dic_inventario { get; set; }        
        public InventarioFacturacionPortugal()
        {
            dic_inventario = new Dictionary<string, EndesaEntity.facturacion.InventarioFacturacion>();
        }

        public InventarioFacturacionPortugal(string nif, DateTime fecha_desde, DateTime fecha_hasta)
        {
            dic_inventario = CargaInventario(nif, fecha_desde, fecha_hasta);
            EndesaBusiness.utilidades.PdteWeb pdte = new EndesaBusiness.utilidades.PdteWeb(dic_inventario.Select(z => z.Value.cups13).ToList());
            EstadoPdte(pdte);
        }
        
        private Dictionary<string, EndesaEntity.facturacion.InventarioFacturacion> CargaInventario(string nif, DateTime fd, DateTime fh)
        {
            Dictionary<string, EndesaEntity.facturacion.InventarioFacturacion> d =
                new Dictionary<string, EndesaEntity.facturacion.InventarioFacturacion>();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;


            try
            {
                strSql = "SELECT error, estado, actualizado, nif, cliente, carpeta_cliente, cups13, cpe, fd, fh, ruta_plantilla, f_ult_mod"
                    + " FROM fact.ag_pt_inventario where ";
                if (nif != null)
                    strSql += "nif = '" + nif + "' and";

                strSql += " (fd <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fh >= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " order by nif";


                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.InventarioFacturacion c = new EndesaEntity.facturacion.InventarioFacturacion();
                    c.actualizado = r["actualizado"].ToString() == "S";

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["cliente"] != System.DBNull.Value)
                        c.cliente = r["cliente"].ToString();

                    if (r["carpeta_cliente"] != System.DBNull.Value)
                        c.carpeta_cliente = r["carpeta_cliente"].ToString();

                    c.cups13 = r["cups13"].ToString();
                    c.cpe = r["cpe"].ToString();
                    c.fecha_desde = Convert.ToDateTime(r["fd"]);
                    c.fecha_hasta = Convert.ToDateTime(r["fh"]);
                    if (r["estado"] != System.DBNull.Value)
                        c.estado = r["estado"].ToString();

                    if (r["ruta_plantilla"] != System.DBNull.Value)
                        c.ruta_plantilla = r["ruta_plantilla"].ToString();

                    this.cpe = c.cpe;
                    this.fecha_desde = c.fecha_desde;
                    this.fecha_hasta = c.fecha_hasta;

                    //Comprobaciones de calidad del inventario
                    if (!c.ruta_plantilla.Contains(".xlsx"))
                    {
                        this.error = "No se ha especificado plantilla.";
                        this.Update();
                    }
                    else if (c.carpeta_cliente == null)
                    {
                        this.error = "Debe especificar el campo carpeta_cliente";
                        this.Update();
                    }
                    else
                    {
                        this.error = "";
                        this.Update();
                    }

                    c.error = this.error;
                    d.Add(c.cpe, c);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Facturadores Portugal",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return null;
            }



        }

        private void EstadoPdte(EndesaBusiness.utilidades.PdteWeb pdte)
        {
            foreach (KeyValuePair<string, EndesaEntity.facturacion.InventarioFacturacion> p in dic_inventario)
            {
                pdte.GetEstado(p.Value.cups13);
                p.Value.ltp = pdte.estado_ltp;
            }
        }

        public void Update()
        {


            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update ag_pt_inventario set";

            if (this.error != null)
                strSql += " error = '" + this.error + "'";

            if (this.actualizado)
                strSql += " ,actualizado = 'S'";

            if (this.nif != null)
                strSql += " ,nif = '" + this.nif + "'";

            if (this.cliente != null)
                strSql += " ,cliente = '" + this.cliente + "'";

            if (this.estado != null)
                strSql += " ,estado = '" + this.estado + "'";

            if (this.carpeta_cliente != null)
                strSql += " ,carpeta_cliente = '" + this.carpeta_cliente + "'";

            if (this.cpe != null)
                strSql += " ,cpe = '" + this.cpe + "'";

            if (this.ruta_plantilla != null)
                strSql += " ,ruta_plantilla = '" + this.ruta_plantilla + "'";

            strSql += " where cpe = '" + this.cpe + "'";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public string GetCUPS13FromCUPS20(string cups20)
        {
            string cups13 = "";
            EndesaEntity.facturacion.InventarioFacturacion o;
            if (dic_inventario.TryGetValue(cups20, out o))
                cups13 = o.cups13;

            return cups13;
        }       

        public bool Existe(string cups20)
        {
            EndesaEntity.facturacion.InventarioFacturacion o;
            return dic_inventario.TryGetValue(cups20, out o);


        }
    }
}
