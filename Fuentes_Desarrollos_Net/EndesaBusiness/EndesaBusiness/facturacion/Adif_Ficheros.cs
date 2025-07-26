using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Fabric.Description;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class Adif_Ficheros
    {
        public Int64 id_fichero { get; set; }
        public String nombre_fichero { get; set; }
        public DateTime fecha_carga { get; set; }
        public String usuario { get; set; }

        Adif_CupsLote cupsLote;
        Adif_Empresa empresa;
        Adif_Producto producto;

        List<Adif_CupsLote> cups = new List<Adif_CupsLote>();
        List<Adif_Empresa> empresas = new List<Adif_Empresa>();
        List<Adif_Producto> productos = new List<Adif_Producto>();

        public Adif_Ficheros(String nombre_archivo)
        {
            String strSql;
            DateTime inicio = new DateTime();
            MySQLDB db;
            MySqlCommand command;
            FileInfo archivo = new System.IO.FileInfo(nombre_archivo);

            inicio = DateTime.Now;
            strSql = "INSERT INTO adif_ficheros SET" +
                " nombre_fichero = '" + archivo.Name + "'," +
                " fecha_carga = '" + inicio.ToString("yyyy-MM-dd HH:mm:ss") + "'," +
                " usuario = '" + Environment.UserName + "';";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            //this.CargaCups();
            //this.CargaEmpresas();
            //this.CargaProductos();

            this.CargaArchivo(nombre_archivo);

        }

        private void CargaArchivo(String nombre_archivo)
        {
            long i = 0;
            string line;
            StringBuilder sb = new StringBuilder();
            Boolean firstOnly = true;
            bool firstOnlyMes = true;
            MySQLDB db;
            MySqlCommand command;
            int mes = 0;
            DateTime fechahora = new DateTime();

            try
            {

                sb.Append("delete from adif_ficheros_facturas_tmp");              
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();


                FileInfo archivo = new FileInfo(nombre_archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(nombre_archivo);
                while ((line = file.ReadLine()) != null)
                {

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO adif_ficheros_facturas_tmp ");
                        sb.Append("(fichero,cups20,lote,comercializadora,fecha_hora,flag_invierno_verano,energia_consumida,");
                        sb.Append("modalidad,precio_aplicar_cierres,precio_horario,constante_v,constante_k,");
                        sb.Append("porcentaje_cierres,energia_facturada,precio,cnpr,cpre,created_by,created_date) values ");
                        firstOnly = false;
                    }

                    if (line.Length > 50)
                    {
                        string[] f = line.Split('\t');
                        i++;

                        if (firstOnlyMes)
                        {
                            if (!nombre_archivo.Contains("CIERRE"))
                                mes = Convert.ToInt32(f[3].Replace("/", "").Substring(0, 6));
                            else
                                mes = Convert.ToInt32(f[3].Replace("/", "").Substring(4, 4) + f[3].Replace("/", "").Substring(2, 2));
                            firstOnlyMes = false;
                        }

                        sb.Append("('").Append(archivo.Name).Append("',"); // fichero                        
                        sb.Append("'").Append(f[0]).Append("',"); // cups20
                        sb.Append(f[1]).Append(","); // lote                                                
                        sb.Append("'").Append(f[2].Trim()).Append("',"); // comercializadora

                        // fecha_hora
                        if (!nombre_archivo.Contains("CIERRE"))
                            sb.Append("'").Append(f[3].Replace("/", "-")).Append("',");
                        else
                        {
                            fechahora = Convert.ToDateTime(f[3]);
                            sb.Append("'").Append(fechahora.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        }



                        sb.Append(f[4]).Append(","); // flag_invierno_verano

                        if (!nombre_archivo.Contains("CIERRE"))
                            sb.Append(f[5]).Append(","); // energia_consumida
                        else
                            sb.Append("null,");

                        sb.Append("'").Append(f[6]).Append("',"); // modalidad
                        sb.Append("'").Append(f[7]).Append("',"); // precio_aplicar_cierres
                        sb.Append(CN(f[8])).Append(","); // precio_horario
                        sb.Append(CN(f[9])).Append(","); // constante_v
                        sb.Append(CN(f[10])).Append(","); // constante_k

                        // porcentaje_cierres
                        if (f[11].Trim() == "-")
                            sb.Append("null,");
                        else
                           sb.Append(CN(f[11])).Append(","); 

                        sb.Append(CN(f[12])).Append(","); // energia_facturada
                        sb.Append(CN(f[13])).Append(","); // precio

                        if (!nombre_archivo.Contains("CIERRE"))
                        {
                            sb.Append(CN(f[14])).Append(","); // cnpr
                            sb.Append("null,"); // cpre
                        }
                        else
                        {
                            sb.Append("null,");
                            sb.Append(CN(f[14])).Append(","); // cpre
                        }


                        sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',"); // created_by
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),"); // created_date

                    }
                                       
                }

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                file.Close();


                sb.Append("REPLACE INTO adif_ficheros_facturas_meses");
                sb.Append(" SELECT f.fichero, f.cups20, f.lote,");
                sb.Append(mes).Append(" as mes,");
                sb.Append(" SUM(f.energia_consumida) AS energia_consumida,");
                sb.Append(" SUM(f.energia_facturada) AS energia_facturada,");
                sb.Append(" SUM(f.cnpr) AS cnpr,");
                sb.Append(" SUM(f.cpre) AS cpre");
                sb.Append(" FROM adif_ficheros_facturas_tmp f ");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                sb = null;
                sb = new StringBuilder();
                sb.Append("REPLACE INTO adif_ficheros_facturas");
                sb.Append(" SELECT * from adif_ficheros_facturas_tmp");               
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Error en la importación de facturas",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }



        private string CD(String t)
        {
            String salida = "";


            salida = "'" + t.Substring(6, 4) + "-" + t.Substring(3, 2) + "-" + t.Substring(0, 2) + "'";
            return salida;
        }

        private string CN(String t)
        {
            t = t.Trim();
            if (t == "")
            {
                return "null";
            }
            else
            {
                t = t.Replace(" ", "");
                //t = t.Replace("+", string.Empty);
                //t = t.Replace("-", string.Empty);
                t = t.Replace(".", string.Empty);
                t = t.Replace(",", ".");

                if (t == "")
                {
                    t = "null";
                }
            }

            return t;
        }


        private Int32 GetCupsID(String cupsree, int lote)
        {
            Int32 i = 0;
            Int32 id = 0;
            foreach (Adif_CupsLote c in cups)
            {
                i++;
                if (c.cupsree == cupsree && c.lote == lote)
                {
                    id = c.id_cups_lote;
                }
            }

            if (id == 0)
            {
                i++;
                cupsLote = new Adif_CupsLote();
                cupsLote.id_cups_lote = i;
                cupsLote.cupsree = cupsree;
                cupsLote.lote = lote;
                cupsLote.Save(i);
                this.cups.Add(cupsLote);
                id = i;
            }

            return id;
        }


        private Int32 GetEmpresaID(String empresa)
        {
            Int32 i = 0;
            Int32 id = 0;
            foreach (Adif_Empresa c in empresas)
            {
                i++;
                if (c.empresa == empresa)
                {
                    id = c.id_empresa;
                }
            }

            if (id == 0)
            {
                i++;
                this.empresa = new Adif_Empresa();
                this.empresa.id_empresa = i;
                this.empresa.empresa = empresa;
                this.empresa.Save(i);
                this.empresas.Add(this.empresa);
                id = i;
            }

            return id;
        }

        private Int32 GetProductoID(String producto)
        {
            Int32 i = 0;
            Int32 id = 0;
            foreach (Adif_Producto c in productos)
            {
                i++;
                if (c.producto == producto)
                {
                    id = c.id_producto;
                }
            }

            if (id == 0)
            {
                i++;
                this.producto = new Adif_Producto();
                this.producto.id_producto = i;
                this.producto.producto = producto;
                this.producto.Save(i);
                this.productos.Add(this.producto);
                id = i;
            }
            return id;

        }


        private void CargaCups()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select * from adif_cups_lotes;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    Adif_CupsLote c = new Adif_CupsLote();
                    c.id_cups_lote = Convert.ToInt32(reader["id_cups_lote"]);
                    c.cupsree = reader["cupsree"].ToString();
                    c.lote = Convert.ToInt32(reader["lote"]);
                    this.cups.Add(c);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "Error en la importación de CUPS",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }

        }


        private void CargaEmpresas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select * from adif_empresas;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    Adif_Empresa c =
                        new Adif_Empresa();
                    c.id_empresa = Convert.ToInt32(reader["id_empresa"]);
                    c.empresa = reader["empresa"].ToString();
                    this.empresas.Add(c);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "Error en la importación de empresas",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }

        }

        private void CargaProductos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select * from adif_productos;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    Adif_Producto c = new Adif_Producto();
                    c.id_producto = Convert.ToInt32(reader["id_producto"]);
                    c.producto = reader["producto"].ToString();
                    this.productos.Add(c);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "Error en la importación de productos.",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }

        }
    }
}
