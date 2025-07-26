using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace EndesaBusiness.facturacion
{
    

    public class Inffact
    {
        System.Data.DataTable dt;
        utilidades.Global g = new utilidades.Global();
        public List<EndesaEntity.facturacion.FacturaInffact> lf { get; set; }

        public Inffact()
        {
            lf = new List<EndesaEntity.facturacion.FacturaInffact>();
        }


        public System.Data.DataTable ConceptosFacturacion(List<string> lista_empresas, string txtNum, string txtCod, string txtDescripcion)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            bool firstOnly = true;
            bool firstOnlyEmpresa = true;

            try
            {
                
                strSql = "SELECT e.descripcion AS empresa, cf.conc_sce as Num, cf.descripcion_corta, cf.descripcion"
                    + " from fo_cf cf INNER JOIN fo_empresas e ON"
                    + " e.empresa_id = cf.empresa_id";

                if (txtNum != "" || txtCod != "" || txtDescripcion != "")
                {
                    strSql = strSql + " where";

                    if (txtNum != "")
                    {
                        strSql = strSql + " cf.conc_sce like '%" + txtNum + "%'";
                        firstOnly = false;
                    }

                    if (txtCod != "")
                    {
                        if (!firstOnly)
                        {
                            strSql = strSql + " and";
                        }

                        strSql = strSql + " cf.descripcion_corta like '%" + txtCod + "%'";
                        firstOnly = false;
                    }

                    if (txtDescripcion != "")
                    {
                        if (!firstOnly)
                        {
                            strSql = strSql + " and";
                        }
                        strSql = strSql + " cf.descripcion like '%" + txtDescripcion + "%'";
                        firstOnly = false;
                    }

                    if (lista_empresas.Count() > 0)
                    {
                        if (!firstOnly)
                        {
                            strSql = strSql + " and";
                        }
                        foreach(string p in lista_empresas)
                        {
                            if (firstOnlyEmpresa)
                            {
                                strSql = strSql + " e.descripcion in ('" + p + "'";
                                firstOnlyEmpresa = false;
                            }
                            else
                                strSql = strSql + ",'" + p + "'";
                        }
                        strSql = strSql + ")";

                        firstOnly = false;
                    }

                }
                strSql = strSql + " group by empresa, cf.CONC_SCE order by cf.CONC_SCE;";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                MySqlDataAdapter da = new MySqlDataAdapter(command);                
                dt = new DataTable();
                da.Fill(dt);
                return dt;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message,
                "Error en la búsqueda de conceptos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }
        }

        public System.Data.DataTable CargadgvFactura(string tabla, DateTime fd, DateTime fh, Boolean usarFechaFactura,
                string cupsree, string cnifdnic, string conceptos,
                string tiponegocio, string empresas, string tiposfactura, string cfactura)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;            

            string strSql = null;
            bool firstOnly = true;
            bool firstOnly2 = true;
            string[] lista_empresas;
            string[] lista_tiponegocio;
            string[] lista_cups;
            string[] lista_nifs;
            string[] lista_conceptos;
            string[] lista_tipos_factura;

            DateTime begin = new DateTime();
           

            try
            {
                #region query
                if (cfactura != null)
                {
                    fd = DateTime.MinValue;
                    fh = DateTime.MaxValue;
                    lista_conceptos = null;

                }

                if (tabla == "fact.fo")
                {
                    strSql = "select concat(fo.CEMPTITU,'-',e.descripcion)as CEMPTITU, e.descripcion as EMPRESA, "
                        + " fo.CNIFDNIC, fo.DAPERSOC, fo.CFACTURA, fo.TESTFACT,"
                        + " fo.CEMPTITU, fo.FFACTURA, fo.FFACTDES, fo.FFACTHAS,"
                        + " FORMAT(fo.VCUOVAFA, 0, 'de_DE') as VCUOVAFA,"
                        + " FORMAT(fo.IFACTURA, 2, 'de_DE') as IFACTURA,"
                        + " FORMAT(fo.IVA, 2, 'de_DE') as IVA,"
                        + " FORMAT(fo.IIMPUES2, 2, 'de_DE') as IIMPUES2,"
                        + " FORMAT(fo.IIMPUES3, 2, 'de_DE') as IIMPUES3,"
                        + " fo.CREFEREN, fo.SECFACTU, fo.CUPSREE,"
                        + " tf.descripcion as TFACTURA, fo.TIPONEGOCIO";

                    for (int i = 1; i <= 9; i++)
                    {
                        strSql = strSql + ", TCONFAC" + i + ", ICONFAC" + i;
                    }

                    for (int i = 10; i <= 20; i++)
                    {
                            strSql = strSql + ", TCONFA" + i + ", ICONFA" + i;
                        }
                    }
                    else
                    {
                        strSql = "select *";
                    }



                    strSql = strSql + " from " + tabla + " inner join fo_p_tipos_factura tf on"
                    + " tf.id_tipo_factura = fo.TFACTURA"
                    + " inner join fo_empresas e on"
                    + " e.cemptitu = fo.CEMPTITU and"
                    + " e.segmento = fo.INDEMPRE"
                    + " where ";

                    if (usarFechaFactura)
                    {
                        strSql = strSql + " (FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                        + " FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                    }
                    else
                    {
                        strSql = strSql + " (FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                    }

                    if (cfactura != null)
                    {
                        strSql = strSql + " and CFACTURA = '" + cfactura + "'";
                    }

                    if (tiponegocio != null)
                    {

                        lista_tiponegocio = tiponegocio.Split(';');
                        for (int i = 0; i < lista_tiponegocio.Count(); i++)
                        {
                            if (firstOnly)
                            {
                                strSql = strSql + " and TIPONEGOCIO in ('"
                                    + lista_tiponegocio[i] + "'";
                                firstOnly = false;
                            }
                            else
                            {
                                strSql = strSql + ", '" + lista_tiponegocio[i] + "'";
                            }
                        }

                        firstOnly = true;
                        strSql = strSql + ")";
                    }

                    if (empresas != null)
                    {
                        lista_empresas = empresas.Split(';');
                        for (int i = 0; i < lista_empresas.Count(); i++)
                        {
                            if (firstOnly)
                            {
                                strSql = strSql + " and e.descripcion in ('"
                                    + lista_empresas[i] + "'";
                                firstOnly = false;
                            }
                            else
                            {
                                strSql = strSql + ", '" + lista_empresas[i] + "'";
                            }
                        }

                        firstOnly = true;
                        strSql = strSql + ")";
                    }

                    if (tiposfactura != null)
                    {
                        lista_tipos_factura = tiposfactura.Split(';');

                        for (int i = 0; i < lista_tipos_factura.Count(); i++)
                        {
                            if (firstOnly)
                            {
                                strSql = strSql + " and tf.descripcion in ('"
                                    + lista_tipos_factura[i] + "'";
                                firstOnly = false;
                            }
                            else
                            {
                                strSql = strSql + ", '" + lista_tipos_factura[i] + "'";
                            }
                        }

                        firstOnly = true;
                        strSql = strSql + ")";
                    }


                    if (cnifdnic != null)
                    {
                        lista_nifs = cnifdnic.Split('|');
                        if (lista_nifs.Count() == 1)
                            strSql = strSql + " and cnifdnic like '" + cnifdnic + "%'";
                        else
                        {
                            for (int i = 0; i < lista_nifs.Count(); i++)
                            {
                                if (firstOnly)
                                {
                                    strSql = strSql + " and cnifdnic in ('"
                                        + lista_nifs[i] + "'";
                                    firstOnly = false;
                                }
                                else
                                {
                                    strSql = strSql + ", '" + lista_nifs[i] + "'";
                                }
                            }

                            firstOnly = true;
                            strSql = strSql + ")";
                        }
                    }

                    if (cupsree != null)
                    {
                        lista_cups = cupsree.Split('|');
                        for (int i = 0; i < lista_cups.Count(); i++)
                        {
                            if (firstOnly)
                            {
                                if (lista_cups[i].Length >= 20)
                                {
                                    strSql = strSql + " and substr(CUPSREE,1,20) in ("
                                    + "'" + lista_cups[i].Substring(0, 20) + "'";
                                }
                                else
                                {
                                    strSql = strSql + " and substr(CCOUNIPS,1,13) in ("
                                    + "'" + lista_cups[i].Substring(0, 13) + "'";
                                }

                                firstOnly = false;
                            }
                            else
                            {
                                if (lista_cups[i].Length >= 20)
                                {
                                    strSql = strSql + ", '" + lista_cups[i].Substring(0, 20) + "'";
                                }
                                else
                                {
                                    strSql = strSql + ", '" + lista_cups[i].Substring(0, 13) + "'";
                                }

                            }
                        }

                        firstOnly = true;
                        strSql = strSql + ")";

                    }

                    // ****************************************************************
                    // *********** TRATAMIENTO DE LOS CONCEPTOS ***********************
                    // ****************************************************************

                    if (conceptos != null)
                    {
                        lista_conceptos = conceptos.Split(';');
                        strSql = strSql + " and (";
                        for (int i = 1; i <= 9; i++)
                        {
                            firstOnly2 = true;
                            if (!firstOnly)
                            {
                                strSql = strSql + " or";
                            }

                            for (int j = 0; j < lista_conceptos.Count(); j++)
                            {

                                if (firstOnly2)
                                {
                                    strSql = strSql + " TCONFAC" + i + " in ("
                                    + lista_conceptos[j];
                                    firstOnly2 = false;
                                }
                                else
                                {
                                    strSql = strSql + ", " + lista_conceptos[j];
                                }

                            } // for(int j = 0;j<lista_conceptos.Count();j++)

                            strSql = strSql + ")";
                            firstOnly = false;

                        } // for(int i = 1; i<= 9; i++)

                        for (int i = 10; i <= 20; i++)
                        {
                            firstOnly2 = true;
                            if (!firstOnly)
                            {
                                strSql = strSql + " or";
                            }
                            for (int j = 0; j < lista_conceptos.Count(); j++)
                            {

                                if (firstOnly2)
                                {
                                    strSql = strSql + " TCONFA" + i + " in ("
                                    + lista_conceptos[j];
                                    firstOnly2 = false;
                                }
                                else
                                {
                                    strSql = strSql + ", " + lista_conceptos[j];
                                }

                            } // foreach(DataGridViewRow r in this.dgvConceptos.Rows)
                            strSql = strSql + ")";
                            firstOnly = false;
                        } // for (int i = 10; i <= 20; i++)
                        strSql = strSql + ")";
                    } // if (this.HayConceptosSeleccionados())



                #endregion

                begin = DateTime.Now;
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.facturacion.FacturaInffact f = new EndesaEntity.facturacion.FacturaInffact();

                    f.cemptitu = reader["CEMPTITU"].ToString();
                    f.cnifdnic = reader["CNIFDNIC"].ToString();
                    f.dapersoc = reader["DAPERSOC"].ToString();

                    if (reader["CFACTURA"] != System.DBNull.Value)
                    {
                        f.cfactura = reader["CFACTURA"].ToString();
                    }
                    if (reader["FFACTURA"] != System.DBNull.Value)
                    {
                        f.ffactura = Convert.ToDateTime(reader["FFACTURA"]);
                    }
                    if (reader["FFACTDES"] != System.DBNull.Value)
                    {
                        f.ffactdes = Convert.ToDateTime(reader["FFACTDES"]);
                        f.ffacthas = Convert.ToDateTime(reader["FFACTHAS"]);
                    }
                    if (reader["VCUOVAFA"] != System.DBNull.Value)
                    {
                        f.vcuovafa = Convert.ToDouble(reader["VCUOVAFA"]);
                    }


                    f.iva = Convert.ToDouble(reader["IVA"]);
                    f.ifactura = Convert.ToDouble(reader["IFACTURA"]);
                    f.creferen = Convert.ToInt64(reader["CREFEREN"]);
                        f.secfactu = Convert.ToInt32(reader["SECFACTU"]);
                        f.testfact = reader["TESTFACT"].ToString();
                        f.tfactura = reader["TFACTURA"].ToString();

                        if (reader["CUPSREE"] != System.DBNull.Value)
                        {
                            f.cupsree = reader["CUPSREE"].ToString();
                        }

                        for (int i = 1; i <= 9; i++)
                        {
                            if (reader["tconfac" + i.ToString()] != System.DBNull.Value)
                            {
                                f.tconfact[i - 1] = Convert.ToInt32(reader["tconfac" + i.ToString()]);
                            }
                            if (reader["iconfac" + i.ToString()] != System.DBNull.Value)
                            {
                                f.iconfact[i - 1] = Convert.ToDouble(reader["iconfac" + i.ToString()]);
                            }

                        }

                        for (int i = 10; i <= 20; i++)
                        {
                            if (reader["tconfa" + i.ToString()] != System.DBNull.Value)
                            {
                                f.tconfact[i - 1] = Convert.ToInt32(reader["tconfa" + i.ToString()]);
                            }
                            if (reader["iconfa" + i.ToString()] != System.DBNull.Value)
                            {
                                f.iconfact[i - 1] = Convert.ToDouble(reader["iconfa" + i.ToString()]);
                            }
                        }

                        lf.Add(f);


                    }

                    reader.Close();
                db.CloseConnection();
                g.SaveQuery("FrmFacturasOperaciones", strSql, begin, DateTime.Now);

                MySqlDataAdapter da = new MySqlDataAdapter(command);
                dt = new DataTable();
                da.Fill(dt);
                return dt;
                
                
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la búsqueda de facturas",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }
        }


        public DateTime UltimaFechaFactura()
        {
            return DateTime.Now.Date;
        }


    }
}
