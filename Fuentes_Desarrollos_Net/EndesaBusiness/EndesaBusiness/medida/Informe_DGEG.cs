using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class Informe_DGEG
    {

        public List<EndesaEntity.facturacion.ErseBTN> fact_list { get; set; }

        public Informe_DGEG()
        {
            fact_list = new List<EndesaEntity.facturacion.ErseBTN>();
        }

        public void Informe(string fichero, DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "delete from facturas_dgeg";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            
            CargaDatosMT(false, fd, fh);
            CargaDatosBTN_BTE(false, fd, fh);
            //CargaDatosBTE(false, fd, fh);
            //CargaDatosBTN(false, fd, fh);

            strSql = "DELETE from facturas_dgeg_agrupado";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            strSql = "REPLACE INTO facturas_dgeg_agrupado"
                 + " SELECT f.Empresa,"
                 + " f.CNIFDNIC, f.DAPERSOC, f.CUPSREE,"                 
                 + " sum(f.VCUOVAFA) AS VCUOVAFA,"
                 + " SUM(f.PUNTA) AS PUNTA, SUM(f.LLANO) AS LLANO, SUM(f.VALLE) AS VALLE, SUM(f.SUPERVALLE) AS SUPERVALLE,"
                 + " SUM(f.A_P1) AS A_P1, SUM(f.A_P2) AS A_P2, SUM(f.A_P3) AS A_P3,"
                 + " SUM(f.A_P4) AS A_P4, SUM(f.A_P5) AS A_P5, SUM(f.A_P6) AS A_P6"
                 + " from facturas_dgeg f"
                 + " GROUP BY f.Empresa, f.CUPSREE";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE from facturas_dgeg_minf";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "REPLACE INTO facturas_dgeg_minf"
                 + " SELECT f.CUPSREE, MIN(f.FFACTDES)"
                 + " FROM fo f INNER JOIN fo_empresas e ON"
                 + " e.empresa_id = f.ID_ENTORNO"
                 + " WHERE (f.FFACTDES <= '" + fh.ToString("yyyy-MM-dd") + "' and f.FFACTHAS >= '" + fd.ToString("yyyy-MM-dd") + "')"
                 + " and e.descripcion IN('BTN-Portugal','BTE-Portugal','MT-Portugal')"                 
                 + " AND f.TIPONEGOCIO = 'L' AND f.CUPSREE <> ''"
                 + " GROUP BY f.CUPSREE";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM facturas_dgeg_maxf";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "REPLACE into facturas_dgeg_maxf"
                 + " SELECT f.CUPSREE, MAX(f.FFACTHAS)"
                 + " FROM fo f INNER JOIN fo_empresas e ON"
                 + " e.empresa_id = f.ID_ENTORNO"
                 + " WHERE (f.FFACTDES <= '" + fh.ToString("yyyy-MM-dd") + "' and f.FFACTHAS >= '" + fd.ToString("yyyy-MM-dd") + "')"
                 + " and e.descripcion IN('BTN-Portugal','BTE-Portugal','MT-Portugal')"
                 + " AND f.TIPONEGOCIO = 'L' AND f.CUPSREE <> ''"
                 + " GROUP BY f.CUPSREE";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM facturas_dgeg_informe";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "REPLACE INTO facturas_dgeg_informe"
                 + " SELECT f.*, fmin.fecha AS Min_Fecha, fmax.fecha AS Max_Fecha"
                 + " FROM fact.facturas_dgeg_agrupado f"
                 + " LEFT OUTER JOIN fact.facturas_dgeg_minf fmin ON"
                 + " fmin.CUPSREE = f.CUPSREE"
                 + " LEFT OUTER JOIN fact.facturas_dgeg_maxf fmax ON"
                 + " fmax.CUPSREE = f.CUPSREE";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public void CargaDatosMT(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            bool hayDatos = false;

            string clave = "";
            string consumo_temp = "";

            try
            {

                #region
                strSql = "select f.CEMPTITU, f.CCOUNIPS, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                + " f.VCUOFIFA as POTENCIA, f.VCUOVAFA as CONSUMO, f.IISE as ISE, f.IFACTURA,"
                 + " f.TFACTURA, f.TESTFACT, f.VCONSACP as PUNTA, f.VCONSACL as LLANO, f.VCONSACV as VALLE,"
                + " f.VCONATH1 as CONSUMO_ACTIVA1,"
                + " f.VCONATH2 as CONSUMO_ACTIVA2,"
                + " f.VCONATH3 as CONSUMO_ACTIVA3,"
                + " f.VCONATH4 as CONSUMO_ACTIVA4,"
                + " f.CUPSREE, f.CNIFDNIC"
                + " from fo_empresas e"
                + " inner join ff_f_temp f on"
                + " f.CEMPTITU = e.cemptitu and"
                + " f.INDEMPRE = e.segmento"
                + " where e.descripcion = 'MT-Portugal' and"
                + " f.TFACTURA in (1,2,3)";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    // CAMBIO GUS/MARTA
                    // strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    // + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                    strSql = strSql + " and (f.FFACTDES <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTHAS >= '" + fd.ToString("yyyy-MM-dd") + "')";
                }


                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    hayDatos = true;
                    EndesaEntity.facturacion.ErseBTN e = new EndesaEntity.facturacion.ErseBTN();
                    EndesaEntity.facturacion.ClaveFactura c = new EndesaEntity.facturacion.ClaveFactura();

                    clave = reader["CEMPTITU"].ToString();
                    clave += "_" + reader["CREFEREN"].ToString();
                    clave += "_" + reader["SECFACTU"].ToString();
                    clave += "_" + reader["TESTFACT"].ToString();

                    e.creferen = reader["CREFEREN"].ToString();
                    e.secfactu = reader["SECFACTU"].ToString();
                    e.ccounips = reader["CCOUNIPS"].ToString();
                    e.cfactura = reader["CFACTURA"] != System.DBNull.Value ? reader["CFACTURA"].ToString() : null;
                    e.ffactura = Convert.ToDateTime(reader["FFACTURA"]);
                    e.ffactdes = Convert.ToDateTime(reader["FFACTDES"]);
                    e.ffacthas = Convert.ToDateTime(reader["FFACTHAS"]);
                    e.potencia = Convert.ToDouble(reader["POTENCIA"]);
                    e.consumo = Convert.ToInt32(reader["CONSUMO"]);
                    e.ise = Convert.ToDouble(reader["ISE"]);

                    e.ifactura = Convert.ToDouble(reader["IFACTURA"]);
                    e.tfactura = Convert.ToInt32(reader["TFACTURA"]);
                    e.testfact = reader["TESTFACT"].ToString();

                    if (reader["PUNTA"] != System.DBNull.Value)
                    {
                        e.consumo_punta = Convert.ToInt32(reader["PUNTA"]);
                        if (e.consumo < 0)
                            e.consumo_punta = e.consumo_punta * -1;
                    }

                    if (reader["LLANO"] != System.DBNull.Value)
                    {
                        e.consumo_llano = Convert.ToInt32(reader["LLANO"]);
                        if (e.consumo < 0)
                            e.consumo_llano = e.consumo_llano * -1;
                    }

                    if (reader["VALLE"] != System.DBNull.Value)
                    {
                        e.consumo_valle = Convert.ToInt32(reader["VALLE"]);
                        if (e.consumo < 0)
                            e.consumo_valle = e.consumo_valle * -1;
                    }

                    if (reader["CONSUMO_ACTIVA1"] != System.DBNull.Value)
                    {
                        e.consumo_activa1 = Convert.ToInt32(reader["CONSUMO_ACTIVA1"]);
                        if (e.consumo < 0)
                            e.consumo_activa1 = e.consumo_activa1 * -1;
                    }

                    if (reader["CONSUMO_ACTIVA2"] != System.DBNull.Value)
                    {
                        e.consumo_activa2 = Convert.ToInt32(reader["CONSUMO_ACTIVA2"]);
                        if (e.consumo < 0)
                            e.consumo_activa2 = e.consumo_activa2 * -1;
                    }

                    if (reader["CONSUMO_ACTIVA3"] != System.DBNull.Value)
                    {
                        e.consumo_activa3 = Convert.ToInt32(reader["CONSUMO_ACTIVA3"]);
                        if (e.consumo < 0)
                            e.consumo_activa3 = e.consumo_activa3 * -1;
                    }


                    if (reader["PUNTA"] != System.DBNull.Value)
                    {
                        e.consumo_punta = Convert.ToInt32(reader["PUNTA"]);
                        if (e.consumo < 0)
                            e.consumo_punta = e.consumo_punta * -1;
                    }

                    if (reader["LLANO"] != System.DBNull.Value)
                    {
                        e.consumo_llano = Convert.ToInt32(reader["LLANO"]);
                        if (e.consumo < 0)
                            e.consumo_llano = e.consumo_llano * -1;
                    }

                    if (reader["VALLE"] != System.DBNull.Value)
                    {
                        e.consumo_valle = Convert.ToInt32(reader["VALLE"]);
                        if (e.consumo < 0)
                            e.consumo_valle = e.consumo_valle * -1;
                    }


                    if (e.consumo_activa1 == 0)
                        e.consumo_activa1 = e.consumo_punta;

                    if (e.consumo_activa2 == 0)
                        e.consumo_activa2 = e.consumo_llano;

                    if (e.consumo_activa3 == 0)
                        e.consumo_activa3 = e.consumo_valle;

                    if (reader["CONSUMO_ACTIVA4"] != System.DBNull.Value)
                    {
                        e.consumo_activa4 = Convert.ToInt32(reader["CONSUMO_ACTIVA4"]);


                        if (e.consumo < 0)
                            e.consumo_activa4 = e.consumo_activa4 * -1;
                    }

                    //if (e.consumo !=
                    //    (e.consumo_activa1 +
                    //    e.consumo_activa2 +
                    //    e.consumo_activa3 +
                    //    e.consumo_activa4))
                    //{
                    //    e.consumo_activa4 = e.consumo_activa4 / 1000;
                    //}

                    if (reader["CUPSREE"] != System.DBNull.Value)
                        e.cupsree = reader["CUPSREE"].ToString();

                    e.cnifdnic = reader["CNIFDNIC"].ToString();

                    fact_list.Add(e);

                }

                db.CloseConnection();
                if (hayDatos)
                {
                    //CargaDatosDetalleMT(usarFechaFactura, fd, fh);
                    //CalculosMT();
                }

                else
                    MessageBox.Show("La consulta no devuelve datos.",
                     "InformeErse.CargaDatos",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "InformeErse.CargaDatos",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

        }

        public void CargaDatosBTN_BTE(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            bool hayDatos = false;

            string clave = "";
                      

            string consumo_temp = "";

            try
            {
                #region query
                strSql = "SELECT e.descripcion AS Empresa," 
                    + " f.CNIFDNIC, f.DAPERSOC, f.CUPSREE,"
                    + " if (f.VCUOVAFA is null, 0, f.VCUOVAFA)  AS CONSUMO,"
                    + " if (f.VCONSACP is null, 0, f.VCONSACP) AS PUNTA,"
                    + " if (f.VCONSACL is null, 0, f.VCONSACL) AS LLANO,"
                    + " if (f.VCONSACV is null, 0, f.VCONSACV) AS VALLE,"
                    + " if (f.SUPERVALLE is null, 0, f.SUPERVALLE) AS SUPERVALLE,"
                    + " if (f.VCONATH1 is null, 0, f.VCONATH1) AS CONSUMO_ACTIVA1,"
                    + " if (f.VCONATH2 is null, 0, f.VCONATH2) AS CONSUMO_ACTIVA2,"
                    + " if (f.VCONATH3 is null, 0, f.VCONATH3) AS CONSUMO_ACTIVA3,"
                    + " if (f.VCONATH4 is null, 0, f.VCONATH4) AS CONSUMO_ACTIVA4,"
                    + " if (f.VCONATH5 is null, 0, f.VCONATH5) AS CONSUMO_ACTIVA5,"
                    + " if (f.VCONATH6 is null, 0, f.VCONATH6) AS CONSUMO_ACTIVA6"                    
                    + " from fact.fo f"
                    + " inner join fo_empresas e on"
                    + " e.empresa_id = f.ID_ENTORNO"               
                    + " where e.descripcion in ('BTN-Portugal','BTE-Portugal') and"
                    + " f.TIPONEGOCIO = 'L' ";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " and (f.FFACTDES <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                     + " f.FFACTHAS >= '" + fd.ToString("yyyy-MM-dd") + "')";
                }

                strSql = strSql + " GROUP BY f.CREFEREN, f.SECFACTU, f.TESTFACT";

                #endregion


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    hayDatos = true;
                    EndesaEntity.facturacion.ErseBTN e =
                        new EndesaEntity.facturacion.ErseBTN();



                    e.empresa = reader["Empresa"].ToString();
                   
                    if (reader["CONSUMO"] != System.DBNull.Value)
                        e.consumo = Convert.ToInt32(reader["CONSUMO"]);
                                      



                    if (reader["CONSUMO_ACTIVA1"] != System.DBNull.Value)
                    {
                        e.consumo_activa1 = Convert.ToInt32(reader["CONSUMO_ACTIVA1"]);
                        if (e.consumo < 0)
                            e.consumo_activa1 = e.consumo_activa1 * -1;
                    }

                    if (reader["CONSUMO_ACTIVA2"] != System.DBNull.Value)
                    {
                        e.consumo_activa2 = Convert.ToInt32(reader["CONSUMO_ACTIVA2"]);
                        if (e.consumo < 0)
                            e.consumo_activa2 = e.consumo_activa2 * -1;
                    }

                    if (reader["CONSUMO_ACTIVA3"] != System.DBNull.Value)
                    {
                        e.consumo_activa3 = Convert.ToInt32(reader["CONSUMO_ACTIVA3"]);
                        if (e.consumo < 0)
                            e.consumo_activa3 = e.consumo_activa3 * -1;
                    }

                    if (reader["CONSUMO_ACTIVA4"] != System.DBNull.Value)
                    {
                        e.consumo_activa4 = Convert.ToInt32(reader["CONSUMO_ACTIVA4"]);
                        if (e.consumo < 0)
                            e.consumo_activa4 = e.consumo_activa4 * -1;
                    }

                    if (reader["PUNTA"] != System.DBNull.Value)
                    {
                        e.consumo_punta = Convert.ToInt32(reader["PUNTA"]);
                        if (e.consumo < 0)
                            e.consumo_punta = e.consumo_punta * -1;
                    }

                    if (reader["LLANO"] != System.DBNull.Value)
                    {
                        e.consumo_llano = Convert.ToInt32(reader["LLANO"]);
                        if (e.consumo < 0)
                            e.consumo_llano = e.consumo_llano * -1;
                    }

                    if (reader["VALLE"] != System.DBNull.Value)
                    {
                        e.consumo_valle = Convert.ToInt32(reader["VALLE"]);
                        if (e.consumo < 0)
                            e.consumo_valle = e.consumo_valle * -1;
                    }
                    if (reader["SUPERVALLE"] != System.DBNull.Value)
                    {
                        e.consumo_supervalle = Convert.ToInt32(reader["SUPERVALLE"]);
                        if (e.consumo < 0)
                            e.consumo_supervalle = e.consumo_supervalle * -1;
                    }


                    if (e.consumo_activa1 == 0)
                        e.consumo_activa1 = e.consumo_punta;

                    if (e.consumo_activa2 == 0)
                        e.consumo_activa2 = e.consumo_llano;

                    if (e.consumo_activa3 == 0)
                        e.consumo_activa3 = e.consumo_valle;

                    if (reader["CUPSREE"] != System.DBNull.Value)
                        e.cupsree = reader["CUPSREE"].ToString();

                   

                    e.cnifdnic = reader["CNIFDNIC"].ToString();
                    e.dapersoc = reader["DAPERSOC"].ToString();

                    fact_list.Add(e);

                }

                db.CloseConnection();
                if (hayDatos)
                {
                    //CargaDatosDetalleBTN(usarFechaFactura, fd, fh);
                    //CalculosBTN();
                    GuardaDatos(fact_list);
                }

                else
                    MessageBox.Show("La consulta no devuelve datos.",
                     "InformeErse.CargaDatos",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "InformeErse.CargaDatos",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

        }
        private void GuardaDatos(List<EndesaEntity.facturacion.ErseBTN> lista)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;

            try
            {
                foreach(EndesaEntity.facturacion.ErseBTN p in lista)
                {
                    x++;
                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO facturas_dgeg ");
                        sb.Append("(Empresa, CNIFDNIC, DAPERSOC, CUPSREE, VCUOVAFA,");
                        sb.Append("PUNTA, LLANO, VALLE, SUPERVALLE,");
                        sb.Append("A_P1, A_P2, A_P3, A_P4, A_P5, A_P6) values");

                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.empresa).Append("',");
                    sb.Append("'").Append(p.cnifdnic).Append("',");
                    sb.Append("'").Append(p.dapersoc).Append("',");
                    sb.Append("'").Append(p.cupsree).Append("',");
                    sb.Append(p.consumo.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_punta.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_llano.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_valle.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_supervalle.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_activa1.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_activa2.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_activa3.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_activa4.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_activa5.ToString().Replace(",", ".")).Append(",");
                    sb.Append(p.consumo_activa6.ToString().Replace(",", ".")).Append("),");


                    if (x == 500)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }

                }

                if (x > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }




            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
