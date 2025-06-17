using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Org.BouncyCastle.Crypto.Tls;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class InformeErse
    {
        utilidades.Param param;

        public DateTime fd { get; set; }
        public DateTime fh { get; set; }

        public bool usarFechaFactura { get; set; }

        public Dictionary<string, EndesaEntity.facturacion.ErseMT> fact_list_MT { get; set; }
        public Dictionary<string, EndesaEntity.facturacion.ErseBTE> fact_list_BTE { get; set; }
        public Dictionary<string, EndesaEntity.facturacion.ErseBTN> fact_list_BTN { get; set; }
        public Dictionary<int, EndesaEntity.facturacion.ErseSAP> fact_list_SAP { get; set; }
        public Dictionary<string, EndesaEntity.facturacion.Tcon> tcon_list { get; set; }

        public InformeErse()
        {
            param = new utilidades.Param("erse_param", MySQLDB.Esquemas.FAC);

            fd = new DateTime();
            fh = new DateTime();

            //fact_list_MT = new Dictionary<string, EndesaEntity.facturacion.ErseMT>();
            //fact_list_BTE = new Dictionary<string, EndesaEntity.facturacion.ErseBTE>();
            //fact_list_BTN = new Dictionary<string, EndesaEntity.facturacion.ErseBTN>();
            fact_list_SAP = new Dictionary<int, EndesaEntity.facturacion.ErseSAP>();
            //tcon_list = new Dictionary<string, EndesaEntity.facturacion.Tcon>();
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
                    strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }


                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    hayDatos = true;
                    EndesaEntity.facturacion.ErseMT e = new EndesaEntity.facturacion.ErseMT();
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

                    fact_list_MT.Add(clave, e);

                }

                db.CloseConnection();
                if (hayDatos)
                {
                    CargaDatosDetalleMT(usarFechaFactura, fd, fh);
                    CalculosMT();
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
        public void CargaDatosBTE(bool usarFechaFactura, DateTime fd, DateTime fh)
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
                //strSql = "select f.CCOUNIPS, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                //+ " f.VCUOFIFA as POTENCIA, f.VCUOVAFA as CONSUMO, f.IISE as ISE, f.IFACTURA,"
                //+ " f.TFACTURA, f.TESTFACT,"
                //+ " (f.VCONSACP + f.VCONATH1) as CONSUMO_ACTIVA1,"
                //+ " (f.VCONSACL + f.VCONATH2) as CONSUMO_ACTIVA2,"
                //+ " (f.VCONSACV + f.VCONATH3) as CONSUMO_ACTIVA3,"
                //+ " (f.INTEG_SUPERVALLE + f.VCONATH4) / 1000 as CONSUMO_ACTIVA4,"
                //+ " f.CUPSREE, f.CNIFDNIC"
                //+ " from fo_empresas e"
                //+ " inner join ff_f_temp f on"
                //+ " f.CEMPTITU = e.cemptitu and"
                //+ " f.INDEMPRE = e.segmento"
                //+ " where e.descripcion = 'BTE-Portugal' and"
                //+ " f.TFACTURA in (1,2,3)";

                strSql = "select f.CEMPTITU, f.CCOUNIPS, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                + " f.VCUOFIFA as POTENCIA, f.VCUOVAFA as CONSUMO, f.IISE as ISE, f.IFACTURA,"
                + " f.TFACTURA, f.TESTFACT, f.VCONSACP as PUNTA, f.VCONSACL as LLANO, f.VCONSACV as VALLE,"
                + " f.VCONATH1 as CONSUMO_ACTIVA1,"
                + " f.VCONATH2 as CONSUMO_ACTIVA2,"
                + " f.VCONATH3 as CONSUMO_ACTIVA3,"
                + " f.INTEG_SUPERVALLE as CONSUMO_ACTIVA4,"
                + " f.CUPSREE, f.CNIFDNIC"
                + " from fo_empresas e"
                + " inner join ff_f_temp f on"
                + " f.CEMPTITU = e.cemptitu and"
                + " f.INDEMPRE = e.segmento"
                + " where e.descripcion = 'BTE-Portugal' and"
                + " f.TFACTURA in (1,2,3)";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                #endregion

                 db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    hayDatos = true;
                    EndesaEntity.facturacion.ErseBTE e = new EndesaEntity.facturacion.ErseBTE();
                    EndesaEntity.facturacion.ClaveFactura c = new EndesaEntity.facturacion.ClaveFactura();

                    if (fact_list_BTE.Count() == 28787)
                        clave = "";


                    clave = reader["CEMPTITU"].ToString();
                    clave += "_"+ reader["CREFEREN"].ToString();
                    clave += "_" + reader["SECFACTU"].ToString();
                    clave += "_" + reader["TESTFACT"].ToString();

                    e.creferen = reader["CREFEREN"].ToString();
                    e.secfactu = reader["SECFACTU"].ToString();
                    e.ccounips = reader["CCOUNIPS"].ToString();
                    e.cfactura = reader["CFACTURA"] != System.DBNull.Value ? reader["CFACTURA"].ToString() : null;

                    if (reader["FFACTURA"] != System.DBNull.Value)
                        e.ffactura = Convert.ToDateTime(reader["FFACTURA"]);
                    if (reader["FFACTDES"] != System.DBNull.Value)
                        e.ffactdes = Convert.ToDateTime(reader["FFACTDES"]);
                    if (reader["FFACTHAS"] != System.DBNull.Value)
                        e.ffacthas = Convert.ToDateTime(reader["FFACTHAS"]);
                    if (reader["POTENCIA"] != System.DBNull.Value)
                        e.potencia = Convert.ToDouble(reader["POTENCIA"]);
                    if (reader["CONSUMO"] != System.DBNull.Value)
                        e.consumo = Convert.ToInt32(reader["CONSUMO"]);
                    if (reader["ISE"] != System.DBNull.Value)
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
                        consumo_temp = reader["CONSUMO_ACTIVA4"].ToString().Replace(".", "");
                        e.consumo_activa4 = Convert.ToInt32(consumo_temp);


                        if (e.consumo < 0)
                            e.consumo_activa4 = e.consumo_activa4 * -1;
                    }

                    if(e.consumo != 
                        (e.consumo_activa1 + 
                        e.consumo_activa2 +
                        e.consumo_activa3 +
                        e.consumo_activa4))
                    {
                        e.consumo_activa4 = e.consumo_activa4 / 1000;
                    }


                    if (reader["CUPSREE"] != System.DBNull.Value)
                        e.cupsree = reader["CUPSREE"].ToString();

                    e.cnifdnic = reader["CNIFDNIC"].ToString();

                    fact_list_BTE.Add(clave, e);

                }

                db.CloseConnection();
                if (hayDatos)
                {
                    CargaDatosDetalleBTE(usarFechaFactura, fd, fh);
                    CalculosBTE();
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

        public void CargaDatosBTN(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            bool hayDatos = false;
            
            string clave = "";

            double iva_intermedio 
                = Convert.ToDouble(param.GetValue("IVA_INTERMEDIO", DateTime.Now, DateTime.Now));

            string consumo_temp = "";

            try
            {
                #region query
                strSql = "select f.CEMPTITU, f.CCOUNIPS, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                + " f.VCUOFIFA as POTENCIA, f.VCUOVAFA as CONSUMO, f.ISE, f.IFACTURA, f.IVA, f.IIMPUES3,"
               + " f.TFACTURA, f.TESTFACT, f.VCONSACP as PUNTA, f.VCONSACL as LLANO, f.VCONSACV as VALLE,"
                + " f.VCONATH1 as CONSUMO_ACTIVA1,"
                + " f.VCONATH2 as CONSUMO_ACTIVA2,"
                + " f.VCONATH3 as CONSUMO_ACTIVA3,"
                + " f.SUPERVALLE as CONSUMO_ACTIVA4,"
                + " f.CUPSREE, f.CNIFDNIC,"
                + " SUM(t.ICONFAC) AS BASE_IVA_INTERMEDIO,"
                + " ROUND(SUM(t.ICONFAC) * " 
                + (iva_intermedio / 100).ToString().Replace(",",".") 
                + " ,2) AS IVA_INTERMEDIO"
                + " from fact.fo f"
                + " inner join fo_empresas e on"
                + " e.empresa_id = f.ID_ENTORNO"
                + " LEFT OUTER JOIN fact.fo_tcon t ON"
                + " t.CREFEREN = f.CREFEREN AND"
                + " t.SECFACTU = f.SECFACTU AND"
                + " t.TESTFACT = f.TESTFACT AND"
                + " t.TCONFAC IN (760, 764, 768)"
                + " where e.descripcion = 'BTN-Portugal' and"
                + " f.TFACTURA in (1, 2, 3)";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
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
                    EndesaEntity.facturacion.ClaveFactura c =
                        new EndesaEntity.facturacion.ClaveFactura();

                    clave = reader["CEMPTITU"].ToString();
                    clave += "_" + reader["CREFEREN"].ToString();
                    clave += "_" + reader["SECFACTU"].ToString();
                    clave += "_" + reader["TESTFACT"].ToString();

                    e.creferen = reader["CREFEREN"].ToString();
                    e.secfactu = reader["SECFACTU"].ToString();
                    e.ccounips = reader["CCOUNIPS"].ToString();
                    e.cfactura = reader["CFACTURA"] != System.DBNull.Value ? reader["CFACTURA"].ToString() : null;

                    if (reader["FFACTURA"] != System.DBNull.Value)
                        e.ffactura = Convert.ToDateTime(reader["FFACTURA"]);
                    if (reader["FFACTDES"] != System.DBNull.Value)
                        e.ffactdes = Convert.ToDateTime(reader["FFACTDES"]);
                    if (reader["FFACTHAS"] != System.DBNull.Value)
                        e.ffacthas = Convert.ToDateTime(reader["FFACTHAS"]);
                    if (reader["POTENCIA"] != System.DBNull.Value)
                        e.potencia = Convert.ToDouble(reader["POTENCIA"]);
                    if (reader["CONSUMO"] != System.DBNull.Value)
                        e.consumo = Convert.ToInt32(reader["CONSUMO"]);
                    if (reader["ISE"] != System.DBNull.Value)
                        e.ise = Convert.ToDouble(reader["ISE"]);

                    e.ifactura = Convert.ToDouble(reader["IFACTURA"]);
                    e.tfactura = Convert.ToInt32(reader["TFACTURA"]);
                    e.testfact = reader["TESTFACT"].ToString();

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

                    if (e.consumo_activa1 == 0)
                        e.consumo_activa1 = e.consumo_punta;

                    if (e.consumo_activa2 == 0)
                        e.consumo_activa2 = e.consumo_llano;

                    if (e.consumo_activa3 == 0)
                        e.consumo_activa3 = e.consumo_valle;

                    if (reader["CUPSREE"] != System.DBNull.Value)
                        e.cupsree = reader["CUPSREE"].ToString();

                    if (reader["IIMPUES3"] != System.DBNull.Value)
                        e.iimpues3 = Convert.ToDouble(reader["IIMPUES3"]);

                    e.cnifdnic = reader["CNIFDNIC"].ToString();

                    if (reader["BASE_IVA_INTERMEDIO"] != System.DBNull.Value)
                        e.base_iva_intermedio = Convert.ToDouble(reader["BASE_IVA_INTERMEDIO"]);

                    if (reader["IVA_INTERMEDIO"] != System.DBNull.Value)
                        e.iva_intermedio = Convert.ToDouble(reader["IVA_INTERMEDIO"]);

                    fact_list_BTN.Add(clave, e);

                }

                db.CloseConnection();
                if (hayDatos)
                {
                    CargaDatosDetalleBTN(usarFechaFactura, fd, fh);
                    CalculosBTN();
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
        public void CargaDatosSAP(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader reader;
            string strSql = "";
            bool hayDatos = false;
            int  clave = 0;


            try
            {

                #region new query ERSE SAP (GUS add. 15/04/2025)
                strSql = ConsultaSAP();
                
                strSql = strSql + " and (FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                #endregion

                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                // command.Connection.ConnectionTimeout = 400; No se puede establecer con la conexión abierta, lo hacemos en RedShiftServer.cs
                command.CommandTimeout = 3000;
                reader = command.ExecuteReader();

                
                while (reader.Read())
                {

                    hayDatos = true;
                    EndesaEntity.facturacion.ErseSAP e = new EndesaEntity.facturacion.ErseSAP();


                    clave++;

                    #region Añadimos registro a diccionario

                    //e.creferen = reader["CREFEREN"].ToString();
                    //e.secfactu = reader["SECFACTU"].ToString();
                    //e.ccounips = reader["CCOUNIPS"].ToString();
                    //e.cfactura = reader["CFACTURA"] != System.DBNull.Value ? reader["CFACTURA"].ToString() : null;
                    
                    if (reader["de_seg_merc_por"] != System.DBNull.Value)
                        e.de_seg_merc_por = reader["de_seg_merc_por"].ToString();
                    if (reader["cd_estado_fact"] != System.DBNull.Value)
                        e.cd_estado_fact = reader["cd_estado_fact"].ToString();
                    if (reader["fh_ult_ejec"] != System.DBNull.Value)
                        e.fh_ult_ejec = Convert.ToDateTime(reader["fh_ult_ejec"]);
                    if (reader["cif"] != System.DBNull.Value)
                        e.cif = reader["cif"].ToString();
                    if (reader["cfactura"] != System.DBNull.Value)
                        e.cfactura = reader["cfactura"].ToString();
                    if (reader["ffactura"] != System.DBNull.Value)
                        e.ffactura = Convert.ToDateTime(reader["ffactura"]);
                    if (reader["ffactdes"] != System.DBNull.Value)
                        e.ffactdes = Convert.ToDateTime(reader["ffactdes"]);
                    if (reader["ffacthas"] != System.DBNull.Value)
                        e.ffacthas = Convert.ToDateTime(reader["ffacthas"]);
                    if (reader["ifactura"] != System.DBNull.Value)
                        e.ifactura = Convert.ToDouble(reader["ifactura"]);
                    if (reader["impuesto1"] != System.DBNull.Value)
                        e.impuesto1 = Convert.ToDouble(reader["impuesto1"]);
                    if (reader["consumo"] != System.DBNull.Value)
                        e.consumo = Convert.ToDouble(reader["consumo"]);
                    if (reader["ise"] != System.DBNull.Value)
                        e.ise = Convert.ToDouble(reader["ise"]);
                    if (reader["base_iva_normal"] != System.DBNull.Value)
                        e.base_iva_normal = Convert.ToDouble(reader["base_iva_normal"]);
                    if (reader["iva_normal"] != System.DBNull.Value)
                        e.iva_normal = Convert.ToDouble(reader["iva_normal"]);
                    if (reader["base_iva_reducido"] != System.DBNull.Value)
                        e.base_iva_reducido = Convert.ToDouble(reader["base_iva_reducido"]);
                    if (reader["iva_reducido"] != System.DBNull.Value)
                        e.iva_reducido = Convert.ToDouble(reader["iva_reducido"]);
                    if (reader["cav"] != System.DBNull.Value)
                        e.cav = Convert.ToDouble(reader["cav"]);
                    if (reader["importe_redes"] != System.DBNull.Value)
                        e.importe_redes = Convert.ToDouble(reader["importe_redes"]);
                    if (reader["consumo_activa1"] != System.DBNull.Value)
                        e.consumo_activa1 = Convert.ToDouble(reader["consumo_activa1"]);
                    if (reader["consumo_activa2"] != System.DBNull.Value)
                        e.consumo_activa2 = Convert.ToDouble(reader["consumo_activa2"]);
                    if (reader["consumo_activa3"] != System.DBNull.Value)
                        e.consumo_activa3 = Convert.ToDouble(reader["consumo_activa3"]);
                    if (reader["consumo_activa4"] != System.DBNull.Value)
                        e.consumo_activa4 = Convert.ToDouble(reader["consumo_activa4"]);
                    if (reader["cupsree"] != System.DBNull.Value)
                        e.cupsree = reader["cupsree"].ToString();
                    if (reader["potencia"] != System.DBNull.Value)
                        e.potencia = Convert.ToDouble(reader["potencia"]);

                    #endregion

                    fact_list_SAP.Add(clave, e);

                }

                db.CloseConnection();
                if (!hayDatos)
                {
                    MessageBox.Show("La consulta no devuelve datos.",
                     "InformeErse.CargaDatos",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "InformeErse.CargaDatos",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

        }


        private void CargaDatosDetalleMT(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string clave;

            try
            {
                #region
                strSql = "select t.*"
                + " from fo_empresas e"
                //+ " inner join ff_f_temp f on"
                + " inner join fo f on"
                + " f.CEMPTITU = e.CEMPTITU and "
                + " f.INDEMPRE = e.SEGMENTO"
                + " inner join fo_tcon t on"
                + " t.CREFEREN = f.CREFEREN and"
                + " t.SECFACTU = f.SECFACTU and"
                + " t.TESTFACT = f.TESTFACT"
                + " where e.descripcion = 'MT-Portugal' and"
                + " f.TFACTURA in (1,2,3)";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                #endregion


                tcon_list.Clear();
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.facturacion.ClaveTcon c =
                        new EndesaEntity.facturacion.ClaveTcon();
                    clave = reader["CREFEREN"].ToString();
                    clave += reader["SECFACTU"].ToString();
                    clave += reader["TESTFACT"].ToString();
                    clave += reader["TCONFAC"].ToString();

                    EndesaEntity.facturacion.Tcon t = new EndesaEntity.facturacion.Tcon();
                    t.iconfac = Convert.ToDouble(reader["ICONFAC"]);

                    EndesaEntity.facturacion.Tcon tt;
                    if (!tcon_list.TryGetValue(clave, out tt))
                        tcon_list.Add(clave, t);

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
            "InformeErse.CargaDatosDetalle",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
            }
        }
        private void CargaDatosDetalleBTE(bool usarFechaFactura, DateTime fd, DateTime fh)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string clave;
            
            try
            {

                #region query
                strSql = "select t.*"
                + " from fo_empresas e"
                //+ " inner join ff_f_temp f on"
                + " inner join fo f on"
                + " f.CEMPTITU = e.CEMPTITU and "
                + " f.INDEMPRE = e.SEGMENTO"
                + " inner join fo_tcon t on"
                + " t.CREFEREN = f.CREFEREN and"
                + " t.SECFACTU = f.SECFACTU and"
                + " t.TESTFACT = f.TESTFACT"
                + " where e.descripcion = 'BTE-Portugal'";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                #endregion


                tcon_list.Clear();

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.facturacion.ClaveTcon c = new EndesaEntity.facturacion.ClaveTcon();
                    clave = reader["CREFEREN"].ToString();
                    clave += reader["SECFACTU"].ToString();
                    clave += reader["TESTFACT"].ToString();
                    clave += reader["TCONFAC"].ToString();

                    EndesaEntity.facturacion.Tcon t = new EndesaEntity.facturacion.Tcon();
                    t.iconfac = Convert.ToDouble(reader["ICONFAC"]);

                    EndesaEntity.facturacion.Tcon tt;
                    if (!tcon_list.TryGetValue(clave, out tt))
                        tcon_list.Add(clave, t);

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
            "InformeErse.CargaDatosDetalle",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
            }
        }
        private void CargaDatosDetalleBTN(bool usarFechaFactura, DateTime fd, DateTime fh)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string clave;            

            try
            {
                #region query
                strSql = "select t.*"
                + " from fo_empresas e"
                + " inner join fo f on"
                + " f.CEMPTITU = e.CEMPTITU and "
                + " f.INDEMPRE = e.SEGMENTO"
                + " inner join fo_tcon t on"
                + " t.CREFEREN = f.CREFEREN and"
                + " t.SECFACTU = f.SECFACTU and"
                + " t.TESTFACT = f.TESTFACT"
                + " where e.descripcion = 'BTN-Portugal'";

                if (usarFechaFactura)
                {
                    strSql = strSql + " and (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " and (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                #endregion


                tcon_list.Clear();

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    EndesaEntity.facturacion.ClaveTcon c = new EndesaEntity.facturacion.ClaveTcon();
                    clave = reader["CREFEREN"].ToString();
                    clave += reader["SECFACTU"].ToString();
                    clave += reader["TESTFACT"].ToString();
                    // clave += reader["TCONFAC"].ToString();


                    EndesaEntity.facturacion.ErseBTN o;
                    if (fact_list_BTN.TryGetValue(clave, out o))
                    {
                        double oo;
                        if (!o.dic.TryGetValue(Convert.ToInt32(reader["TCONFAC"]), out oo))
                            o.dic.Add(Convert.ToInt32(reader["TCONFAC"]), Convert.ToDouble(reader["ICONFAC"]));
                    }

                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
            "InformeErse.CargaDatosDetalle",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
            }
        }

        private void CalculosMT()
        {
            string clave;

            try
            {
                foreach (KeyValuePair<string, EndesaEntity.facturacion.ErseMT> p in fact_list_MT)
                {

                    clave = p.Key;
                    clave += param.GetValue("TAV", p.Value.ffactdes, p.Value.ffacthas);

                    p.Value.base_iva_r = 0;
                    p.Value.iva_r = 0;

                    EndesaEntity.facturacion.Tcon tt;
                    if (tcon_list.TryGetValue(clave, out tt))
                    {
                        p.Value.cav = tt.iconfac;
                        p.Value.base_iva_r = tt.iconfac;
                        p.Value.iva_r = (p.Value.base_iva_r *
                            (Convert.ToDouble(param.GetValue("R_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                        //p.Value.base_iva_n = (p.Value.ifactura - p.Value.base_iva_r - p.Value.iva_r) /
                        //    (1 + (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));
                        //p.Value.iva_n = (p.Value.base_iva_n *
                        //    (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                    }

                    p.Value.base_iva_n = (p.Value.ifactura - p.Value.base_iva_r - p.Value.iva_r) /
                        (1 + (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                    p.Value.iva_n = (p.Value.base_iva_n *
                        (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                    p.Value.iva_r = Math.Round(p.Value.iva_r, 2);
                    p.Value.base_iva_n = Math.Round(p.Value.base_iva_n, 2);
                    p.Value.iva_n = Math.Round(p.Value.iva_n, 2);

                }
            }
            catch (KeyNotFoundException e)
            {
                MessageBox.Show(e.Message,
                   "InformeErse.CalculosMT",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }

        private void CalculosBTE()
        {
            string clave;

            try
            {
                foreach (KeyValuePair<string, EndesaEntity.facturacion.ErseBTE> p in fact_list_BTE)
                {

                    clave = p.Key;
                    clave += param.GetValue("TAV2", p.Value.ffactdes, p.Value.ffacthas);

                    p.Value.base_iva_r = 0;
                    p.Value.iva_r = 0;

                    EndesaEntity.facturacion.Tcon tt;
                    if (tcon_list.TryGetValue(clave, out tt))
                    {
                        p.Value.cav = tt.iconfac;
                        p.Value.base_iva_r = tt.iconfac;
                        p.Value.iva_r = (p.Value.base_iva_r *
                            (Convert.ToDouble(param.GetValue("R_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                    }

                    p.Value.base_iva_n = (p.Value.ifactura - p.Value.base_iva_r - p.Value.iva_r) /
                            (1 + (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));
                    p.Value.iva_n = (p.Value.base_iva_n *
                        (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                    p.Value.iva_r = Math.Round(p.Value.iva_r, 2);
                    p.Value.base_iva_n = Math.Round(p.Value.base_iva_n, 2);
                    p.Value.iva_n = Math.Round(p.Value.iva_n, 2);

                    clave = p.Key;
                    clave += param.GetValue("DGGE", p.Value.ffactdes, p.Value.ffacthas);

                    if (tcon_list.TryGetValue(clave, out tt))
                    {
                        p.Value.tasa_dge = tt.iconfac;
                    }
                                       


                }
            }
            catch (KeyNotFoundException e)
            {
                MessageBox.Show(e.Message,
                  "InformeErse.CalculosBTE",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }
        private void CalculosBTN()
        {
            int clave;
            int clave2;
            int clave_dgge_btn;



            try
            {

                foreach (KeyValuePair<string, EndesaEntity.facturacion.ErseBTN> p in fact_list_BTN)
                {
                    
                    clave = Convert.ToInt32(param.GetValue("TAV_BTN", p.Value.ffactdes, p.Value.ffacthas));
                    clave2 = 1001;

                    p.Value.base_iva_r = 0;
                    p.Value.iva_r = p.Value.iimpues3 - p.Value.iva_intermedio;

                    p.Value.base_iva_r = Math.Round(p.Value.iva_r / (Convert.ToDouble(param.GetValue("R_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100), 2);


                    double iconfac;
                    // Buscamos el importe de la Tasa Audivisual                    
                    if (p.Value.dic.TryGetValue(clave, out iconfac))
                    {

                        p.Value.cav = iconfac;
                        // Si el importe de la tasa Audiv es <> de la base_iva_r entonces debemos sumar
                        // el termino de pontencia

                        if (iconfac != p.Value.base_iva_r)
                        {
                            p.Value.base_iva_r = 0;
                            foreach (KeyValuePair<int, double> pp in p.Value.dic)
                            {
                                if (pp.Key == clave || pp.Key == clave2)
                                    p.Value.base_iva_r += pp.Value;
                            }
                        }
                                                
                    }


                    clave_dgge_btn = Convert.ToInt32(param.GetValue("DGGE_BTN", p.Value.ffactdes, p.Value.ffacthas));
                    if (p.Value.dic.TryGetValue(clave_dgge_btn, out iconfac))
                    {
                        p.Value.tasa_dge = iconfac;                      
                    }






                    p.Value.base_iva_n = (p.Value.ifactura - p.Value.base_iva_r - p.Value.iva_r - p.Value.base_iva_intermedio -  p.Value.iva_intermedio) /
                        (1 + (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));
                    p.Value.iva_n = (p.Value.base_iva_n *
                        (Convert.ToDouble(param.GetValue("S_VAT", p.Value.ffactdes, p.Value.ffacthas)) / 100));

                    p.Value.iva_r = Math.Round(p.Value.iva_r, 2);
                    p.Value.base_iva_n = Math.Round(p.Value.base_iva_n, 2);
                    p.Value.iva_n = Math.Round(p.Value.iva_n, 2);

                }
            }
            catch (KeyNotFoundException e)
            {
                MessageBox.Show(e.Message,
                  "InformeErse.CalculosBTN",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        public void ExportExcel2(string fichero, string entorno)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add(entorno);




            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            if (entorno == "MT")
            {
                //headerFont.SetFromFont(new Font("Times New Roman", 12)); //Do this first
                headerFont.Bold = true;
                //headerFont.Italic = true;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
                workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                workSheet.Cells[f, c].Value = "CFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTDES"; c++;
                workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO"; c++;
                workSheet.Cells[f, c].Value = "ISE"; c++;
                workSheet.Cells[f, c].Value = "IFACTURA"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "IVA REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA NORMAL"; c++;
                workSheet.Cells[f, c].Value = "IVA NORMAL"; c++;
                workSheet.Cells[f, c].Value = "CAV"; c++;
                workSheet.Cells[f, c].Value = "TFACTURA"; c++;
                workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "CUPSREE"; c++;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            }
            else if (entorno == "BTE")
            {
                //headerFont.SetFromFont(new Font("Times New Roman", 12)); //Do this first
                headerFont.Bold = true;
                //headerFont.Italic = true;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
                workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                workSheet.Cells[f, c].Value = "CFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTDES"; c++;
                workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO"; c++;
                workSheet.Cells[f, c].Value = "ISE"; c++;
                workSheet.Cells[f, c].Value = "IFACTURA"; c++;
                workSheet.Cells[f, c].Value = "TASA DGE"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "IVA REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA NORMAL"; c++;
                workSheet.Cells[f, c].Value = "IVA NORMAL"; c++;
                workSheet.Cells[f, c].Value = "CAV"; c++;
                workSheet.Cells[f, c].Value = "TFACTURA"; c++;
                workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "CUPSREE"; c++;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            }
            else
            {
                //headerFont.SetFromFont(new Font("Times New Roman", 12)); //Do this first
                headerFont.Bold = true;
                //headerFont.Italic = true;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
                workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                workSheet.Cells[f, c].Value = "CFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTDES"; c++;
                workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO"; c++;
                workSheet.Cells[f, c].Value = "ISE"; c++;
                workSheet.Cells[f, c].Value = "IFACTURA"; c++;
                workSheet.Cells[f, c].Value = "TASA DGE"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA INTERMEDIO"; c++;
                workSheet.Cells[f, c].Value = "IVA INTERMEDIO"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "IVA REDUCIDO"; c++;
                workSheet.Cells[f, c].Value = "BASE IVA NORMAL"; c++;
                workSheet.Cells[f, c].Value = "IVA NORMAL"; c++;
                workSheet.Cells[f, c].Value = "CAV"; c++;
                workSheet.Cells[f, c].Value = "TFACTURA"; c++;
                workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "CUPSREE"; c++;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
            }

            //if (entorno == "MT")
            //    workSheet.Cells["A2"].LoadFromCollection(erse.fact_list_MT.Select(z => z.Value).ToList(), false, OfficeOpenXml.Table.TableStyles.Medium20);            
            //else
            //    workSheet.Cells["A2"].LoadFromCollection(erse.fact_list_BTE.Select(z => z.Value).ToList(), false, OfficeOpenXml.Table.TableStyles.Medium20);

            if (entorno == "MT")
            {
                foreach (KeyValuePair<string, EndesaEntity.facturacion.ErseMT> p in fact_list_MT)
                {
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                    workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                    workSheet.Cells[f, c].Value = p.Value.secfactu; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.potencia; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.ise; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.base_iva_r; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.iva_r; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.base_iva_n; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.iva_n; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.cav; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.tfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.testfact; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                    workSheet.Cells[f, c].Value = p.Value.cnifdnic; c++;
                }



            }
            else if (entorno == "BTE")
            {
                foreach (KeyValuePair<string, EndesaEntity.facturacion.ErseBTE> p in fact_list_BTE)
                {
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                    workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                    workSheet.Cells[f, c].Value = p.Value.secfactu; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.potencia; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.ise; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.tasa_dge; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.base_iva_r; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.iva_r; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.base_iva_n; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.iva_n; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.cav; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.tfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.testfact; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                    workSheet.Cells[f, c].Value = p.Value.cnifdnic; c++;
                }

            }
            else
            {
                foreach (KeyValuePair<string, EndesaEntity.facturacion.ErseBTN> p in fact_list_BTN)
                {
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                    workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                    workSheet.Cells[f, c].Value = p.Value.secfactu; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.potencia; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.ise; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.tasa_dge; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    if (p.Value.base_iva_intermedio > 0)                    
                        workSheet.Cells[f, c].Value = p.Value.base_iva_intermedio; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (p.Value.iva_intermedio > 0)
                        workSheet.Cells[f, c].Value = p.Value.iva_intermedio; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    
                    
                    workSheet.Cells[f, c].Value = p.Value.base_iva_r; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.iva_r; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.base_iva_n; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.iva_n; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.cav; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = p.Value.tfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.testfact; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.consumo_activa4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                    workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                    workSheet.Cells[f, c].Value = p.Value.cnifdnic; c++;
                }


            }

            var allCells = workSheet.Cells[1, 1, f, c];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            if (entorno == "MT")
                workSheet.Cells["A1:X1"].AutoFilter = true;
            else if (entorno == "BTE")
                workSheet.Cells["A1:Y1"].AutoFilter = true;
            else
                workSheet.Cells["A1:AA1"].AutoFilter = true;

            allCells.AutoFitColumns();


            excelPackage.Save();
            

        }

        public void ExportExcelSAP(string fichero)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("SAP");


            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            #region CABECERA EXCEL
          
            headerFont.Bold = true;
            workSheet.Cells[f, c].Value = "ENTORNO"; c++;
            workSheet.Cells[f, c].Value = "ESTADO FACTURA"; c++;
            workSheet.Cells[f, c].Value = "FH_ULT_EJEC"; c++;
            workSheet.Cells[f, c].Value = "CIF"; c++;
            workSheet.Cells[f, c].Value = "CFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTURA"; c++;
            workSheet.Cells[f, c].Value = "FFACTDES"; c++;
            workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
            workSheet.Cells[f, c].Value = "IFACTURA"; c++;
            //workSheet.Cells[f, c].Value = "IMPUESTO1"; c++;
            workSheet.Cells[f, c].Value = "CONSUMO"; c++;
            workSheet.Cells[f, c].Value = "ISE"; c++;
            workSheet.Cells[f, c].Value = "BASE IVA NORMAL"; c++;
            workSheet.Cells[f, c].Value = "IVA NORMAL"; c++;
            workSheet.Cells[f, c].Value = "BASE IVA REDUCIDO"; c++;
            workSheet.Cells[f, c].Value = "IVA REDUCIDO"; c++;
            workSheet.Cells[f, c].Value = "CAV"; c++;
            workSheet.Cells[f, c].Value = "IMPORTE REDES"; c++;
            workSheet.Cells[f, c].Value = "CONSUMO ACTIVA 1"; c++;
            workSheet.Cells[f, c].Value = "CONSUMO ACTIVA 2"; c++;
            workSheet.Cells[f, c].Value = "CONSUMO ACTIVA 3"; c++;
            workSheet.Cells[f, c].Value = "CONSUMO ACTIVA 4"; c++;
            workSheet.Cells[f, c].Value = "CUPSREE"; c++;
            workSheet.Cells[f, c].Value = "POTENCIA"; c++;

            #endregion

           
            #region DATOS FILAS

            foreach (KeyValuePair<int, EndesaEntity.facturacion.ErseSAP> p in fact_list_SAP)
            {
                f++;
                c = 1;
                workSheet.Cells[f, c].Value = p.Value.de_seg_merc_por; c++;
                workSheet.Cells[f, c].Value = p.Value.cd_estado_fact; c++;
                workSheet.Cells[f, c].Value = p.Value.fh_ult_ejec; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.cif; c++;
                workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                workSheet.Cells[f, c].Value = p.Value.ffactura; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.ffactdes; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.ffacthas; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.Value.ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                //workSheet.Cells[f, c].Value = p.Value.impuesto1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.consumo; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.ise; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.base_iva_normal; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.iva_normal; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.base_iva_reducido; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.iva_reducido; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.cav; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.importe_redes; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                workSheet.Cells[f, c].Value = p.Value.consumo_activa1; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.consumo_activa2; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.consumo_activa3; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.consumo_activa4; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                workSheet.Cells[f, c].Value = p.Value.potencia; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
            }

            #endregion


            var allCells = workSheet.Cells[1, 1, f, c];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Calibri", 8));
            workSheet.View.FreezePanes(2, 1);
            //workSheet.Cells["A1:V1"].AutoFilter = true;
            allCells.AutoFilter = true;
            allCells.AutoFitColumns();


            excelPackage.Save();


        }

        private string ConsultaSAP()
        {
            string consulta;
            
            #region query ERSE SAP
            consulta = "SELECT de_seg_merc_por,cd_estado_fact, n.fh_ult_ejec, nm_id as CIF, nm_doc_oficial as CFACTURA, " +
            "fh_contab as FFACTURA,fh_desde_fact as FFACTDES, fh_hasta_fact as FFACTHAS, im_factura as IFACTURA, im_impuesto1 as IMPUESTO1, " +
            "nm_base_ise as CONSUMO, im_base_ise as ISE, " +
            "im_base_imp1 as BASE_IVA_NORMAL, im_total_cuota_1 as IVA_NORMAL, " +
            "case when im_base_imp2 =0 then im_base_imp3 else im_base_imp2 end as BASE_IVA_REDUCIDO,  " +
            "case when im_total_cuota_2=0 then im_total_cuota_3 else im_total_cuota_2 end as IVA_REDUCIDO, " +
            "case when CD_CONCEP1 in ('ZP07','ZTAUDI') then im_comcep1 else 0 end + " +
            "case when CD_CONCEP2 in ('ZP07','ZTAUDI') then im_comcep2 else 0 end + " +
            "case when CD_CONCEP3 in ('ZP07','ZTAUDI') then im_comcep3 else 0 end + " +
            "case when CD_CONCEP4 in ('ZP07','ZTAUDI') then im_comcep4 else 0 end + " +
            "case when CD_CONCEP5 in ('ZP07','ZTAUDI') then im_comcep5 else 0 end + " +
            "case when CD_CONCEP6 in ('ZP07','ZTAUDI') then im_comcep6 else 0 end + " +
            "case when CD_CONCEP7 in ('ZP07','ZTAUDI') then im_comcep7 else 0 end + " +
            "case when CD_CONCEP8 in ('ZP07','ZTAUDI') then im_comcep8 else 0 end + " +
            "case when CD_CONCEP9 in ('ZP07','ZTAUDI') then im_comcep9 else 0 end + " +
            "case when CD_CONCEP10 in ('ZP07','ZTAUDI') then im_comcep10 else 0 end + " +
            "case when CD_CONCEP11 in ('ZP07','ZTAUDI') then im_comcep11 else 0 end + " +
            "case when CD_CONCEP12 in ('ZP07','ZTAUDI') then im_comcep12 else 0 end + " +
            "case when CD_CONCEP13 in ('ZP07','ZTAUDI') then im_comcep13 else 0 end + " +
            "case when CD_CONCEP14 in ('ZP07','ZTAUDI') then im_comcep14 else 0 end + " +
            "case when CD_CONCEP15 in ('ZP07','ZTAUDI') then im_comcep15 else 0 end + " +
            "case when CD_CONCEP16 in ('ZP07','ZTAUDI') then im_comcep16 else 0 end + " +
            "case when CD_CONCEP17 in ('ZP07','ZTAUDI') then im_comcep17 else 0 end + " +
            "case when CD_CONCEP18 in ('ZP07','ZTAUDI') then im_comcep18 else 0 end + " +
            "case when CD_CONCEP19 in ('ZP07','ZTAUDI') then im_comcep19 else 0 end + " +
            "case when CD_CONCEP20 in ('ZP07','ZTAUDI') then im_comcep20 else 0 end + " +
            "case when CD_CONCEP21 in ('ZP07','ZTAUDI') then im_comcep21 else 0 end + " +
            "case when CD_CONCEP22 in ('ZP07','ZTAUDI') then im_comcep22 else 0 end + " +
            "case when CD_CONCEP23 in ('ZP07','ZTAUDI') then im_comcep23 else 0 end + " +
            "case when CD_CONCEP24 in ('ZP07','ZTAUDI') then im_comcep24 else 0 end + " +
            "case when CD_CONCEP25 in ('ZP07','ZTAUDI') then im_comcep25 else 0 end + " +
            "case when CD_CONCEP26 in ('ZP07','ZTAUDI') then im_comcep26 else 0 end + " +
            "case when CD_CONCEP27 in ('ZP07','ZTAUDI') then im_comcep27 else 0 end + " +
            "case when CD_CONCEP28 in ('ZP07','ZTAUDI') then im_comcep28 else 0 end + " +
            "case when CD_CONCEP29 in ('ZP07','ZTAUDI') then im_comcep29 else 0 end + " +
            "case when CD_CONCEP30 in ('ZP07','ZTAUDI') then im_comcep30 else 0 end + " +
            "case when CD_CONCEP31 in ('ZP07','ZTAUDI') then im_comcep31 else 0 end + " +
            "case when CD_CONCEP32 in ('ZP07','ZTAUDI') then im_comcep32 else 0 end + " +
            "case when CD_CONCEP33 in ('ZP07','ZTAUDI') then im_comcep33 else 0 end + " +
            "case when CD_CONCEP34 in ('ZP07','ZTAUDI') then im_comcep34 else 0 end + " +
            "case when CD_CONCEP35 in ('ZP07','ZTAUDI') then im_comcep35 else 0 end + " +
            "case when CD_CONCEP36 in ('ZP07','ZTAUDI') then im_comcep36 else 0 end + " +
            "case when CD_CONCEP37 in ('ZP07','ZTAUDI') then im_comcep37 else 0 end + " +
            "case when CD_CONCEP38 in ('ZP07','ZTAUDI') then im_comcep38 else 0 end + " +
            "case when CD_CONCEP39 in ('ZP07','ZTAUDI') then im_comcep39 else 0 end + " +
            "case when CD_CONCEP40 in ('ZP07','ZTAUDI') then im_comcep40 else 0 end + " +
            "case when CD_CONCEP41 in ('ZP07','ZTAUDI') then im_comcep41 else 0 end + " +
            "case when CD_CONCEP42 in ('ZP07','ZTAUDI') then im_comcep42 else 0 end + " +
            "case when CD_CONCEP43 in ('ZP07','ZTAUDI') then im_comcep43 else 0 end + " +
            "case when CD_CONCEP44 in ('ZP07','ZTAUDI') then im_comcep44 else 0 end + " +
            "case when CD_CONCEP45 in ('ZP07','ZTAUDI') then im_comcep45 else 0 end + " +
            "case when CD_CONCEP46 in ('ZP07','ZTAUDI') then im_comcep46 else 0 end + " +
            "case when CD_CONCEP47 in ('ZP07','ZTAUDI') then im_comcep47 else 0 end + " +
            "case when CD_CONCEP48 in ('ZP07','ZTAUDI') then im_comcep48 else 0 end + " +
            "case when CD_CONCEP49 in ('ZP07','ZTAUDI') then im_comcep49 else 0 end + " +
            "case when CD_CONCEP50 in ('ZP07','ZTAUDI') then im_comcep50 else 0 end  " +
            "as CAV,  " +
            "case when CD_CONCEP1 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep1 else 0 end + " +
            "case when CD_CONCEP2 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep2 else 0 end + " +
            "case when CD_CONCEP3 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep3 else 0 end + " +
            "case when CD_CONCEP4 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep4 else 0 end + " +
            "case when CD_CONCEP5 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep5 else 0 end + " +
            "case when CD_CONCEP6 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep6 else 0 end + " +
            "case when CD_CONCEP7 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep7 else 0 end + " +
            "case when CD_CONCEP8 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep8 else 0 end + " +
            "case when CD_CONCEP9 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep9 else 0 end + " +
            "case when CD_CONCEP10 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep10 else 0 end + " +
            "case when CD_CONCEP11 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep11 else 0 end + " +
            "case when CD_CONCEP12 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep12 else 0 end + " +
            "case when CD_CONCEP13 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep13 else 0 end + " +
            "case when CD_CONCEP14 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep14 else 0 end + " +
            "case when CD_CONCEP15 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep15 else 0 end + " +
            "case when CD_CONCEP16 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep16 else 0 end + " +
            "case when CD_CONCEP17 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep17 else 0 end + " +
            "case when CD_CONCEP18 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep18 else 0 end + " +
            "case when CD_CONCEP19 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep19 else 0 end + " +
            "case when CD_CONCEP20 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep20 else 0 end + " +
            "case when CD_CONCEP21 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep21 else 0 end + " +
            "case when CD_CONCEP22 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep22 else 0 end + " +
            "case when CD_CONCEP23 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep23 else 0 end + " +
            "case when CD_CONCEP24 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep24 else 0 end + " +
            "case when CD_CONCEP25 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep25 else 0 end + " +
            "case when CD_CONCEP26 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep26 else 0 end + " +
            "case when CD_CONCEP27 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep27 else 0 end + " +
            "case when CD_CONCEP28 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep28 else 0 end + " +
            "case when CD_CONCEP29 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep29 else 0 end + " +
            "case when CD_CONCEP30 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep30 else 0 end + " +
            "case when CD_CONCEP31 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep31 else 0 end + " +
            "case when CD_CONCEP32 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep32 else 0 end + " +
            "case when CD_CONCEP33 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep33 else 0 end + " +
            "case when CD_CONCEP34 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep34 else 0 end + " +
            "case when CD_CONCEP35 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep35 else 0 end + " +
            "case when CD_CONCEP36 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep36 else 0 end + " +
            "case when CD_CONCEP37 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep37 else 0 end + " +
            "case when CD_CONCEP38 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep38 else 0 end + " +
            "case when CD_CONCEP39 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep39 else 0 end + " +
            "case when CD_CONCEP40 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep40 else 0 end + " +
            "case when CD_CONCEP41 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep41 else 0 end + " +
            "case when CD_CONCEP42 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep42 else 0 end + " +
            "case when CD_CONCEP43 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep43 else 0 end + " +
            "case when CD_CONCEP44 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep44 else 0 end + " +
            "case when CD_CONCEP45 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep45 else 0 end + " +
            "case when CD_CONCEP46 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep46 else 0 end + " +
            "case when CD_CONCEP47 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep47 else 0 end + " +
            "case when CD_CONCEP48 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep48 else 0 end + " +
            "case when CD_CONCEP49 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep49 else 0 end + " +
            "case when CD_CONCEP50 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then im_comcep50 else 0 end  " +
            "as IMPORTE_REDES,  " +
            "nm_cons_act_tramo1 as CONSUMO_ACTIVA1,nm_cons_act_tramo2 as CONSUMO_ACTIVA2,nm_cons_act_tramo3 as CONSUMO_ACTIVA3,nm_cons_act_tramo4 as CONSUMO_ACTIVA4, " +
            "cd_pun_notif as CUPSREE,nm_vol_calc + nm_num_decimales as POTENCIA " +
            "FROM ed_owner.sap_tfactura_n0 n " +
            "inner join ed_owner.t_ed_h_sap_dc dc " +
            "on n.cd_dc= dc.cd_dc " +
            "LEFT outer JOIN ed_owner.t_ed_h_sap_dc_lineas l " +
            "ON dc.cd_dc= l.cd_dc " +
            "where cd_clase_pos_doc in ('ZP03', 'ZPOBTE','ZPOBTN') " +
            "and n.lg_origen='SAP' " +
            "and cd_territorio1='INTERNACIONAL OFICIN' " +
            "and n.fh_ult_ejec >= '2024-10-01 00:00:00' ";
            #endregion

            return consulta;
        }
        private string Consulta_ERSE_SAP_INDIVIDUALES(DateTime fd, DateTime fh)
        {
            string consulta;

            #region query ERSE SAP INDIVIDUALES
            consulta =
            "select " +
            "dc.cd_empr as CEMPTITU,  " +
            "cf.de_seg_merc_por,	 " +
            "f.cd_est_fact as TESTFACT,  " +
            "f.fh_actual_dtmco as fh_ult_ejec, " +
            "c.cd_nif_cif_cli as CIF,  " +
            "f.id_fact as CFACTURA,  " +
            "f.fh_fact as FFACTURA,  " +
            "f.fh_ini_fact as FFACTDES,   " +
            "f.fh_fin_fact as FFACTHAS,   " +
            "f.im_factdo_con_iva as IFACTURA,  " +
            "f.im_impuesto_1 as IMPUESTO1, " +
            "f.im_impuesto_1 as IMPUESTO2, " +
            "f.im_impuesto_1 as IMPUESTO3, " +
            "f.im_base_imp1, " +
            "f.im_base_imp2, " +
            "f.im_base_imp3, " +
            "(f.im_factdo_sin_iva - f.im_impto_elect) as IBASEISE, " +
            "f.im_impto_elect as ISE,  " +
            "f.im_factdo_sin_iva as BASEIVA, " +
            "case when f.nm_total_iva is null then (f.im_factdo_con_iva-f.im_factdo_sin_iva) else f.nm_total_iva end as IVA, ";
            #region CAV
            for(int cont = 1; cont<50; cont++)
            {
                consulta += $"case when f.cd_concepto_{cont} in ('ZP07','ZTAUDI') then f.im_concepto_{cont} else 0 end + ";
            }
            consulta += "case when f.cd_concepto_50 in ('ZP07','ZTAUDI') then f.im_concepto_50 else 0 end  as CAV, ";
            #endregion

            #region importe_redes
            for (int cont = 1; cont <50; cont++)
            {
                consulta += $"case when f.cd_concepto_{cont} in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then f.im_concepto_{cont} else 0 end + ";
            }
            consulta += "case when f.cd_concepto_50 in ('ZEN_A3','ZEN_A2','ZEN_A1','ZENATR','ZEXATR','ZATBTE','ZEN_T2','ZEN_T3','ZEN_T1','ZENAFV','ZENTFV','ZPO9') then f.im_concepto_50 else 0 end as IMPORTE_REDES, " +
            #endregion
            "case when f.nm_cons_per_1=0 then f.nm_cons_punta else f.nm_cons_per_1 end as CONSUMO_ACTIVA1, " +
            "case when f.nm_cons_per_2=0 then f.nm_cons_llano else f.nm_cons_per_1 end as CONSUMO_ACTIVA2, " +
            "case when f.nm_cons_per_3=0 then f.nm_cons_valle else f.nm_cons_per_1 end as CONSUMO_ACTIVA3, " +
            "case when f.nm_cons_per_4=0 then f.nm_cons_valle_bajo else f.nm_cons_per_1 end as CONSUMO_ACTIVA4, " +
            "f.NM_ENER_CONSMDA as ENERGIA_CONSUMIDA, " +
            "f.NM_ENER_FACTDA as ENERGIA_FACTURADA, " +
            "case when dc.cd_cups_ext is null then dc.cd_cups_gas_ext else dc.cd_cups_ext end as CUPSREE, " +
            "GREATEST(f.nm_pot_max_1, " +
                "f.nm_pot_max_2, " +
                "f.nm_pot_max_3, " +
                "f.nm_pot_max_4, " +
                "f.nm_pot_max_5, " +
                "f.nm_pot_max_6, " +
                "f.nm_pot_max_7, " +
                "f.nm_pot_punta, " +
                "f.nm_pot_llano, " +
                "f.nm_pot_valle) as POTENCIA, " +
                "f.cd_di as DI,   " +
                "f.cd_tp_fact  as TFACTURA,  " +
                "c.tx_apell_cli as DAPERSOC,  	  " +
                "dc.cd_linea_negocio as CSEGMERC,  " +
                "case dc.cd_linea_negocio when '001' then 'L' when '002' then 'G' else 'L' end as TIPONEGOCIO " +
            "from ed_owner.t_ed_h_sap_facts f  " +
            "inner join  " +
                "ed_owner.t_ed_h_sap_dc dc on f.cd_di=dc.cd_di  " +
            "left join  " +
                "ed_owner.t_ed_f_clis c on f.cl_cli = c.cl_cli  " +
            "left join " +
                "(SELECT id_crto_ext, de_seg_merc_por, MAX(cd_sec_crto) AS valor_maximo_version " +
                "FROM ed_owner.t_ed_h_sap_crto_front " +
                "GROUP BY id_crto_ext, de_seg_merc_por) cf on cf.id_crto_ext = dc.id_crto_ext " +
            "where   " +
                //"f.im_factdo_con_iva > 0 and   " +
                "f.fh_ini_fact is not null and " +
                "f.fh_fin_fact is not null and " +
                "f.fh_fact >='" + fd.ToString("yyyy-MM-dd") + "' and   " +
                "f.fh_fact <='" + fh.ToString("yyyy-MM-dd") + "' and   " +
                "f.cl_empr = '006' and dc.cd_empr ='PT1Q' and   " +
                "f.id_fact is not null and " +
                "dc.cd_linea_negocio = '001' " +
            "group by cemptitu,de_seg_merc_por,testfact,fh_ult_ejec,cif,cfactura,ffactura,ffactdes,ffacthas,ifactura,impuesto1,impuesto2,impuesto3,im_base_imp1,im_base_imp2,im_base_imp3,ibaseise,ise,baseiva,iva,cav,importe_redes,consumo_activa1,consumo_activa2,consumo_activa3,consumo_activa4,ENERGIA_CONSUMIDA,ENERGIA_FACTURADA,cupsree,potencia,DI,tfactura,dapersoc,csegmerc,tiponegocio;";
            #endregion

            return consulta;
        }
        private string Consulta_ERSE_SAP_AGRUPADAS()
        {
            string consulta;

            #region query ERSE SAP AGRUPADAS
            consulta = " ";
            #endregion

            return consulta;
        }
    }
}
