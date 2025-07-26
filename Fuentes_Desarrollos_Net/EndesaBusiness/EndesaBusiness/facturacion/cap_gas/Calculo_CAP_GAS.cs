using EndesaBusiness.gas;
using EndesaBusiness.servidores;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;

namespace EndesaBusiness.facturacion.cap_gas
{
    public class Calculo_CAP_GAS
    {

        public bool hayError { get; set; }
        public string descripcion_error { get; set; }


        List<EndesaEntity.facturacion.Factura> lista_facturas;
        
        Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas;
        Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas_psudos;

        Dictionary<string, List<EndesaEntity.facturacion.Factura>> dic_facturas_cups;
        Dictionary<string, List<EndesaEntity.facturacion.Factura>> dic_facturas_pseudos_cups;

        Dictionary<string, string> dic_cups;

        Dictionary<string, string> dic_nif;

        public DateTime fecha_min { get; set; }
        public DateTime fecha_max { get; set; }
        public Calculo_CAP_GAS()
        {
            dic_facturas = new Dictionary<string, EndesaEntity.facturacion.Factura>();
            dic_facturas_psudos = new Dictionary<string, EndesaEntity.facturacion.Factura>();

            dic_facturas_cups = new Dictionary<string, List<EndesaEntity.facturacion.Factura>>();
            dic_nif = new Dictionary<string, string>();

            dic_cups = new Dictionary<string, string>();

            lista_facturas = new List<EndesaEntity.facturacion.Factura>();


        }


        public void ExcelFacturas(string fichero)
        {
            fecha_min = new DateTime();
            fecha_max = new DateTime();
            fecha_min = DateTime.MaxValue;
            fecha_max = DateTime.MinValue;

            hayError = false;
            descripcion_error = "";
            CargaExcel(fichero);

        }


        public void ExcelCUPS(string fichero)
        {
            fecha_min = new DateTime();
            fecha_max = new DateTime();
            fecha_min = DateTime.MaxValue;
            fecha_max = DateTime.MinValue;

            hayError = false;
            descripcion_error = "";            
            CargaExcelCUPS(fichero);

        }

        public void ExcelAgrupadas(string fichero)
        {
            fecha_min = new DateTime();
            fecha_max = new DateTime();
            fecha_min = DateTime.MaxValue;
            fecha_max = DateTime.MinValue;

            hayError = false;
            descripcion_error = "";
            CargaExcelAgrupadas(fichero);

        }

        private void CargaExcel(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";
            string clave = "";

            try
            {
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 100000; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString()
                            + workSheet.Cells[1, 2].Value.ToString()
                            + workSheet.Cells[1, 3].Value.ToString();

                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString()
                        + workSheet.Cells[f, 2].Value.ToString() == "")
                    {
                        break;
                    }
                    else
                    {

                        EndesaEntity.facturacion.Factura x = new EndesaEntity.facturacion.Factura();
                        x.cnifdnic = workSheet.Cells[f, 2].Value.ToString();
                        x.cupsree = workSheet.Cells[f, 5].Value.ToString();
                        x.ffactura = Convert.ToDateTime(workSheet.Cells[f, 11].Value);
                        x.ffactdes = Convert.ToDateTime(workSheet.Cells[f, 12].Value);
                        x.ffacthas = Convert.ToDateTime(workSheet.Cells[f, 13].Value);
                        x.ifactura = Convert.ToDouble(workSheet.Cells[f, 14].Value);
                        x.tfactura_desc = workSheet.Cells[f, 9].Value.ToString();
                        x.comentario = "FACTURA NO ENCONTRADA";

                        fecha_min = fecha_min > x.ffactura ? x.ffactura : fecha_min;
                        fecha_max = fecha_max < x.ffactura ? x.ffactura : fecha_max;


                        if (workSheet.Cells[f, 9].Value.ToString() != "PSEUDOFACTURA")
                        {
                            x.cfactura = workSheet.Cells[f, 6].Value.ToString();                            
                            dic_facturas.Add(x.cfactura, x);
                        }
                        else
                        {
                            // CUPS_FFACTURA_FFACTDES_FFACTHAS_IMPORTE
                            clave = x.cupsree + "_"
                                + x.ffactura.ToString("yyyyMMdd") + "_"
                                + x.ffactdes.ToString("yyyyMMdd") + "_"
                                + x.ffacthas.ToString("yyyyMMdd") + "_"
                                + x.ifactura.ToString().Replace(",", "."); 

                            dic_facturas_psudos.Add(clave, x);

                        }

                        string o;
                        if(!dic_nif.TryGetValue(x.cnifdnic, out o))
                            dic_nif.Add(x.cnifdnic, x.cnifdnic);
                        
                        id++;
                    }

                }


                fs = null;
                excelPackage = null;


                BusquedaFacturas(dic_facturas, dic_facturas_psudos, dic_nif, fecha_min, fecha_max);
                MarcaExcelOrigen(fichero);

            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                         "Error en el formato del fichero",
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
            }
        }

        private void CargaExcelCUPS(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";
            string clave = "";

            try
            {
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 100000; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString()
                            + workSheet.Cells[1, 2].Value.ToString()
                            + workSheet.Cells[1, 3].Value.ToString();

                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString()
                        + workSheet.Cells[f, 2].Value.ToString() == "")
                    {
                        break;
                    }
                    else
                    {

                        
                        EndesaEntity.facturacion.Factura x = new EndesaEntity.facturacion.Factura();
                        //x.cnifdnic = workSheet.Cells[f, 2].Value.ToString();
                        x.cupsree = workSheet.Cells[f, 2].Value.ToString();
                        //x.ffactura = Convert.ToDateTime(workSheet.Cells[f, 11].Value);
                        //x.ffactdes = Convert.ToDateTime(workSheet.Cells[f, 12].Value);
                        //x.ffacthas = Convert.ToDateTime(workSheet.Cells[f, 13].Value);
                        //x.ifactura = Convert.ToDouble(workSheet.Cells[f, 14].Value);
                        //x.tfactura_desc = workSheet.Cells[f, 9].Value.ToString();
                        x.comentario = "FACTURA NO ENCONTRADA";
                        

                        dic_cups.Add(x.cupsree, x.cupsree);
                        
                                               

                        id++;
                    }

                }


                fs = null;
                excelPackage = null;


                BusquedaFacturasCUPS(dic_facturas);
                //MarcaExcelOrigen(fichero);

            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                         "Error en el formato del fichero",
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
            }
        }

        private void CargaExcelAgrupadas(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";
            string clave = "";

            try
            {
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 100000; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        cabecera = workSheet.Cells[1, 1].Value.ToString();
                           

                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;                  
                    else
                    {


                        EndesaEntity.facturacion.Factura x = new EndesaEntity.facturacion.Factura();
                        //x.cnifdnic = workSheet.Cells[f, 2].Value.ToString();
                        //x.cupsree = workSheet.Cells[f, 2].Value.ToString();
                        //x.ffactura = Convert.ToDateTime(workSheet.Cells[f, 11].Value);
                        //x.ffactdes = Convert.ToDateTime(workSheet.Cells[f, 12].Value);
                        //x.ffacthas = Convert.ToDateTime(workSheet.Cells[f, 13].Value);
                        //x.ifactura = Convert.ToDouble(workSheet.Cells[f, 14].Value);
                        //x.tfactura_desc = workSheet.Cells[f, 9].Value.ToString();
                        x.cfactura = workSheet.Cells[f, 1].Value.ToString();
                        x.comentario = "FACTURA NO ENCONTRADA";

                        dic_cups.Add(x.cfactura, x.cfactura);



                        id++;
                    }

                }


                fs = null;
                excelPackage = null;


                BusquedaFacturasAgrupadas(dic_facturas);
                //MarcaExcelOrigen(fichero);

            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                         "Error en el formato del fichero",
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
            }
        }

        private void MarcaExcelOrigen(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";
            string clave = "";

            EndesaEntity.facturacion.Factura o;

            string ruta_plantilla_facturas = fichero;
            FileInfo origen_facturas = new FileInfo(ruta_plantilla_facturas);
            FileInfo nombreSalidaExcel = new FileInfo(@"C:\Temp\" + 
                origen_facturas.Name.Replace(".xlsx", "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"));

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(origen_facturas);
            var workSheet = excelPackage.Workbook.Worksheets.First();
            

            f = 1; // Porque la primera fila es la cabecera
            for (int i = 0; i < 100000; i++)
            {
                c = 1;

                if (firstOnly)
                {
                    cabecera = workSheet.Cells[1, 1].Value.ToString()
                        + workSheet.Cells[1, 2].Value.ToString()
                        + workSheet.Cells[1, 3].Value.ToString();

                    if (!EstructuraCorrecta(cabecera))
                    {
                        this.hayError = true;
                        this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                        break;
                    }

                    firstOnly = false;
                }

                f++;

                if (workSheet.Cells[f, 1].Value == null)
                    break;

                if (workSheet.Cells[f, 1].Value.ToString()
                    + workSheet.Cells[f, 2].Value.ToString() == "")
                {
                    break;
                }
                else
                {

                    EndesaEntity.facturacion.Factura x = new EndesaEntity.facturacion.Factura();
                    x.cnifdnic = workSheet.Cells[f, 2].Value.ToString();
                    x.cupsree = workSheet.Cells[f, 5].Value.ToString();
                    x.ffactura = Convert.ToDateTime(workSheet.Cells[f, 11].Value);
                    x.ffactdes = Convert.ToDateTime(workSheet.Cells[f, 12].Value);
                    x.ffacthas = Convert.ToDateTime(workSheet.Cells[f, 13].Value);
                    x.ifactura = Convert.ToDouble(workSheet.Cells[f, 14].Value);
                    x.tfactura_desc = workSheet.Cells[f, 9].Value.ToString();                    

                    fecha_min = fecha_min > x.ffactura ? x.ffactura : fecha_min;
                    fecha_max = fecha_max < x.ffactura ? x.ffactura : fecha_max;


                    if (workSheet.Cells[f, 9].Value.ToString() != "PSEUDOFACTURA")
                    {
                        x.cfactura = workSheet.Cells[f, 6].Value.ToString();


                        if (dic_facturas.TryGetValue(x.cfactura, out o))
                            if (o.comentario == "FACTURA NO ENCONTRADA")
                                workSheet.Cells[f, 25].Value = o.comentario;

                    }
                    else
                    {
                        // CUPS_FFACTURA_FFACTDES_FFACTHAS_IMPORTE
                        clave = x.cupsree + "_"
                            + x.ffactura.ToString("yyyyMMdd") + "_"
                            + x.ffactdes.ToString("yyyyMMdd") + "_"
                            + x.ffacthas.ToString("yyyyMMdd") + "_"
                            + x.ifactura.ToString().Replace(",", ".");
                        
                        if (dic_facturas_psudos.TryGetValue(clave, out o))
                            if (o.comentario == "FACTURA NO ENCONTRADA")
                                workSheet.Cells[f, 25].Value = o.comentario;

                    }

                    

                    id++;
                }

            }

            excelPackage.SaveAs(nombreSalidaExcel);
            excelPackage = null;
        }

        private bool EstructuraCorrecta(string cabecera)
        {

            return true;
          
        }


        private void BusquedaFacturas(Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas,
            Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas_pseudos,
            Dictionary<string, string> dic_nif, 
            DateTime fd, DateTime fh)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;
            EndesaEntity.facturacion.Factura factura;
            EndesaEntity.facturacion.Factura_Detalle factura_detalle;
            string clave = "";

            int c = 0;
            int f = 0;

            string error_fila = "";


            strSql = "SELECT  e.descripcion AS entorno, f.CNIFDNIC, f.CUPSREE, f.CFACTURA,"
                + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.CREFEREN, f.SECFACTU, f.CUPSREE,"
                + " f.IVA, f.ISE, f.IBASEISE, f.VCUOVAFA, t.TCONFAC, t.ICONFAC, f.IFACTURA,"
                + " f.TFACTURA, f.CCOUNIPS, f.IIMPUES2, f.IIMPUES3, ag.CFACAGP"
                + " FROM fo f"
                + " INNER JOIN fo_empresas e on"
                + " e.empresa_id = f.ID_ENTORNO"
                + " LEFT OUTER JOIN fo_tcon t ON"
                + " t.CREFEREN = f.CREFEREN AND"
                + " t.SECFACTU = f.SECFACTU AND"
                + " t.TESTFACT = f.TESTFACT AND"
                + " t.TCONFAC IN (1222,1256,5162,660,1261)"
                + " LEFT OUTER JOIN fo_agrupadas ag on"
                + " ag.CREFEREN = f.CREFEREN AND"
                + " ag.SECFACTU = f.SECFACTU AND"
                + " ag.TESTFACT = f.TESTFACT"
                + " WHERE f.CNIFDNIC in ";

            foreach(KeyValuePair<string, string> p in dic_nif)
            {
                if (firstOnly)
                {
                    strSql += "('" + p.Key + "'";
                    firstOnly = false;
                }
                else
                    strSql += ",'" + p.Key + "'";

            }

            strSql += ") AND f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "' AND"
                + " f.TFACTURA IN (1,2,3,8,9) AND t.TESTFACT IN ('N','Y','S')";
                //+ " AND f.CFACTURA = 'P0Z202N0157733'";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["CFACTURA"] != System.DBNull.Value)
                {
                    if (r["CFACTURA"].ToString().Trim() != "")
                    {

                        if (dic_facturas.TryGetValue(r["CFACTURA"].ToString(), out factura))
                        {

                            factura.comentario = "FACTURA ENCONTRADA";

                            factura.entorno = r["ENTORNO"].ToString();

                            if (r["CUPSREE"] != System.DBNull.Value)
                                factura.cupsree = r["CUPSREE"].ToString();

                            if (r["CCOUNIPS"] != System.DBNull.Value)
                                factura.ccounips = r["CCOUNIPS"].ToString();

                            if (r["CREFEREN"] != System.DBNull.Value)
                                factura.creferen = Convert.ToInt64(r["CREFEREN"]);

                            if (r["SECFACTU"] != System.DBNull.Value)
                                factura.secfactu = Convert.ToInt32(r["SECFACTU"]);

                            if (r["IIMPUES2"] != System.DBNull.Value)
                                factura.iimpues2 = Convert.ToDouble(r["IIMPUES2"]);

                            if (r["IIMPUES3"] != System.DBNull.Value)
                                factura.iimpues3 = Convert.ToDouble(r["IIMPUES3"]);

                            if (r["IVA"] != System.DBNull.Value)
                                factura.iva = Convert.ToDouble(r["IVA"]);

                            if (r["ISE"] != System.DBNull.Value)
                                factura.ise = Convert.ToDouble(r["ISE"]);

                            if (r["IBASEISE"] != System.DBNull.Value)
                                factura.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                            if (r["VCUOVAFA"] != System.DBNull.Value)
                                factura.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                            if (r["CFACAGP"] != System.DBNull.Value)
                                factura.cfacagp = r["CFACAGP"].ToString();


                            if (r["TCONFAC"] != System.DBNull.Value)
                            {
                                factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                                factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                                factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                                factura.lista_conceptos.Add(factura_detalle);
                            }





                        }
                    }
                    else
                    {
                        // CUPS_FFACTURA_FFACTDES_FFACTHAS_IMPORTE
                        //clave = x.cupsree + "_"
                        //    + x.ffactura.ToString("yyyyMMdd") + "_"
                        //    + x.ffactdes.ToString("yyyyMMdd") + "_"
                        //    + x.ffacthas.ToString("yyyyMMdd") + "_"
                        //    + x.ifactura.ToString();


                        clave = r["CUPSREE"].ToString() + "_"
                            + Convert.ToDateTime(r["FFACTURA"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDateTime(r["FFACTDES"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDateTime(r["FFACTHAS"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDouble(r["IFACTURA"]).ToString().Replace(",", ".");

                        if (dic_facturas_pseudos.TryGetValue(clave, out factura))
                        {
                            factura.comentario = "FACTURA ENCONTRADA";

                            factura.entorno = r["ENTORNO"].ToString();

                            if (r["CUPSREE"] != System.DBNull.Value)
                                factura.cupsree = r["CUPSREE"].ToString();

                            if (r["CCOUNIPS"] != System.DBNull.Value)
                                factura.ccounips = r["CCOUNIPS"].ToString();

                            if (r["CREFEREN"] != System.DBNull.Value)
                                factura.creferen = Convert.ToInt64(r["CREFEREN"]);

                            if (r["SECFACTU"] != System.DBNull.Value)
                                factura.secfactu = Convert.ToInt32(r["SECFACTU"]);

                            if (r["IVA"] != System.DBNull.Value)
                                factura.iva = Convert.ToDouble(r["IVA"]);

                            if (r["IIMPUES2"] != System.DBNull.Value)
                                factura.iimpues2 = Convert.ToDouble(r["IIMPUES2"]);

                            if (r["IIMPUES3"] != System.DBNull.Value)
                                factura.iimpues3 = Convert.ToDouble(r["IIMPUES3"]);

                            if (r["ISE"] != System.DBNull.Value)
                                factura.ise = Convert.ToDouble(r["ISE"]);

                            if (r["IBASEISE"] != System.DBNull.Value)
                                factura.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                            if (r["VCUOVAFA"] != System.DBNull.Value)
                                factura.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                            if (r["TCONFAC"] != System.DBNull.Value)
                            {
                                factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                                factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                                factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                                factura.lista_conceptos.Add(factura_detalle);
                            }
                        }
                    }
                }
                

            }
            db.CloseConnection();


            // Pintamos los datos en las facturas

            string ruta_plantilla_cim = System.Environment.CurrentDirectory + @"\media\SEPARACION CAP DE GAS - FACTURA (CIM).xlsx";
            string ruta_plantilla_ie = System.Environment.CurrentDirectory + @"\media\SEPARACION CAP DE GAS - FACTURA.xlsx";


            FileInfo nombreSalidaExcel_CIM = new FileInfo(@"C:\Temp\SEPARACION CAP DE GAS - FACTURA(CIM)_"
                   + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            FileInfo nombreSalidaExcel = new FileInfo(@"C:\Temp\SEPARACION CAP DE GAS - FACTURA_"
                   + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            FileInfo plantillaExcel = new FileInfo(ruta_plantilla_cim);

            #region calculos-cim

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
            var workSheet = excelPackage.Workbook.Worksheets["CALCULOS - CIM"];

            f = 5;


            workSheet = excelPackage.Workbook.Worksheets["CALCULOS - CIM"];

            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas)
            {
                if (p.Value.comentario == "FACTURA ENCONTRADA")
                {
                                       

                    if (p.Value.ise == 0)
                    {
                        error_fila = "";
                        f++;
                        c = 2;

                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;

                        workSheet.Cells[f, c].Value = p.Value.ifactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        if (p.Value.iva != 0)
                            workSheet.Cells[f, c].Value = p.Value.iva;
                        else if (p.Value.iimpues2 != 0)
                            workSheet.Cells[f, c].Value = p.Value.iimpues2;
                        else
                            workSheet.Cells[f, c].Value = p.Value.iimpues3;

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        workSheet.Cells[f, c].Value = p.Value.ise;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;


                        if(p.Value.ibaseise != 0)                         
                            workSheet.Cells[f, c].Value = p.Value.ibaseise;                    
                        else
                        {
                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                            if (factura_detalle == null)
                                factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 650);

                            if (factura_detalle != null)
                                workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            else
                                workSheet.Cells[f, c].Value = 0;
                        }

                       

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; 
                        c++;

                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 5162);
                        if (factura_detalle != null)
                        {
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                        }

                        c++;

                        workSheet.Cells[f, c].Value = p.Value.vcuovafa;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;


                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222);

                        if (factura_detalle == null)
                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1256);

                        if (factura_detalle != null)
                        {
                            workSheet.Cells[f, 14].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, 14].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                            error_fila = error_fila == "" ? "No tiene concepto 1222 ó 1256" : error_fila + " ,1222 ó 1256";

                        workSheet.Cells[f, 1].Value = error_fila;
                        c++;


                        // Ponemos el resto de campos a partir de la columna 24
                        c = 27;
                        workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                        workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                        workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                        workSheet.Cells[f, c].Value = p.Value.secfactu; c++;

                        if (p.Value.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffactdes > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactdes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffacthas > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffacthas;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }

                        c++;

                        workSheet.Cells[f, c].Value = p.Value.entorno; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfacagp; c++;


                    }
                }

               
            }


            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas_pseudos)
            {
                if (p.Value.comentario == "FACTURA ENCONTRADA")
                {
                    
                    if (p.Value.ise == 0)
                    {

                        error_fila = "";
                        f++;
                        c = 2;

                        if (p.Value.cfactura != null)
                            workSheet.Cells[f, c].Value = p.Value.cfactura;
                        c++;

                        workSheet.Cells[f, c].Value = p.Value.ifactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        if (p.Value.iva != 0)                        
                            workSheet.Cells[f, c].Value = p.Value.iva;
                        else if (p.Value.iimpues2 != 0)
                            workSheet.Cells[f, c].Value = p.Value.iimpues2;
                        else
                            workSheet.Cells[f, c].Value = p.Value.iimpues3;


                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        workSheet.Cells[f, c].Value = p.Value.ise;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        if (p.Value.ibaseise != 0)
                            workSheet.Cells[f, c].Value = p.Value.ibaseise;
                        else
                        {
                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                            if (factura_detalle == null)
                                factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 650);

                            if (factura_detalle != null)
                                workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            else
                                workSheet.Cells[f, c].Value = 0;
                        }

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; 
                        c++;

                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 5162);
                        if(factura_detalle != null)
                        {
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                        }
                        
                         c++;


                        workSheet.Cells[f, c].Value = p.Value.vcuovafa;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222 || z.tconfac == 1256);

                        if (factura_detalle != null)
                        {
                            workSheet.Cells[f, 14].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, 14].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                            error_fila = error_fila == "" ? "No tiene concepto 1222 ó 1256" : error_fila + " ,1222 ó 1256";

                        workSheet.Cells[f, 1].Value = error_fila;

                        c++;

                        // Ponemos el resto de campos a partir de la columna 24
                        c = 27;
                        workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                        workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                        workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                        workSheet.Cells[f, c].Value = p.Value.secfactu; c++;

                        if (p.Value.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffactdes > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactdes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffacthas > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffacthas;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.Value.entorno; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfacagp; c++;
                    }
                }


            }

            excelPackage.SaveAs(nombreSalidaExcel_CIM);
            excelPackage = null;
            #endregion

            plantillaExcel = new FileInfo(ruta_plantilla_ie);


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            excelPackage = new ExcelPackage(plantillaExcel);
            workSheet = excelPackage.Workbook.Worksheets["CALCULOS"];

            f = 4;

            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas)
            {

                if (p.Value.comentario == "FACTURA ENCONTRADA")
                {

                    if (p.Value.ise != 0)
                    {
                        error_fila = "";
                        f++;
                        c = 2;


                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;

                        workSheet.Cells[f, c].Value = p.Value.ifactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        if (p.Value.iva != 0)
                            workSheet.Cells[f, c].Value = p.Value.iva;
                        else if (p.Value.iimpues2 != 0)
                            workSheet.Cells[f, c].Value = p.Value.iimpues2;
                        else
                            workSheet.Cells[f, c].Value = p.Value.iimpues3;

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        // ISE / TCONFAC 660                                      

                        workSheet.Cells[f, c].Value = p.Value.ise;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                        c++;


                        // BASE ISE FACTURA (IBASEISE) / TCONFAC 1261
                        if (p.Value.ibaseise != 0)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ibaseise;

                        }
                        else
                        {

                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                            if (factura_detalle != null)
                            {
                                workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            }
                            else
                                error_fila = error_fila == "" ? "No tiene concepto 1261" : error_fila + " ,1261";
                        }

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222 || z.tconfac == 1256);

                        if (factura_detalle != null)
                        {
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                            error_fila = error_fila == "" ? "No tiene concepto 1222 ó 1256" : error_fila + " ,1222 ó 1256";

                        workSheet.Cells[f, 1].Value = error_fila;
                        c++;

                        // Ponemos el resto de campos a partir de la columna 24
                        c = 25;
                        workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                        workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                        workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                        workSheet.Cells[f, c].Value = p.Value.secfactu; c++;

                        if (p.Value.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffactdes > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactdes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffacthas > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffacthas;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.Value.entorno; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfacagp; c++;


                    }

                }
            }


            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas_pseudos)
            {

                if (p.Value.comentario == "FACTURA ENCONTRADA")
                {

                    if (p.Value.ise != 0)
                    {

                        error_fila = "";
                        f++;
                        c = 2;

                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;

                        workSheet.Cells[f, c].Value = p.Value.ifactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        if (p.Value.iva != 0)
                            workSheet.Cells[f, c].Value = p.Value.iva;
                        else if (p.Value.iimpues2 != 0)
                            workSheet.Cells[f, c].Value = p.Value.iimpues2;
                        else
                            workSheet.Cells[f, c].Value = p.Value.iimpues3;

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                        workSheet.Cells[f, c].Value = p.Value.ise;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;


                        // ISE / TCONFAC 660
                        //factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 660);

                        //if (factura_detalle != null)
                        //{
                        //    workSheet.Cells[f, c].Value = p.Value.ise / factura_detalle.iconfac;
                        //    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        //}
                        //else
                        //    error_fila = "No tiene concepto 660";

                        // BASE ISE FACTURA (IBASEISE) / TCONFAC 1261

                        if (p.Value.ibaseise != 0)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ibaseise;

                        }
                        else
                        {

                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                            if (factura_detalle != null)
                            {
                                workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                            }
                            else
                                error_fila = error_fila == "" ? "No tiene concepto 1261" : error_fila + " ,1261";
                        }

                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;


                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222);
                        if (factura_detalle == null)
                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1256);

                        if (factura_detalle != null)
                        {
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                            error_fila = error_fila == "" ? "No tiene concepto 1222 ó 1256" : error_fila + " ,1222 ó 1256";
                        c++;

                        workSheet.Cells[f, 1].Value = error_fila;


                        // Ponemos el resto de campos a partir de la columna 24
                        c = 25;
                        workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                        workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                        workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                        workSheet.Cells[f, c].Value = p.Value.secfactu; c++;

                        if (p.Value.ffactura > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactura;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffactdes > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffactdes;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        if (p.Value.ffacthas > DateTime.MinValue)
                        {
                            workSheet.Cells[f, c].Value = p.Value.ffacthas;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        c++;

                        workSheet.Cells[f, c].Value = p.Value.entorno; c++;
                        workSheet.Cells[f, c].Value = p.Value.cfacagp; c++;
                    }
                  


                }


            }

            excelPackage.SaveAs(nombreSalidaExcel);
            excelPackage = null;

        }

        private void BusquedaFacturasCUPS(Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas)           
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;
            EndesaEntity.facturacion.Factura factura;
            EndesaEntity.facturacion.Factura_Detalle factura_detalle;
            string clave = "";

            int c = 0;
            int f = 0;

            string error_fila = "";


            strSql = "SELECT  e.descripcion AS entorno, f.CNIFDNIC, f.DAPERSOC, f.CUPSREE, f.CFACTURA,"
                + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.CREFEREN, f.SECFACTU, f.CUPSREE,"
                + " f.IVA, f.ISE, f.IBASEISE, f.VCUOVAFA, t.TCONFAC, t.ICONFAC, f.IFACTURA,"
                + " f.TFACTURA, f.CCOUNIPS, f.IIMPUES2, f.IIMPUES3, f.IFACTURA, ag.CFACAGP"
                + " FROM fo f"
                + " INNER JOIN fo_empresas e on"
                + " e.empresa_id = f.ID_ENTORNO"
                + " LEFT OUTER JOIN fo_tcon t ON"
                + " t.CREFEREN = f.CREFEREN AND"
                + " t.SECFACTU = f.SECFACTU AND"
                + " t.TESTFACT = f.TESTFACT AND"
                + " t.TCONFAC IN (1222,1256,5162,660,1261)"
                + " LEFT OUTER JOIN fo_agrupadas_aux ag on"                
                + " ag.CREFEREN = f.CREFEREN AND"
                + " ag.SECFACTU = f.SECFACTU AND"
                + " ag.TESTFACT = f.TESTFACT"
                + " WHERE f.CUPSREE in ";

            foreach (KeyValuePair<string, string> p in dic_cups)
            {
                if (firstOnly)
                {
                    strSql += "('" + p.Key + "'";
                    firstOnly = false;
                }
                else
                    strSql += ",'" + p.Key + "'";

            }

            strSql += ") AND f.FFACTDES >= '2022-06-15'";
                
                
            //+ " AND f.CFACTURA = 'P0Z202N0157733'";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                
                

                factura = new EndesaEntity.facturacion.Factura();                    

                factura.comentario = "FACTURA ENCONTRADA";

                factura.entorno = r["ENTORNO"].ToString();

                if (r["CNIFDNIC"] != System.DBNull.Value)
                    factura.cnifdnic = r["CNIFDNIC"].ToString();

                if (r["DAPERSOC"] != System.DBNull.Value)
                    factura.dapersoc = r["DAPERSOC"].ToString();

                if (r["CFACTURA"] != System.DBNull.Value)
                    factura.cfactura = r["CFACTURA"].ToString();

                if (r["FFACTURA"] != System.DBNull.Value)
                    factura.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                if (r["FFACTDES"] != System.DBNull.Value)
                    factura.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                if (r["FFACTHAS"] != System.DBNull.Value)
                    factura.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                if (r["CUPSREE"] != System.DBNull.Value)
                    factura.cupsree = r["CUPSREE"].ToString();

                if (r["CCOUNIPS"] != System.DBNull.Value)
                    factura.ccounips = r["CCOUNIPS"].ToString();

                if (r["CREFEREN"] != System.DBNull.Value)
                    factura.creferen = Convert.ToInt64(r["CREFEREN"]);

                if (r["SECFACTU"] != System.DBNull.Value)
                    factura.secfactu = Convert.ToInt32(r["SECFACTU"]);

                if (r["IIMPUES2"] != System.DBNull.Value)
                    factura.iimpues2 = Convert.ToDouble(r["IIMPUES2"]);

                if (r["IIMPUES3"] != System.DBNull.Value)
                    factura.iimpues3 = Convert.ToDouble(r["IIMPUES3"]);

                if (r["IVA"] != System.DBNull.Value)
                    factura.iva = Convert.ToDouble(r["IVA"]);

                if (r["ISE"] != System.DBNull.Value)
                    factura.ise = Convert.ToDouble(r["ISE"]);

                if (r["IBASEISE"] != System.DBNull.Value)
                    factura.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                if (r["IFACTURA"] != System.DBNull.Value)
                    factura.ifactura = Convert.ToDouble(r["IFACTURA"]);

                if (r["VCUOVAFA"] != System.DBNull.Value)
                    factura.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

                if (r["CFACAGP"] != System.DBNull.Value)
                    factura.cfacagp = r["CFACAGP"].ToString();

                if (factura.cfactura.Trim() != "")
                {
                    EndesaEntity.facturacion.Factura o;
                    if (!dic_facturas.TryGetValue(factura.cfactura, out o))
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            factura.lista_conceptos.Add(factura_detalle);
                        }

                        dic_facturas.Add(factura.cfactura, factura);

                    }
                    else
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            o.lista_conceptos.Add(factura_detalle);
                        }
                    }
                }                    
                else // Caso que no tenemos CFACTURA
                {
                    clave = r["CUPSREE"].ToString() + "_"
                            + Convert.ToDateTime(r["FFACTURA"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDateTime(r["FFACTDES"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDateTime(r["FFACTHAS"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDouble(r["IFACTURA"]).ToString().Replace(",", ".");


                    EndesaEntity.facturacion.Factura o;
                    if (!dic_facturas.TryGetValue(clave, out o))
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            factura.lista_conceptos.Add(factura_detalle);
                        }

                        dic_facturas.Add(clave, factura);

                    }
                    else
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            o.lista_conceptos.Add(factura_detalle);
                        }
                    }

                }


            }
            db.CloseConnection();


            // Pintamos los datos en las facturas

            string ruta_plantilla_cim = System.Environment.CurrentDirectory + @"\media\SEPARACION CAP DE GAS - FACTURA (CIM).xlsx";
            string ruta_plantilla_ie = System.Environment.CurrentDirectory + @"\media\SEPARACION CAP DE GAS - FACTURA.xlsx";


            FileInfo nombreSalidaExcel_CIM = new FileInfo(@"C:\Temp\SEPARACION CAP DE GAS - FACTURA(CIM)_"
                   + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            FileInfo nombreSalidaExcel = new FileInfo(@"C:\Temp\SEPARACION CAP DE GAS - FACTURA_"
                   + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            FileInfo plantillaExcel = new FileInfo(ruta_plantilla_cim);

            #region calculos-cim

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
            var workSheet = excelPackage.Workbook.Worksheets["CALCULOS - CIM"];

            f = 5;


            workSheet = excelPackage.Workbook.Worksheets["CALCULOS - CIM"];

            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas)
            {
                
                

                if (p.Value.ise == 0)
                {
                    error_fila = "";
                    f++;
                    c = 2;

                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;

                    workSheet.Cells[f, c].Value = p.Value.ifactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (p.Value.iva != 0)
                        workSheet.Cells[f, c].Value = p.Value.iva;
                    else if (p.Value.iimpues2 != 0)
                        workSheet.Cells[f, c].Value = p.Value.iimpues2;
                    else
                        workSheet.Cells[f, c].Value = p.Value.iimpues3;

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    workSheet.Cells[f, c].Value = p.Value.ise;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;


                    if (p.Value.ibaseise != 0)
                        workSheet.Cells[f, c].Value = p.Value.ibaseise;
                    else
                    {
                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                        if (factura_detalle == null)
                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 650);

                        if (factura_detalle != null)
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                        else
                            workSheet.Cells[f, c].Value = 0;
                    }



                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    c++;

                    factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 5162);
                    if (factura_detalle != null)
                    {
                        workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                    }

                    c++;

                    workSheet.Cells[f, c].Value = p.Value.vcuovafa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;


                    factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222);

                    if (factura_detalle == null)
                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1256);

                    if (factura_detalle != null)
                    {
                        workSheet.Cells[f, 14].Value = factura_detalle.iconfac;
                        workSheet.Cells[f, 14].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                        error_fila = error_fila == "" ? "No tiene concepto 1222 ó 1256" : error_fila + " ,1222 ó 1256";

                    workSheet.Cells[f, 1].Value = error_fila;
                    c++;


                    // Ponemos el resto de campos a partir de la columna 26
                    c = 26;
                    workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                    workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                    workSheet.Cells[f, c].Value = p.Value.secfactu; c++;

                    if (p.Value.ffactura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value.ffactdes > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffactdes;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value.ffacthas > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffacthas;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }

                    c++;

                    workSheet.Cells[f, c].Value = p.Value.entorno; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfacagp; c++;
                    workSheet.Cells[f, c].Value = p.Value.cnifdnic; c++;
                    workSheet.Cells[f, c].Value = p.Value.dapersoc; c++;

                }


            }


            

            excelPackage.SaveAs(nombreSalidaExcel_CIM);
            excelPackage = null;
            #endregion

            plantillaExcel = new FileInfo(ruta_plantilla_ie);


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            excelPackage = new ExcelPackage(plantillaExcel);
            workSheet = excelPackage.Workbook.Worksheets["CALCULOS"];

            f = 4;

            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas)
            {

                

                if (p.Value.ise != 0)
                {
                    error_fila = "";
                    f++;
                    c = 2;


                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;

                    workSheet.Cells[f, c].Value = p.Value.ifactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (p.Value.iva != 0)
                        workSheet.Cells[f, c].Value = p.Value.iva;
                    else if (p.Value.iimpues2 != 0)
                        workSheet.Cells[f, c].Value = p.Value.iimpues2;
                    else
                        workSheet.Cells[f, c].Value = p.Value.iimpues3;

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    // ISE / TCONFAC 660                                      

                    workSheet.Cells[f, c].Value = p.Value.ise;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";

                    c++;


                    // BASE ISE FACTURA (IBASEISE) / TCONFAC 1261
                    if (p.Value.ibaseise != 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ibaseise;

                    }
                    else
                    {

                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                        if (factura_detalle != null)
                        {
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                            workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                        }
                        else
                            error_fila = error_fila == "" ? "No tiene concepto 1261" : error_fila + " ,1261";
                    }

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222 || z.tconfac == 1256);

                    if (factura_detalle != null)
                    {
                        workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                        error_fila = error_fila == "" ? "No tiene concepto 1222 ó 1256" : error_fila + " ,1222 ó 1256";

                    workSheet.Cells[f, 1].Value = error_fila;
                    c++;

                    // Ponemos el resto de campos a partir de la columna 26
                    c = 26;
                    workSheet.Cells[f, c].Value = p.Value.ccounips; c++;
                    workSheet.Cells[f, c].Value = p.Value.cupsree; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                    workSheet.Cells[f, c].Value = p.Value.secfactu; c++;

                    if (p.Value.ffactura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value.ffactdes > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffactdes;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (p.Value.ffacthas > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffacthas;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.Value.entorno; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfacagp; c++;
                    workSheet.Cells[f, c].Value = p.Value.cnifdnic; c++;
                    workSheet.Cells[f, c].Value = p.Value.dapersoc; c++;
                    workSheet.Cells[f, c].Value = p.Value.vcuovafa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";




                }
                    
                
            }


            

            excelPackage.SaveAs(nombreSalidaExcel);
            excelPackage = null;

        }

        private void BusquedaFacturasAgrupadas(Dictionary<string, EndesaEntity.facturacion.Factura> dic_facturas)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;
            EndesaEntity.facturacion.Factura factura;
            EndesaEntity.facturacion.Factura_Detalle factura_detalle;
            string clave = "";

            int c = 0;
            int f = 0;

            string error_fila = "";


            strSql = "SELECT  e.descripcion AS entorno, f.CNIFDNIC, f.DAPERSOC, f.CUPSREE, f.CFACTURA,"
                + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.CREFEREN, f.SECFACTU, f.TESTFACT,  f.CUPSREE,"
                + " f.IVA, f.ISE, f.IBASEISE, f.VCUOVAFA, t.TCONFAC, t.ICONFAC, f.IFACTURA,"
                + " f.TFACTURA, f.CCOUNIPS, f.IIMPUES2, f.IIMPUES3, f.IFACTURA"
                + " FROM fo f"
                + " INNER JOIN fo_empresas e on"
                + " e.empresa_id = f.ID_ENTORNO"
                + " LEFT OUTER JOIN fo_tcon t ON"
                + " t.CREFEREN = f.CREFEREN AND"
                + " t.SECFACTU = f.SECFACTU AND"
                + " t.TESTFACT = f.TESTFACT AND"
                + " t.TCONFAC IN (1222,1256,5162,5163, 660,1261)"               
                + " WHERE f.CFACTURA in ";

            foreach (KeyValuePair<string, string> p in dic_cups)
            {
                if (firstOnly)
                {
                    strSql += "('" + p.Key + "'";
                    firstOnly = false;
                }
                else
                    strSql += ",'" + p.Key + "'";

            }

            strSql += ")";
                        
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {



                factura = new EndesaEntity.facturacion.Factura();

                factura.comentario = "FACTURA ENCONTRADA";

                factura.entorno = r["ENTORNO"].ToString();

                if (r["CNIFDNIC"] != System.DBNull.Value)
                    factura.cnifdnic = r["CNIFDNIC"].ToString();

                if (r["TFACTURA"] != System.DBNull.Value)
                    factura.tfactura = Convert.ToInt32(r["TFACTURA"]);

                if (r["DAPERSOC"] != System.DBNull.Value)
                    factura.dapersoc = r["DAPERSOC"].ToString();

                if (r["CFACTURA"] != System.DBNull.Value)
                    factura.cfactura = r["CFACTURA"].ToString();

                if (r["TESTFACT"] != System.DBNull.Value)
                    factura.testfact = r["TESTFACT"].ToString();

                if (r["FFACTURA"] != System.DBNull.Value)
                    factura.ffactura = Convert.ToDateTime(r["FFACTURA"]);

                if (r["FFACTDES"] != System.DBNull.Value)
                    factura.ffactdes = Convert.ToDateTime(r["FFACTDES"]);

                if (r["FFACTHAS"] != System.DBNull.Value)
                    factura.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);

                if (r["CUPSREE"] != System.DBNull.Value)
                    factura.cupsree = r["CUPSREE"].ToString();

                if (r["CCOUNIPS"] != System.DBNull.Value)
                    factura.ccounips = r["CCOUNIPS"].ToString();

                if (r["CREFEREN"] != System.DBNull.Value)
                    factura.creferen = Convert.ToInt64(r["CREFEREN"]);

                if (r["SECFACTU"] != System.DBNull.Value)
                    factura.secfactu = Convert.ToInt32(r["SECFACTU"]);

                if (r["IIMPUES2"] != System.DBNull.Value)
                    factura.iimpues2 = Convert.ToDouble(r["IIMPUES2"]);

                if (r["IIMPUES3"] != System.DBNull.Value)
                    factura.iimpues3 = Convert.ToDouble(r["IIMPUES3"]);

                if (r["IVA"] != System.DBNull.Value)
                    factura.iva = Convert.ToDouble(r["IVA"]);

                if (r["ISE"] != System.DBNull.Value)
                    factura.ise = Convert.ToDouble(r["ISE"]);

                if (r["IBASEISE"] != System.DBNull.Value)
                    factura.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                if (r["IFACTURA"] != System.DBNull.Value)
                    factura.ifactura = Convert.ToDouble(r["IFACTURA"]);

                if (r["VCUOVAFA"] != System.DBNull.Value)
                    factura.vcuovafa = Convert.ToInt32(r["VCUOVAFA"]);

               

                if (factura.cfactura.Trim() != "")
                {
                    EndesaEntity.facturacion.Factura o;
                    if (!dic_facturas.TryGetValue(factura.cfactura, out o))
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            factura.lista_conceptos.Add(factura_detalle);
                        }

                        dic_facturas.Add(factura.cfactura, factura);

                    }
                    else
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            o.lista_conceptos.Add(factura_detalle);
                        }
                    }
                }
                else // Caso que no tenemos CFACTURA
                {
                    clave = r["CUPSREE"].ToString() + "_"
                            + Convert.ToDateTime(r["FFACTURA"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDateTime(r["FFACTDES"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDateTime(r["FFACTHAS"]).ToString("yyyyMMdd") + "_"
                            + Convert.ToDouble(r["IFACTURA"]).ToString().Replace(",", ".");


                    EndesaEntity.facturacion.Factura o;
                    if (!dic_facturas.TryGetValue(clave, out o))
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            factura.lista_conceptos.Add(factura_detalle);
                        }

                        dic_facturas.Add(clave, factura);

                    }
                    else
                    {
                        if (r["TCONFAC"] != System.DBNull.Value)
                        {
                            factura_detalle = new EndesaEntity.facturacion.Factura_Detalle();
                            factura_detalle.tconfac = Convert.ToInt32(r["TCONFAC"]);
                            factura_detalle.iconfac = Convert.ToDouble(r["ICONFAC"]);
                            o.lista_conceptos.Add(factura_detalle);
                        }
                    }

                }


            }
            db.CloseConnection();


            // Pintamos los datos en las facturas

            string ruta_plantilla = System.Environment.CurrentDirectory + @"\media\SEPARACION CAP DE GAS - FACTURA AGRUPADAS.xlsx";           

            FileInfo nombreSalidaExcel = new FileInfo(@"C:\Temp\SEPARACION CAP DE GAS - FACTURA AGRUPADAS_"
                   + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            FileInfo plantillaExcel = new FileInfo(ruta_plantilla);

            #region calculos-cim

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);
            var workSheet = excelPackage.Workbook.Worksheets["AGRUPADAS"];

            f = 4;


            workSheet = excelPackage.Workbook.Worksheets["AGRUPADAS"];

            foreach (KeyValuePair<string, EndesaEntity.facturacion.Factura> p in dic_facturas)
            {


                
                
                    error_fila = "";
                    f++;
                    c = 2;


                    workSheet.Cells[f, c].Value = p.Value.cnifdnic; c++;
                    workSheet.Cells[f, c].Value = p.Value.dapersoc; c++;
                    workSheet.Cells[f, c].Value = p.Value.creferen; c++;
                    workSheet.Cells[f, c].Value = p.Value.secfactu; c++;
                    workSheet.Cells[f, c].Value = p.Value.cfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.tfactura; c++;
                    workSheet.Cells[f, c].Value = p.Value.testfact; c++;

                    if (p.Value.ffactura > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = p.Value.ffactura;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.Value.vcuovafa;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;

                    workSheet.Cells[f, c].Value = p.Value.ifactura;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (p.Value.iva != 0)
                        workSheet.Cells[f, c].Value = p.Value.iva;
                    else
                        workSheet.Cells[f, c].Value = 0;

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    //else if (p.Value.iimpues2 != 0)
                    //    workSheet.Cells[f, c].Value = p.Value.iimpues2;
                    //else
                    //    workSheet.Cells[f, c].Value = p.Value.iimpues3;


                    if (p.Value.ise != 0)
                        workSheet.Cells[f, c].Value = p.Value.ise;
                    else
                    {
                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 660);
                        if (factura_detalle != null)
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                        else
                            workSheet.Cells[f, c].Value = 0;
                    }

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (p.Value.ibaseise != 0)
                        workSheet.Cells[f, c].Value = p.Value.ibaseise;
                    else
                    {
                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1261);

                        if (factura_detalle == null)
                            factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 650);

                        if (factura_detalle != null)
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                        else
                            workSheet.Cells[f, c].Value = 0;
                    }

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    c++;

                    factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 5162);
                    if (factura_detalle != null)                    
                        workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                    else
                    {
                        factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 5163);
                        if (factura_detalle != null)
                            workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                        else
                            workSheet.Cells[f, c].Value = 0;
                    }
                        
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    c++;
                                      


                    factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1256);

                    if (factura_detalle != null)
                        workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                    else
                        workSheet.Cells[f, c].Value = 0;

                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    c++;

                    factura_detalle = p.Value.lista_conceptos.Find(z => z.tconfac == 1222);

                    if (factura_detalle != null)                    
                        workSheet.Cells[f, c].Value = factura_detalle.iconfac;
                    else
                        workSheet.Cells[f, c].Value = 0;


                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";                    
                    c++;


                    

                


            }




            excelPackage.SaveAs(nombreSalidaExcel);
            excelPackage = null;
            #endregion
                       

        }

    }
}
