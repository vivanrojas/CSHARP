using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class FacturasBTN
    {
        public bool hayError;
        public string descripcion_error = "";
        public FacturasBTN()
        {
            hayError = false;
        }


        public void InformeFacturasBTN(DateTime fecha_factura_desde, DateTime fecha_factura_hasta, string fichero)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int f = 1;
            int c = 1;
            int total_registros = 0;
            double percent = 0;

            FileInfo fileInfo;            

            try
            {
                fileInfo = new FileInfo(fichero);
                if (fileInfo.Exists)
                    fileInfo.Delete();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("BTN");                

                var headerCells = workSheet.Cells[1, 1, 1, 26];
                var headerFont = headerCells.Style.Font;

                //headerFont.SetFromFont(new Font("Times New Roman", 12)); //Do this first
                headerFont.Bold = true;
                //headerFont.Italic = true;
                workSheet.Cells[f, c].Value = "CCOUNIPS"; c++;
                workSheet.Cells[f, c].Value = "CEMPTITU"; c++;
                workSheet.Cells[f, c].Value = "CREFEREN"; c++;
                workSheet.Cells[f, c].Value = "SECFACTU"; c++;
                workSheet.Cells[f, c].Value = "CREFFACT"; c++;
                workSheet.Cells[f, c].Value = "CFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTURA"; c++;
                workSheet.Cells[f, c].Value = "FFACTDES"; c++;
                workSheet.Cells[f, c].Value = "FFACTHAS"; c++;
                workSheet.Cells[f, c].Value = "VCUOVAFA"; c++;
                workSheet.Cells[f, c].Value = "VENEREAC"; c++;
                workSheet.Cells[f, c].Value = "VCUOFIFA"; c++;
                workSheet.Cells[f, c].Value = "IFACTURA"; c++;
                workSheet.Cells[f, c].Value = "IVA"; c++;
                workSheet.Cells[f, c].Value = "IBASEISE"; c++;
                workSheet.Cells[f, c].Value = "ISE"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_MEDIO"; c++;
                workSheet.Cells[f, c].Value = "TFACTURA"; c++;
                workSheet.Cells[f, c].Value = "TESTFACT"; c++;
                workSheet.Cells[f, c].Value = "DAPERSOC"; c++;
                workSheet.Cells[f, c].Value = "CNIFDNIC"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC1"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC1"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC2"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC2"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC3"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC3"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC4"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC4"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC5"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC5"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC6"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC6"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC7"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC7"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC8"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC8"; c++;
                workSheet.Cells[f, c].Value = "TCONFAC9"; c++;
                workSheet.Cells[f, c].Value = "ICONFAC9"; c++;
                workSheet.Cells[f, c].Value = "TCONFA10"; c++;
                workSheet.Cells[f, c].Value = "ICONFA10"; c++;
                workSheet.Cells[f, c].Value = "TCONFA11"; c++;
                workSheet.Cells[f, c].Value = "ICONFA11"; c++;
                workSheet.Cells[f, c].Value = "TCONFA12"; c++;
                workSheet.Cells[f, c].Value = "ICONFA12"; c++;
                workSheet.Cells[f, c].Value = "TCONFA13"; c++;
                workSheet.Cells[f, c].Value = "ICONFA13"; c++;
                workSheet.Cells[f, c].Value = "TCONFA14"; c++;
                workSheet.Cells[f, c].Value = "ICONFA14"; c++;
                workSheet.Cells[f, c].Value = "TCONFA15"; c++;
                workSheet.Cells[f, c].Value = "ICONFA15"; c++;
                workSheet.Cells[f, c].Value = "TCONFA16"; c++;
                workSheet.Cells[f, c].Value = "ICONFA16"; c++;
                workSheet.Cells[f, c].Value = "TCONFA17"; c++;
                workSheet.Cells[f, c].Value = "ICONFA17"; c++;
                workSheet.Cells[f, c].Value = "TCONFA18"; c++;
                workSheet.Cells[f, c].Value = "ICONFA18"; c++;
                workSheet.Cells[f, c].Value = "TCONFA19"; c++;
                workSheet.Cells[f, c].Value = "ICONFA19"; c++;
                workSheet.Cells[f, c].Value = "TCONFA20"; c++;
                workSheet.Cells[f, c].Value = "ICONFA20"; c++;
                workSheet.Cells[f, c].Value = "COMENTARIOS"; c++;
                workSheet.Cells[f, c].Value = "ID"; c++;
                workSheet.Cells[f, c].Value = "IIMPUES2"; c++;
                workSheet.Cells[f, c].Value = "IIMPUES3"; c++;
                workSheet.Cells[f, c].Value = "TIPO_CATEGORIA"; c++;
                workSheet.Cells[f, c].Value = "TINDGCPY"; c++;
                workSheet.Cells[f, c].Value = "TMODOPTA"; c++;
                workSheet.Cells[f, c].Value = "KPERFACT"; c++;
                workSheet.Cells[f, c].Value = "CREFAEXT"; c++;
                workSheet.Cells[f, c].Value = "TIPO_MOTIVO"; c++;
                workSheet.Cells[f, c].Value = "DESCRIPCION"; c++;
                workSheet.Cells[f, c].Value = "TIPO_COMENTARIO"; c++;
                workSheet.Cells[f, c].Value = "COMENTARIO_MOTIVO"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA1"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA2"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA3"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA4"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA5"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTRATADA6"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAXIMA1"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAXIMA2"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAXIMA3"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAXIMA4"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAXIMA5"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAXIMA6"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA5"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA6"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA_TOT"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA5"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA6"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA_TOT"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA5"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA6"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACTIVA1"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACTIVA2"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACTIVA3"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACTIVA4"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACTIVA5"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACTIVA6"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POTENCIA1"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POTENCIA2"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POTENCIA3"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POTENCIA4"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POTENCIA5"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POTENCIA6"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA1"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA2"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA3"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA4"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA5"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA6"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTR_LLANO"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTR_VALLE"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_CONTR_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAX_LLANO"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAX_VALLE"; c++;
                workSheet.Cells[f, c].Value = "POTENCIA_MAX_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA_LLANO"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA_VALLE"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_ACTIVA_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA_LLANO"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA_VALLE"; c++;
                workSheet.Cells[f, c].Value = "CONSUMO_REACTIVA_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA_LLANO"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA_VALLE"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_REACTIVA_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACT_BLQ1"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACT_BLQ2"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_ACT_BLQ3"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POT_BLQ1"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POT_BLQ2"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_POT_BLQ3"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA_LLANO"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA_VALLE"; c++;
                workSheet.Cells[f, c].Value = "EXCESO_POTENCIA_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_DH_LLANO"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_DH_VALLE"; c++;
                workSheet.Cells[f, c].Value = "PRECIO_DH_PUNTA"; c++;
                workSheet.Cells[f, c].Value = "CUPSREE"; c++;
                workSheet.Cells[f, c].Value = "CFACTREC"; c++;
                workSheet.Cells[f, c].Value = "CLINNEG"; c++;
                workSheet.Cells[f, c].Value = "CFACAGP"; c++;
                workSheet.Cells[f, c].Value = "CONTRATO COMERCIAL"; c++;




                total_registros = TotalRegistros(fecha_factura_desde, fecha_factura_hasta) + 1;

                forms.FrmProgressBar pb = new forms.FrmProgressBar();               
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = total_registros;
                pb.Text = "Generando informe... ";
                pb.Show();

                #region Query
                strSql = "SELECT" 
                    + " f.CCOUNIPS, f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.CREFFACT, f.CFACTURA,"
                    + " f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.VCUOVAFA, f.VENEREAC, f.VCUOFIFA," 
                    + " f.IFACTURA, f.IVA, f.IBASEISE, f.ISE, f.PRECIO_MEDIO,"
                    + " f.TFACTURA, f.TESTFACT, f.DAPERSOC, f.CNIFDNIC,"
                    + " f.TCONFAC1, f.ICONFAC1, f.TCONFAC2, f.ICONFAC2, f.TCONFAC3, f.ICONFAC3,"
                    + " f.TCONFAC4, f.ICONFAC4, f.TCONFAC5, f.ICONFAC5, f.TCONFAC6, f.ICONFAC6,"
                    + " f.TCONFAC7, f.ICONFAC7, f.TCONFAC8, f.ICONFAC8, f.TCONFAC9, f.ICONFAC9,"
                    + " f.TCONFA10, f.ICONFA10, f.TCONFA11, f.ICONFA11, f.TCONFA12, f.ICONFA12,"
                    + " f.TCONFA13, f.ICONFA13, f.TCONFA14, f.ICONFA14, f.TCONFA15, f.ICONFA15,"
                    + " f.TCONFA16, f.ICONFA16, f.TCONFA17, f.ICONFA17, f.TCONFA18, f.ICONFA18,"
                    + " f.TCONFA19, f.ICONFA19, f.TCONFA20, f.ICONFA20, f.COMENTARIOS, NULL as ID,"
                    + " f.IIMPUES2, f.IIMPUES3, f.CSEGMERC as TIPO_CATEGORIA, f.TINDGCPY, f.TMODOPTA,"
                    + " f.KPERFACT, f.CREFAEXT, f.MOTIVO_REFACTURACION AS TIPO_MOTIVO,"
                    + " f.SUBMOTIVO AS DESCRIPCION, f.TIPO_COMENTARIO,"
                    + " f.COMENTARIO_REFACT AS COMENTARIO_MOTIVO,"
                    + " f.VPOTCON1 AS POTENCIA_CONTRATADA1,"
                    + " f.VPOTCON2 AS POTENCIA_CONTRATADA2,"
                    + " f.VPOTCON3 AS POTENCIA_CONTRATADA3,"
                    + " f.VPOTCON4 AS POTENCIA_CONTRATADA4,"
                    + " f.VPOTCON5 AS POTENCIA_CONTRATADA5,"
                    + " f.VPOTCON6 AS POTENCIA_CONTRATADA6,"
                    + " f.VPOTMAX1 AS POTENCIA_MAXIMA1,"
                    + " f.VPOTMAX2 AS POTENCIA_MAXIMA2,"
                    + " f.VPOTMAX3 AS POTENCIA_MAXIMA3,"
                    + " f.VPOTMAX4 AS POTENCIA_MAXIMA4,"
                    + " f.VPOTMAX5 AS POTENCIA_MAXIMA5,"
                    + " f.VPOTMAX6 AS POTENCIA_MAXIMA6,"
                    + " f.VCONATH1 AS CONSUMO_ACTIVA1,"
                    + " f.VCONATH2 AS CONSUMO_ACTIVA2,"
                    + " f.VCONATH3 AS CONSUMO_ACTIVA3,"
                    + " f.VCONATH4 AS CONSUMO_ACTIVA4,"
                    + " f.VCONATH5 AS CONSUMO_ACTIVA5,"
                    + " f.VCONATH6 AS CONSUMO_ACTIVA6,"
                    + " f.VCONATHP AS CONSUMO_ACTIVA_TOT,"
                    + " f.VCONRTH1 AS CONSUMO_REACTIVA1,"
                    + " f.VCONRTH2 AS CONSUMO_REACTIVA2,"
                    + " f.VCONRTH3 AS CONSUMO_REACTIVA3,"
                    + " f.VCONRTH4 AS CONSUMO_REACTIVA4,"
                    + " f.VCONRTH5 AS CONSUMO_REACTIVA5,"
                    + " f.VCONRTH6 AS CONSUMO_REACTIVA6,"
                    + " f.VCONRTHP AS CONSUMO_REACTIVA_TOT,"
                    + " f.VEXCERE1 AS EXCESO_REACTIVA1,"
                    + " f.VEXCERE2 AS EXCESO_REACTIVA2,"
                    + " f.VEXCERE3 AS EXCESO_REACTIVA3,"
                    + " f.VEXCERE4 AS EXCESO_REACTIVA4,"
                    + " f.VEXCERE5 AS EXCESO_REACTIVA5,"
                    + " f.VEXCERE6 AS EXCESO_REACTIVA6,"
                    + " f.PRECIAC1 AS PRECIO_ACTIVA1,"
                    + " f.PRECIAC2 AS PRECIO_ACTIVA2,"
                    + " f.PRECIAC3 AS PRECIO_ACTIVA3,"
                    + " f.PRECIAC4 AS PRECIO_ACTIVA4,"
                    + " f.PRECIAC5 AS PRECIO_ACTIVA5,"
                    + " f.PRECIAC6 AS PRECIO_ACTIVA6,"
                    + " f.PRECIPO1 AS PRECIO_POTENCIA1,"
                    + " f.PRECIPO2 AS PRECIO_POTENCIA2,"
                    + " f.PRECIPO3 AS PRECIO_POTENCIA3,"
                    + " f.PRECIPO4 AS PRECIO_POTENCIA4,"
                    + " f.PRECIPO5 AS PRECIO_POTENCIA5,"
                    + " f.PRECIPO6 AS PRECIO_POTENCIA6,"
                    + " f.VEXCEPO1 AS EXCESO_POTENCIA1,"
                    + " f.VEXCEPO2 AS EXCESO_POTENCIA2,"
                    + " f.VEXCEPO3 AS EXCESO_POTENCIA3,"
                    + " f.VEXCEPO4 AS EXCESO_POTENCIA4,"
                    + " f.VEXCEPO5 AS EXCESO_POTENCIA5,"
                    + " f.VEXCEPO6 AS EXCESO_POTENCIA6,"
                    + " f.VPOTCALL AS POTENCIA_CONTR_LLANO,"
                    + " f.VPOTCALV AS POTENCIA_CONTR_VALLE,"
                    + " f.VPOTCALP AS POTENCIA_CONTR_PUNTA,"
                    + " f.VPOTMAXL AS POTENCIA_MAX_LLANO,"
                    + " f.VPOTMAXV AS POTENCIA_MAX_VALLE,"
                    + " f.VPOTMAXP AS POTENCIA_MAX_PUNTA,"
                    + " f.VCONSACL AS CONSUMO_ACTIVA_LLANO,"
                    + " f.VCONSACV AS CONSUMO_ACTIVA_VALLE,"
                    + " f.VCONSACP AS CONSUMO_ACTIVA_PUNTA,"
                    + " f.VCONSREL AS CONSUMO_REACTIVA_LLANO,"
                    + " f.VCONSREV AS CONSUMO_REACTIVA_VALLE,"
                    + " f.VCONSREP AS CONSUMO_REACTIVA_PUNTA,"
                    + " f.VEXCEREL AS EXCESO_REACTIVA_LLANO,"
                    + " f.VEXCEREV AS EXCESO_REACTIVA_VALLE,"
                    + " f.VEXCEREP AS EXCESO_REACTIVA_PUNTA,"
                    + " f.PRECIAB1 AS PRECIO_ACT_BLQ1,"
                    + " f.PRECIAB2 AS PRECIO_ACT_BLQ2,"
                    + " f.PRECIAB3 AS PRECIO_ACT_BLQ3,"
                    + " f.PRECIPB1 AS PRECIO_POT_BLQ1,"
                    + " f.PRECIPB2 AS PRECIO_POT_BLQ2,"
                    + " f.PRECIPB3 AS PRECIO_POT_BLQ3,"
                    + " f.VEXCEPOL AS EXCESO_POTENCIA_LLANO,"
                    + " f.VEXCEPOV AS EXCESO_POTENCIA_VALLE,"
                    + " f.VEXCEPOP AS EXCESO_POTENCIA_PUNTA,"
                    + " f.PRECIDHL AS PRECIO_DH_LLANO,"
                    + " f.PRECIDHV AS PRECIO_DH_VALLE,"
                    + " f.PRECIDHP AS PRECIO_DH_PUNTA,"
                    + " f.CUPSREE,"
                    + " f.CFACTREC,"
                    + " f.CLINNEG,"
                    + " ag.CFACAGP,"
                    + " ag.CONTRATO_COMERCIAL"
                    + " FROM fact.fo f"
                    + " INNER JOIN fact.fo_empresas e ON"
                    + " e.empresa_id = f.ID_ENTORNO"
                    + " INNER JOIN fact.fo_p_tipos_factura tf ON"
                    + " tf.id_tipo_factura = f.TFACTURA"
                    + " LEFT OUTER JOIN fact.fo_agrupadas ag ON"
                    + " ag.empresa_id = f.ID_ENTORNO and"
                    + " ag.CREFEREN = f.CREFEREN AND"
                    + " ag.SECFACTU = f.SECFACTU AND"
                    + " ag.TESTFACT = f.TESTFACT"                   
                    + " WHERE e.descripcion = 'BTN-Portugal'"
                    + " AND (f.FFACTURA >= '" + fecha_factura_desde.ToString("yyyy-MM-dd") + "'"
                    + " AND f.FFACTURA <= '" + fecha_factura_hasta.ToString("yyyy-MM-dd") + "')" 
                    + " ORDER BY f.CREFEREN, f.SECFACTU";

                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    percent = (f / Convert.ToDouble(total_registros)) * 100;
                    pb.txtDescripcion.Text = "Exportando: "
                       + f + " / " + total_registros;                       

                    pb.progressBar.Value = f;
                    pb.progressBar.Value = f - 1;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    f++;
                    c = 1;
                    if(r["CCOUNIPS"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CCOUNIPS"].ToString(); 
                    c++;

                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["CEMPTITU"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt64(r["CREFEREN"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["SECFACTU"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt64(r["CREFFACT"]); c++;

                    if (r["CFACTURA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CFACTURA"].ToString(); 
                    c++;
                    if (r["FFACTURA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FFACTURA"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (r["FFACTDES"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FFACTDES"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (r["FFACTHAS"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["FFACTHAS"]); 
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    if (r["VCUOVAFA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt64(r["VCUOVAFA"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }                        
                    c++;

                    if (r["VENEREAC"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["VENEREAC"]);
                    }                         
                    c++;

                    if (r["VCUOFIFA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["VCUOFIFA"].ToString();
                    }
                    c++;

                    if (r["IFACTURA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["IFACTURA"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    c++;

                    if (r["IVA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["IVA"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    c++;

                    if (r["IBASEISE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt64(r["IBASEISE"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }                    
                    c++;

                    if (r["ISE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["ISE"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                    c++;

                    if (r["PRECIO_MEDIO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_MEDIO"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.000000";
                    }
                    c++;

                    if (r["TFACTURA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["TFACTURA"]); 
                    }
                    c++;

                    if (r["TESTFACT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TESTFACT"].ToString(); 
                    }
                    c++;


                    workSheet.Cells[f, c].Value = r["DAPERSOC"].ToString(); c++;
                    workSheet.Cells[f, c].Value = r["CNIFDNIC"].ToString(); c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC1"]); c++;

                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC1"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC2"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC2"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC3"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC3"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC4"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC4"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC5"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC5"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC6"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC6"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC7"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC7"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC8"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC8"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFAC9"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFAC9"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA10"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA10"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA11"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA11"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA12"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA12"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA13"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA13"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA14"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA14"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA15"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA15"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA16"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA16"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA17"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA17"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA18"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA18"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA19"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA19"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = Convert.ToInt32(r["TCONFA20"]); c++;
                    workSheet.Cells[f, c].Value = Convert.ToDouble(r["ICONFA20"]); workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                    if (r["COMENTARIOS"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["COMENTARIOS"].ToString(); 
                    c++;

                    workSheet.Cells[f, c].Value = ""; c++; // ID

                    if (r["IIMPUES2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["IIMPUES2"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }
                        
                    c++;
                    if (r["IIMPUES3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["IIMPUES3"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00";
                    }                        
                    c++;

                    if (r["TIPO_CATEGORIA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TIPO_CATEGORIA"].ToString();
                    }
                    c++;

                    if (r["TINDGCPY"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TINDGCPY"].ToString();
                    }
                    c++;

                    if (r["TMODOPTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TMODOPTA"].ToString();
                    }
                    c++;

                    if (r["KPERFACT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["KPERFACT"].ToString();
                    }
                    c++;
                    if (r["CREFAEXT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CREFAEXT"].ToString();
                    }
                    c++;
                    if (r["TIPO_MOTIVO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TIPO_MOTIVO"].ToString();
                    }

                    c++;
                    if (r["DESCRIPCION"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["DESCRIPCION"].ToString();
                    }
                    c++;

                    if (r["TIPO_COMENTARIO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["TIPO_COMENTARIO"].ToString();
                    }
                    c++;
                    if (r["COMENTARIO_MOTIVO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["COMENTARIO_MOTIVO"].ToString();
                    }
                    c++;
                    if (r["POTENCIA_CONTRATADA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTRATADA1"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_CONTRATADA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTRATADA2"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_CONTRATADA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTRATADA3"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_CONTRATADA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTRATADA4"].ToString(); 
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                        c++;
                    if (r["POTENCIA_CONTRATADA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTRATADA5"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                        c++;
                    if (r["POTENCIA_CONTRATADA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTRATADA6"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;
                    if (r["POTENCIA_MAXIMA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAXIMA1"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_MAXIMA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAXIMA2"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_MAXIMA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAXIMA3"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_MAXIMA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAXIMA4"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_MAXIMA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAXIMA5"].ToString();
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["POTENCIA_MAXIMA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["POTENCIA_MAXIMA6"]); 
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA1"]);
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA2"]);
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA3"]);
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA4"]);
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA5"]); 
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA6"]);
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA_TOT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_ACTIVA_TOT"]);
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA1"]);
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA2"]);
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA3"]);
                    }
                        c++;

                    if (r["CONSUMO_REACTIVA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA4"]);
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA5"]);
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA6"]);
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA_TOT"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["CONSUMO_REACTIVA_TOT"]);
                    }
                    c++;

                    if (r["EXCESO_REACTIVA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_REACTIVA1"]);
                    }
                    c++;

                    if (r["EXCESO_REACTIVA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_REACTIVA2"].ToString();
                    }
                    c++;

                    if (r["EXCESO_REACTIVA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_REACTIVA3"]);
                    }
                    c++;

                    if (r["EXCESO_REACTIVA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_REACTIVA4"]);
                    }
                    c++;

                    if (r["EXCESO_REACTIVA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_REACTIVA5"]);
                    }
                    c++;

                    if (r["EXCESO_REACTIVA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_REACTIVA6"]);
                    }
                    c++;

                    if (r["PRECIO_ACTIVA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_ACTIVA1"]);
                    }
                    c++;

                    if (r["PRECIO_ACTIVA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_ACTIVA2"]);
                    }
                    c++;

                    if (r["PRECIO_ACTIVA3"] != System.DBNull.Value) 
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_ACTIVA3"]); 
                    }
                    c++;

                    if (r["PRECIO_ACTIVA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_ACTIVA4"]);
                    }
                    c++;

                    if (r["PRECIO_ACTIVA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_ACTIVA5"]);
                    }
                    c++;

                    if (r["PRECIO_ACTIVA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_ACTIVA6"]);
                    }
                    c++;

                    if (r["PRECIO_POTENCIA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_POTENCIA1"]);
                    }
                        c++;
                    if (r["PRECIO_POTENCIA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_POTENCIA2"]);
                    }
                    c++;

                    if (r["PRECIO_POTENCIA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_POTENCIA3"]);
                    }
                    c++;

                    if (r["PRECIO_POTENCIA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_POTENCIA4"]);
                    }
                    c++;

                    if (r["PRECIO_POTENCIA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["PRECIO_POTENCIA5"].ToString(); 
                    }
                    c++;

                    if (r["PRECIO_POTENCIA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["PRECIO_POTENCIA6"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_POTENCIA1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_POTENCIA1"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_POTENCIA2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_POTENCIA2"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_POTENCIA3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_POTENCIA3"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_POTENCIA4"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_POTENCIA4"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_POTENCIA5"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_POTENCIA5"].ToString(); 
                    }
                    c++;
                    
                    if (r["EXCESO_POTENCIA6"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_POTENCIA6"].ToString(); 
                    }
                    c++;

                    if (r["POTENCIA_CONTR_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTR_LLANO"].ToString(); 
                    }
                    c++;

                    if (r["POTENCIA_CONTR_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTR_VALLE"].ToString();
                    }
                    c++;

                    if (r["POTENCIA_CONTR_PUNTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_CONTR_PUNTA"].ToString();
                    }
                    c++;

                    if (r["POTENCIA_MAX_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAX_LLANO"].ToString(); 
                    }
                    c++;

                    if (r["POTENCIA_MAX_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAX_VALLE"].ToString(); 
                    }
                    c++;

                    if (r["POTENCIA_MAX_PUNTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["POTENCIA_MAX_PUNTA"].ToString(); 
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONSUMO_ACTIVA_LLANO"].ToString(); 
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONSUMO_ACTIVA_VALLE"].ToString(); 
                    }
                    c++;

                    if (r["CONSUMO_ACTIVA_PUNTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONSUMO_ACTIVA_PUNTA"].ToString(); 
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONSUMO_REACTIVA_LLANO"].ToString(); 
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONSUMO_REACTIVA_VALLE"].ToString(); 
                    }
                    c++;

                    if (r["CONSUMO_REACTIVA_PUNTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONSUMO_REACTIVA_PUNTA"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_REACTIVA_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_REACTIVA_LLANO"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_REACTIVA_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_REACTIVA_VALLE"].ToString(); 
                    }
                    c++;

                    if (r["EXCESO_REACTIVA_PUNTA"] != System.DBNull.Value) 
                    {
                        workSheet.Cells[f, c].Value = r["EXCESO_REACTIVA_PUNTA"].ToString(); c++;
                    }
                        
                    if (r["PRECIO_ACT_BLQ1"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["PRECIO_ACT_BLQ1"].ToString(); 
                    }
                    c++;

                    if (r["PRECIO_ACT_BLQ2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["PRECIO_ACT_BLQ2"].ToString();
                    }
                    c++;

                    if (r["PRECIO_ACT_BLQ3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["PRECIO_ACT_BLQ3"].ToString();
                    }
                    c++;

                    if (r["PRECIO_POT_BLQ1"] != System.DBNull.Value) 
                    {
                        workSheet.Cells[f, c].Value = r["PRECIO_POT_BLQ1"].ToString(); 
                    }
                    c++;

                    if (r["PRECIO_POT_BLQ2"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_POT_BLQ2"]);
                    }
                    c++;

                    if (r["PRECIO_POT_BLQ3"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_POT_BLQ3"]); 
                    }
                    c++;

                    if (r["EXCESO_POTENCIA_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_POTENCIA_LLANO"]);
                    }
                    c++;    

                    if (r["EXCESO_POTENCIA_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_POTENCIA_VALLE"]);
                    }
                    c++;

                    if (r["EXCESO_POTENCIA_PUNTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["EXCESO_POTENCIA_PUNTA"]);
                    }
                    c++;

                    if (r["PRECIO_DH_LLANO"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_DH_LLANO"]);
                    }
                    c++;

                    if (r["PRECIO_DH_VALLE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_DH_VALLE"]);
                    }
                    c++;

                    if (r["PRECIO_DH_PUNTA"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToDouble(r["PRECIO_DH_PUNTA"]);
                    }
                    c++;

                    if (r["CUPSREE"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CUPSREE"].ToString();
                    }
                    c++;

                    if (r["CFACTREC"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CFACTREC"].ToString();
                    }
                    c++;

                    if (r["CLINNEG"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["CLINNEG"]);
                    }
                    c++;

                    if (r["CFACAGP"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CFACAGP"].ToString();
                    }

                    c++;

                    if (r["CONTRATO_COMERCIAL"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CONTRATO_COMERCIAL"].ToString();
                    }

                    c++;
                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, c];
                //var cellFont = allCells.Style.Font;
                //cellFont.SetFromFont(new Font("Calibri", 8));
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:EZ1"].AutoFilter = true;
                allCells.AutoFitColumns();
                excelPackage.Save();
                pb.Close();

            }
            catch(Exception e)
            {
                hayError = true;
                descripcion_error = e.Message;
                
            }
        }

        private int TotalRegistros(DateTime fecha_factura_desde, DateTime fecha_factura_hasta)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int total_registros = 0;

            #region Query
            strSql = "SELECT count(*) as total_registros"
                + " FROM fact.fo f"
                + " INNER JOIN fact.fo_empresas e ON"
                + " e.empresa_id = f.ID_ENTORNO"
                + " INNER JOIN fact.fo_p_tipos_factura tf ON"
                + " tf.id_tipo_factura = f.TFACTURA"
                + " WHERE e.descripcion = 'BTN-Portugal'"
                + " AND (f.FFACTURA >= '" + fecha_factura_desde.ToString("yyyy-MM-dd") + "'"
                + " AND f.FFACTURA <= '" + fecha_factura_hasta.ToString("yyyy-MM-dd") + "')";
                

            #endregion

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                total_registros = Convert.ToInt32(r["total_registros"]);                    
            }
            db.CloseConnection();

            return total_registros;
        }


    }
}
