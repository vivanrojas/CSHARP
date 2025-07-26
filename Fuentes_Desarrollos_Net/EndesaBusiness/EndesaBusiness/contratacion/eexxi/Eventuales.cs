using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class Eventuales
    {

        utilidades.Param param;
        public Eventuales()
        {
            param = new utilidades.Param("eexxi_param", servidores.MySQLDB.Esquemas.CON);
        }

        public void CargaExcelEventuales(string fichero)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;

            int x = 0;
            int c = 1;
            int f = 1;
            int secuencial = 0;
            FileStream fs;
            ExcelPackage excelPackage;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            List<EndesaEntity.contratacion.xxi.Eventuales> lista
                = new List<EndesaEntity.contratacion.xxi.Eventuales>();
            Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> dic_solicitud_cups =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud>();

            try
            {
                EndesaBusiness.xml.XMLFunciones xml_A301 = new EndesaBusiness.xml.XMLFunciones();

                fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();
                f = 2; // Porque las 2 primeras filas son la cabecera
                for (int i = 1; i < 5000; i++)
                {
                    c = 1;
                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString() == "")
                        break;

                    EndesaEntity.contratacion.xxi.Eventuales xml = new EndesaEntity.contratacion.xxi.Eventuales();

                    secuencial = Convert.ToInt32(param.GetValue("secuencial_solicitud")) + 1;

                    xml.codigoDeSolicitud = param.GetValue("prefijo_solicitud") 
                        + DateTime.Now.ToString("yyyy") + secuencial.ToString().PadLeft(4, '0');

                    xml.codigoREEEmpresaEmisora = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.distribuidora = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.cups = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.cnae = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.indActivacion = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    
                    if(xml.indActivacion == "F")
                        xml.fechaActivacion = Convert.ToDateTime(workSheet.Cells[f, c].Value); 
                    
                    c++;

                    xml.tarifaATR = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                    for(int j = 1; j <= 6; j++)
                    {
                        if (workSheet.Cells[f, 1].Value.ToString() != "")
                            xml.potenciaPeriodo[j] = Convert.ToDouble(workSheet.Cells[f, c].Value); 
                        c++;
                    }

                    xml.tipoIdentificador = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.identificador = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.tipoPersona = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.razonSocial = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.telefono = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                    xml.indicadorTipoDireccion = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                    if(xml.indicadorTipoDireccion == "F")
                    {
                        xml.direccionPS_Calle = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        xml.contacto = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        xml.contacto_Numero = Convert.ToString(workSheet.Cells[f, c].Value); c++;
                        

                    }
                    else
                    {
                        c++; c++;c++;
                    }

                    xml.tipoContratoATR = Convert.ToString(workSheet.Cells[f, c].Value); c++;

                    if (workSheet.Cells[f, 1].Value.ToString() != "")
                        xml.fecha_de_baja  = Convert.ToDateTime(workSheet.Cells[f, c].Value); 
                    c++;

                    if (workSheet.Cells[f, 1].Value.ToString() != "")
                        xml.tipoDocAportado = Convert.ToString(workSheet.Cells[f, c].Value); 
                    c++;

                    if (workSheet.Cells[f, 1].Value.ToString() != "")
                        xml.direccionURL = Convert.ToString(workSheet.Cells[f, c].Value); 
                    c++;

                    if (workSheet.Cells[f, 1].Value.ToString() != "")
                        xml.observaciones = Convert.ToString(workSheet.Cells[f, c].Value);
                    c++;

                    #region Generamos XML A301

                    GuardaNumSecuencialTemporal(secuencial);

                    FileInfo ficheroSalida = new FileInfo(param.GetValue("RutaSalidaXML_A3101")
                    + xml.codigoREEEmpresaEmisora + "_"
                    + xml.codigoREEEmpresaDestino + "_"
                    + "A3_01_"
                    + xml.cups + "_"
                    + "01_"
                    + secuencial.ToString().PadLeft(4, '0')
                    + ".xml");

                    #endregion

                    lista.Add(xml);



                }
                fs = null;
                excelPackage = null;


                #region Guardamos los datos en MySQL


                foreach (EndesaEntity.contratacion.xxi.Eventuales p in lista)
                {
                    x++;

                    if (firstOnly)
                    {
                        sb.Append("replace into eexxi_eventuales (empresa_emisora, distribuidora,");
                        sb.Append("cups22, cnae, indActivacion, fecha_activacion, tarifa,");
                        sb.Append("potencia_periodo1, potencia_periodo2, potencia_periodo3,");
                        sb.Append("potencia_periodo4, potencia_periodo5, potencia_periodo6,");
                        sb.Append("tension, tipo_identificador, identificador, tipo_persona,");
                        sb.Append("razon_social, telefono, indicador_tipodireccion, direccion,");
                        sb.Append("contacto, telf_contacto, tipo_contrato, fecha_baja, tipo_doc_aportado,");
                        sb.Append("direccion_url, observaciones, fichero, created_by, last_update_date) values ");
                        firstOnly = false;
                    }
                    sb.Append("('").Append(p.codigoREEEmpresaEmisora).Append("',");
                    sb.Append("'").Append(p.distribuidora).Append("',");
                    sb.Append("'").Append(p.cups).Append("',");
                    sb.Append(p.cnae).Append(",");
                    sb.Append("'").Append(p.indActivacion).Append("',");

                    if (p.fechaActivacion > DateTime.MinValue)
                        sb.Append("'").Append(p.fechaActivacion.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(p.tarifaATR).Append("',");

                    for(int i = 1; i <= 6; i++)
                    {
                        if (p.potenciaPeriodo[i] > 0)
                            sb.Append(p.potenciaPeriodo[i]).Append(",");
                        else
                            sb.Append("null,");
                    }

                    sb.Append("'").Append(p.tipoIdentificador).Append("',");
                    sb.Append("'").Append(p.identificador).Append("',");
                    sb.Append("'").Append(p.tipoPersona).Append("',");
                    sb.Append("'").Append(p.razonSocial).Append("',");

                    if(p.contacto_Numero != null)
                        sb.Append("'").Append(p.contacto_Numero).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(p.indicadorTipoDireccion).Append("',");

                    if (p.direccionPS_Calle != null)
                        sb.Append("'").Append(p.direccionPS_Calle).Append("',");
                    else
                        sb.Append("null,");

                    if (p.contacto != null)
                        sb.Append("'").Append(p.contacto).Append("',");
                    else
                        sb.Append("null,");

                    if (p.contacto_Numero != null)
                        sb.Append("'").Append(p.contacto_Numero).Append("',");
                    else
                        sb.Append("null,");

                    if (p.tipoContratoATR != null)
                        sb.Append("'").Append(p.tipoContratoATR).Append("',");
                    else
                        sb.Append("null,");

                    if (p.fecha_de_baja > DateTime.MinValue)
                        sb.Append("'").Append(p.fecha_de_baja.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.tipoDocAportado != null)
                        sb.Append("'").Append(p.tipoDocAportado).Append("',");
                    else
                        sb.Append("null,");

                    if (p.direccionURL != null)
                        sb.Append("'").Append(p.direccionURL).Append("',");
                    else
                        sb.Append("null,");

                    if (p.observaciones != null)
                        sb.Append("'").Append(p.observaciones).Append("'),");
                    else
                        sb.Append("null),");


                    if (x > 250)
                    {

                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
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
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }

                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "CargaExcelEventuales",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }    
            
        }

        private void GuardaNumSecuencialTemporal(int secuencial_solicitud)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update eexxi_param set value = '" + secuencial_solicitud + "'"
                + " where code = 'secuencial_solicitud'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // Volvemos a cargar parametros para tener el ultimo valor en memoria
            param = new utilidades.Param("eexxi_param", servidores.MySQLDB.Esquemas.CON);

        }

    }
}
