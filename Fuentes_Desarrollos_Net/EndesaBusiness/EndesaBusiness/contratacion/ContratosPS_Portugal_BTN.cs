using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using Oracle.ManagedDataAccess.Client;
using Renci.SshNet.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class ContratosPS_Portugal_BTN
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ContratosPS_Portugal_BTN");
        utilidades.Param param;

        Dictionary<string, List<EndesaEntity.contratacion.ComplementosContrato>> dic_contratos;

        public Dictionary<string, EndesaEntity.facturacion.InventarioTipologias> inventario { get; set; }

        // Para el control de la ejecucion
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;


        public ContratosPS_Portugal_BTN()
        {
            param = new utilidades.Param("contratos_ps_btn_param", servidores.MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            dic_contratos = new Dictionary<string, List<EndesaEntity.contratacion.ComplementosContrato>>();
            inventario = new Dictionary<string, EndesaEntity.facturacion.InventarioTipologias>();
        }


        public void CargaInventarioBTN()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            
            try
            {

                //Dictionary<string, string> dic_cups_compor = CargaPuntos_Compor();

                strSql = "SELECT b.EMPRESA, b.IDU, b.DDISTRIB, b.CCONTATR, b.CNUMCATR,"
                    + " b.FPREALTA, b.FALTACON, b.FPSERCON, b.FPREVBAJ, b.TESTCONT,"
                    + " b.CTARIFA, b.VTENSIOM, b.TDISCHOR, b.CONTREXT, b.CNUMSCPS,"
                    + " b.VPOTCAL1, b.VPOTCAL2, b.VPOTCAL3, b.VPOTCAL4, b.VPOTCAL5,"
                    + " b.VPOTCAL6, b.CONSESTI, b.TINDGCPY, b.CCONTCOM, b.TTICONPS,"
                    + " b.FBAJACON, b.CSEGMERC, b.DMUNICIP, b.DPROVINC, b.CNIFDNIC,"
                    + " b.DAPERSOC, b.CUPSREE, b.TPERFCNS, b.TFACTURA, b.TPUNTMED,"
                    + " b.TINDTUR, b.TEQMEDIN, b.PROVMUNI, b.CLASESUM, b.TDIHORAA,"
                    + " b.CODFINCA, b.TELEMEDIDA, e.Descripcion as desc_estado_contrato,"
                    + " t.tipo_tarifa, ps.CCOMPOBJ, ps.CTARIFA as producto,"
                    + " agr.fh_periodo AS PERIODO_PENDIENTE, agr.cd_estado, de.de_estado ,agr.cd_subestado, ds.de_subestado"
                    + " FROM cont.contratos_ps_btn b"
                    + " LEFT OUTER JOIN cont.contratos_ps_complementos_btn ps ON"
                    + " ps.CUPS = b.IDU"
                    + " INNER JOIN cont.cont_estadoscontrato e ON"
                    + " e.Cod_Estado = b.TESTCONT"
                    + " LEFT OUTER JOIN cont.tarifas_btn t ON"
                    + " t.codigo_tarifa = b.CTARIFA"
                    + " LEFT OUTER JOIN fact.t_ed_h_sap_pendiente_facturar_agrupado agr ON" 
                    + " agr.cd_cups = b.CUPSREE AND agr.fh_envio = (SELECT MAX(fh_envio) FROM fact.t_ed_h_sap_pendiente_facturar_agrupado)"
                    + " LEFT OUTER JOIN fact.t_ed_p_estado_sap_pendiente_facturar de ON"
                    + " de.cd_estado = agr.cd_estado"
                    + " LEFT OUTER JOIN fact.t_ed_p_subestado_sap_pendiente_facturar ds ON"
                    + " ds.cd_subestado = agr.cd_subestado;";


                //db = new MySQLDB(MySQLDB.Esquemas.CON);
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                Console.WriteLine("Consultando BBDD ...");               
                
                while (r.Read())
                {
                    EndesaEntity.facturacion.InventarioTipologias c = new EndesaEntity.facturacion.InventarioTipologias();

                    if (r["IDU"] != System.DBNull.Value)
                    c.cups13 = r["IDU"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cups22 = r["CUPSREE"].ToString();



                    if (r["FALTACON"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FALTACON"]);

                    if (r["FBAJACON"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FBAJACON"]);

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.cif = r["CNIFDNIC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.razon_social = r["DAPERSOC"].ToString();

                    if (r["tipo_tarifa"] != System.DBNull.Value)
                        c.tipo_tarifa = r["tipo_tarifa"].ToString();

                    c.convertido = false;

                    c.empresa = "BTN";

                    if (r["CTARIFA"] != System.DBNull.Value)
                        c.tarifa = r["CTARIFA"].ToString();

                    c.tipo_punto_suministro = "";

                    if (r["VPOTCAL1"] != System.DBNull.Value)
                        c.kaveas = Convert.ToDouble(r["VPOTCAL1"]) / 1000;

                    //c.provincia = r["TX_MUNICIPIO"].ToString();

                    if (r["CONTREXT"] != System.DBNull.Value)
                        c.contrato_ext = r["CONTREXT"].ToString();

                    if (r["CNUMCATR"] != System.DBNull.Value)
                        c.version = Convert.ToInt32(r["CNUMCATR"]);

                    if (r["FPSERCON"] != System.DBNull.Value)
                    {
                        if(r["FPSERCON"].ToString() != "0000-00-00")
                            c.fecha_puesta_servicio = Convert.ToDateTime(r["FPSERCON"]);
                    }

                    if (r["DPROVINC"] != System.DBNull.Value)
                        c.provincia = r["DPROVINC"].ToString();

                    if (r["DMUNICIP"] != System.DBNull.Value)
                        c.territorio = r["DMUNICIP"].ToString();

                    if (r["TPUNTMED"] != System.DBNull.Value)
                        c.tipo_punto_suministro = r["TPUNTMED"].ToString();


                    if (r["desc_estado_contrato"] != System.DBNull.Value)
                        c.estado_contrato = r["desc_estado_contrato"].ToString();

                    //Periodo pendiente
                    if (r["PERIODO_PENDIENTE"] != System.DBNull.Value)
                        c.periodo_pendiente = r["PERIODO_PENDIENTE"].ToString();
                    
                    //Estado pendiente
                    if (r["cd_estado"] != System.DBNull.Value)
                    {
                        c.estado_pendiente = r["cd_estado"].ToString();
                        if (r["de_estado"] != System.DBNull.Value)
                            c.estado_pendiente = c.estado_pendiente + " " + r["de_estado"].ToString();
                    }
                    
                    //Subestado pendiente
                    if (r["cd_subestado"] != System.DBNull.Value)
                    {
                        c.subestado_pendiente = r["cd_subestado"].ToString();
                        if (r["de_subestado"] != System.DBNull.Value)
                            c.subestado_pendiente = c.subestado_pendiente + " " + r["de_subestado"].ToString();
                    }
                    

                    EndesaEntity.contratacion.ComplementosContrato cc = new EndesaEntity.contratacion.ComplementosContrato();

                    if (r["CCOMPOBJ"] != System.DBNull.Value)
                        cc.ccompobj = r["CCOMPOBJ"].ToString();

                    if (r["producto"] != System.DBNull.Value)
                        cc.producto = r["producto"].ToString();

                    EndesaEntity.facturacion.InventarioTipologias o;
                    if (!inventario.TryGetValue(c.cups22, out o))
                    {
                        if (r["CCOMPOBJ"] != System.DBNull.Value)
                            c.dic_complementos.Add(r["CCOMPOBJ"].ToString(), cc);

                        inventario.Add(c.cups22, c);
                    }                        
                    else
                    {
                        if (r["CCOMPOBJ"] != System.DBNull.Value)
                        {
                            EndesaEntity.contratacion.ComplementosContrato oo;
                            if (!c.dic_complementos.TryGetValue(r["CCOMPOBJ"].ToString(), out oo))
                                c.dic_complementos.Add(r["CCOMPOBJ"].ToString(), cc);
                        }
                            
                    }

                }
                db.CloseConnection();
            }
            catch(Exception ex)
            {
                ficheroLog.addError("CargaInventarioBTN: " + ex.Message);
            }
        }

        public void ImportacionContratos()
        {
            EndesaBusiness.utilidades.Fechas ff = new utilidades.Fechas();
            StringBuilder sb = new StringBuilder();
            string md5 = "";

            DateTime ultima_ejecucion = new DateTime();

            try
            {
                Console.WriteLine("Última Ejecución: "
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_BTN", "2_Importar_Contratos_BTN").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_BTN", "2_Importar_Contratos_BTN");

                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {


                    string archivo = param.GetValue("ruta_entrada")
                        + param.GetValue("prefijo_archivo").Replace("*", ff.UltimoDiaHabil().ToString("MMdd"))
                        + ".txt";

                    DescargaContratosBTN(ff.UltimoDiaHabil());

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_contratosPS_BTN")))
                        {


                            ImportarArchivoContratos(archivo);

                            param.code = "MD5_archivo_contratosPS_BTN";
                            param.from_date = new DateTime(2022, 09, 16);
                            param.to_date = new DateTime(4999, 12, 31);
                            param.value = md5;
                            param.Save();
                        }

                    }
                    else
                        fileInfo.Delete();
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.AddError(ex.Message);
            }
        }       

        public void ImportacionContratosComplementos()
        {

            StringBuilder sb = new StringBuilder();
            string md5 = "";

            DateTime ultima_ejecucion = new DateTime();

            try
            {
                Console.WriteLine("Última Ejecución: "
                   + ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_BTN", "2_Importar_Contratos_Complementos_BTN").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_BTN", "2_Importar_Contratos_Complementos_BTN");


                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {
                    string archivo = param.GetValue("ruta_entrada")
                        + param.GetValue("nombre_archivo");

                    DescargaContratosComplementosBTN();

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_contratos_complementos_PS_BTN")))
                        {
                            ImportarArchivoContratosComplementos(archivo);

                            param.code = "MD5_archivo_contratos_complementos_PS_BTN";
                            param.from_date = new DateTime(2022, 09, 16);
                            param.to_date = new DateTime(4999, 12, 31);
                            param.value = md5;
                            param.Save();
                        }


                    }
                    else
                    {
                        ss_pp.Update_Comentario("Contratación", "contratos_complementos_PS_Portugal_BTN",
                            "2_Importar_Contratos_Complementos_BTN",
                            "El archivo " + fileInfo.Name + " no se ha actualizado.");
                        fileInfo.Delete();
                    }
                        
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.AddError(ex.Message);
            }
        }

        public void ImportarArchivoContratos(string archivo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int c = 0;
            string strSql = "";
            int total_registros = 0;


            try
            {

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_PS_Portugal_BTN", "2_Importar_Contratos_BTN");

                strSql = "delete from contratos_ps_btn";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando contratos BTN");
                Console.WriteLine("========================");

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split('|');
                    i++;
                    c = 0;
                    total_registros++;

                    if (firstOnly)
                    {
                        sb.Append("replace into contratos_ps_btn");
                        sb.Append(" (EMPRESA, IDU, DDISTRIB, CCONTATR, CNUMCATR,");
                        sb.Append("FPREALTA, FALTACON, FPSERCON, FPREVBAJ,");
                        sb.Append("TESTCONT, CTARIFA, VTENSIOM, TDISCHOR,");
                        sb.Append("CONTREXT, CNUMSCPS, VPOTCAL1, VPOTCAL2,");
                        sb.Append("VPOTCAL3, VPOTCAL4, VPOTCAL5, VPOTCAL6,");
                        sb.Append("CONSESTI, TINDGCPY, CCONTCOM, TTICONPS,");
                        sb.Append("FBAJACON, CSEGMERC, DMUNICIP, DPROVINC,");
                        sb.Append("CNIFDNIC, DAPERSOC, CUPSREE, TPERFCNS,");
                        sb.Append("TFACTURA, TPUNTMED, TINDTUR, TEQMEDIN,");
                        sb.Append("PROVMUNI, CLASESUM, TDIHORAA, CODFINCA, TELEMEDIDA");
                        sb.Append(") values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++; // EMPRESA
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // IDU
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // DDISTRIB
                    sb.Append(CN(campos[c])).Append(","); c++; // CCONTATR
                    sb.Append(CN(campos[c])).Append(","); c++; // CNUMCATR


                    sb.Append(CF(campos[c])).Append(","); c++; // FPREALTA
                    sb.Append(CF(campos[c])).Append(","); c++; // FALTACON
                    sb.Append(CF(campos[c])).Append(","); c++; // FPSERCON
                    sb.Append(CF(campos[c])).Append(","); c++; // FPREVBAJ

                    sb.Append(CN(campos[c])).Append(","); c++; // TESTCONT
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CTARIFA
                    sb.Append(CN(campos[c])).Append(","); c++; // VTENSIOM
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TDISCHOR

                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CONTREXT
                    sb.Append(CN(campos[c])).Append(","); c++; // CNUMSCPS


                    // VPOTCALn
                    for (int k = 1; k <= 6; k++)
                    {
                        sb.Append(CN(campos[c])).Append(",");
                        c++;
                    }

                    sb.Append(CN(campos[c])).Append(","); c++; // CONSESTI
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TINDGCPY
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CCONTCOM
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TTICONPS

                    sb.Append(CF(campos[c])).Append(","); c++; // FBAJACON
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CSEGMERC 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // DMUNICIP 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // DPROVINC

                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CNIFDNIC 
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(campos[c])).Append("',"); c++; // DAPERSOC 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CUPSREE 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TPERFCNS

                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TFACTURA 
                    sb.Append(CN(campos[c])).Append(","); c++; // TPUNTMED 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TINDTUR 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TEQMEDIN

                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // PROVMUNI 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CLASESUM 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TDIHORAA 
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CODFINCA 
                    sb.Append("'").Append(campos[c].Trim()).Append("'"); c++; // TELEMEDIDA

                    sb.Append("),");

                    if (i == 250)
                    {

                        Console.CursorLeft = 0;
                        Console.Write(total_registros.ToString("N0"));

                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }
                file.Close();

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                ficheroLog.Add("Fin importación");
                ficheroLog.Add("Se han importado " + total_registros.ToString("N0") + " registros");
                FileInfo archivo_info = new FileInfo(archivo);
                archivo_info.Delete();

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_PS_Portugal_BTN", "2_Importar_Contratos_BTN");
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("ImportarArchivoContratos: " + ex.Message);
                Console.WriteLine("ImportarArchivoContratos: " + ex.Message);
            }
        }

        private void DescargaContratosBTN(DateTime fecha)
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("script_contratos_ps_btn", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_PS_Portugal_BTN", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("script_contratos_ps_btn", DateTime.Now, DateTime.Now), fecha.ToString("MMdd"));

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_PS_Portugal_BTN", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
            }


        }        

        private void DescargaContratosComplementosBTN()
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("script_contratos_complementos_ps_btn", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_complementos_PS_Portugal_BTN", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("script_contratos_complementos_ps_btn"));

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_complementos_PS_Portugal_BTN", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
            }


        }

        private Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> CargaContratosComplementosCalendarios(string ccounips)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string tarifa = "";
            int faltacon = 0;
            int fpsercon = 0;
            int fprevbaj = 0;
            int fbajacon = 0;

            Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos> d
                = new Dictionary<string, EndesaEntity.facturacion.InformeContratosComplementos>();

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            try
            {



                strSql = "SELECT ps.IDU AS CCOUNIPS, ps.CUPS22, ps.CONTREXT, ps.estadoCont AS TESTCONT,"
                    + " ps.fAltaCont AS FALTACON, ps.FPSERCON, ps.fPrevBajaCont AS FPREVBAJ,"
                    + " ps.fBajaCont AS FBAJACON, cc.CEMPTITU, cc.CCONTRPS, cc.CNUMSCPS, cc.CLINNEG,"
                    + " cc.CCLIENTE, cc.CCALENPO, "
                    + " cc.CTARIFA,  cc.CCOMPOBJ, cc.VNSEGHOR,"
                    + " cc.VPARAM01, cc.VPARAM02, cc.VPARAM03, cc.VPARAM04, cc.VPARAM05,"
                    + " ps.TENSION"
                    + " FROM PS_AT ps"
                    + " LEFT OUTER JOIN cont_comp_contratos_calendarios cc ON"
                    + " cc.CCOUNIPS = ps.IDU";
                //+ " cc.CONTREXT = ps.CONTREXT";

                if (ccounips != null)
                    strSql += " where ps.IDU = '" + ccounips + "'";

                strSql += " group by ps.IDU, cc.CCOMPOBJ";


                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.InformeContratosComplementos o;
                    if (!d.TryGetValue(r["CCOUNIPS"].ToString(), out o))
                    {

                        tarifa = r["CTARIFA"].ToString();

                        if (r["FALTACON"] != System.DBNull.Value)
                            faltacon = Convert.ToInt32(r["FALTACON"]);

                        if (r["FPSERCON"] != System.DBNull.Value)
                            fpsercon = Convert.ToInt32(r["FPSERCON"]);

                        if (r["FPREVBAJ"] != System.DBNull.Value)
                            fprevbaj = Convert.ToInt32(r["FPREVBAJ"]);

                        if (r["FBAJACON"] != System.DBNull.Value)
                            fbajacon = Convert.ToInt32(r["FBAJACON"]);

                        o = new EndesaEntity.facturacion.InformeContratosComplementos();
                        o.cups = r["CCOUNIPS"].ToString();

                        if (r["CUPS22"] != System.DBNull.Value)
                        {
                            if (r["CUPS22"].ToString().Length == 22)
                            {
                                o.cups20 = r["CUPS22"].ToString().Substring(0, 20);
                                o.cups22 = r["CUPS22"].ToString();
                            }
                        }


                        if (r["CONTREXT"] != System.DBNull.Value)
                            o.contrato = r["CONTREXT"].ToString();

                        if (r["CNUMSCPS"] != System.DBNull.Value)
                            o.version = Convert.ToInt32(r["CNUMSCPS"]);

                        if (r["TENSION"] != System.DBNull.Value)
                            o.tension = Convert.ToInt32(r["TENSION"]);

                        //o.dic = Cabecera(tarifa, faltacon, fpsercon, fprevbaj, fbajacon);

                        if (r["CCOMPOBJ"] == System.DBNull.Value)
                            o.dic.Add("Sin Datos Contrato", "Sin Datos Contrato");
                        else
                        {
                            string oo;
                            oo = r["CCOMPOBJ"].ToString();
                            o.dic.Add(oo, oo);

                        }



                        //d.Add(o.contrato, o);                        
                        d.Add(o.cups, o);

                    }
                    else
                    {
                        string oo;
                        if (r["CCOMPOBJ"] != System.DBNull.Value)
                        {
                            oo = r["CCOMPOBJ"].ToString();
                            o.dic.Add(oo, oo);
                        }

                    }


                }
                r.Close();
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                db.CloseConnection();
                ficheroLog.AddError("CargaContratos: " + e.Message);
                return null;
            }
        }

        private Dictionary<string, string> CargaPuntos_Compor()
        {
            // Ora
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string strSql = "";

            string cups13 = "";
            string cups20 = "";

            try
            {
                Dictionary<string, string> d = new Dictionary<string, string>();

                strSql = "SELECT TX_CUPS_INT, TX_CPE FROM APL_INVENTARIO_PUNTOS_ACTIVOS";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    if (r["TX_CUPS_INT"] != System.DBNull.Value)
                        cups13 = r["TX_CUPS_INT"].ToString();

                    if(r["TX_CPE"] != System.DBNull.Value)
                        cups20 = r["TX_CPE"].ToString();

                    string o;
                    if (!d.TryGetValue(cups13, out o))
                        d.Add(cups13, cups20);
                }
                ora_db.CloseConnection();
                return d;
            }
            catch(Exception ex)
            {
                ficheroLog.AddError("CargaPuntos_Compor: " + ex.Message);
                return null;
            }
        }

        public void ImportarArchivoContratosComplementos(string archivo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            string line = "";
            string[] campos;
            int i = 0;
            int c = 0;
            string strSql = "";
            int total_registros = 0;


            try
            {

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_complementos_PS_Portugal_BTN", "2_Importar_Contratos_Complementos_BTN");

                strSql = "delete from contratos_ps_complementos_btn";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando contratos complementos BTN");
                Console.WriteLine("======================================");

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split(';');
                    i++;
                    c = 1;
                    total_registros++;

                    if (firstOnly)
                    {
                        sb.Append("replace into contratos_ps_complementos_btn");
                        sb.Append(" (CUPS, CONTREXT, TESTCONT, FALTACON, FPSERCON, FPREVBAJ, FBAJACON,");
                        sb.Append(" CEMPTITU, CCONTRPS, VERSION, CLINNEG, CCLIENTE, CCALENPO,");
                        sb.Append(" VDIAFACT, CSEGMERC, FSIGFACT, FFINVESU, CTARIFA, CCOMPOBJ, VNSEGHOR,");
                        sb.Append(" VPARAM01, VPARAM02, VPARAM03, VPARAM04, VPARAM05, CNIFDNIC, NOMBRE");
                        sb.Append(" ) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++; // CUPS
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CONTREXT                  
                    sb.Append(CN(campos[c])).Append(","); c++; // TESTCONT
                    sb.Append(CF(campos[c])).Append(","); c++; // FALTACON
                    sb.Append(CF(campos[c])).Append(","); c++; // FPSERCON
                    sb.Append(CF(campos[c])).Append(","); c++; // FPREVBAJ
                    sb.Append(CF(campos[c])).Append(","); c++; // FBAJACON

                    sb.Append(CN(campos[c])).Append(","); c++; // CEMPTITU
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CCONTRPS     
                    sb.Append(CN(campos[c])).Append(","); c++; // VERSION                    
                    sb.Append(CN(campos[c])).Append(","); c++; // CLINNEG
                    sb.Append(CN(campos[c])).Append(","); c++; // CCLIENTE
                    sb.Append(CN(campos[c])).Append(","); c++; // CCALENPO

                    sb.Append(CN(campos[c])).Append(","); c++; // VDIAFACT 
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CSEGMERC
                    sb.Append(CF(campos[c])).Append(","); c++; // FSIGFACT
                    sb.Append(CF(campos[c])).Append(","); c++; // FFINVESU
                    sb.Append(CN(campos[c])).Append(","); c++; // CTARIFA                                        
                    sb.Append(CN(campos[c])).Append(","); c++; // CCOMPOBJ
                    sb.Append(CN(campos[c])); c++; // VNSEGHOR


                    // VPARAM0n
                    for (int j = 1; j <= 5; j++)
                    {
                        sb.Append(",").Append(CDouble(campos[c])); 
                        c++;
                    }                   
                    
                    sb.Append(",'").Append(utilidades.FuncionesTexto.RT(campos[c])).Append("',"); c++; // CNIFDNIC
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(campos[c])).Append("'"); // NOMBRE
                    sb.Append("),");

                    if (i == 250)
                    {

                        Console.CursorLeft = 0;
                        Console.Write(total_registros.ToString("N0"));

                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }

                }
                file.Close();

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                ficheroLog.Add("Fin importación");
                ficheroLog.Add("Se han importado " + total_registros + " registros");
                FileInfo archivo_info = new FileInfo(archivo);
                archivo_info.Delete();

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_complementos_PS_Portugal_BTN", "2_Importar_Contratos_Complementos_BTN");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.AddError("ImportarArchivoContratosComplementos: " + ex.Message);
            }
        }

        public bool TieneComplemento(string cups20, string complemento)
        {
            EndesaEntity.facturacion.InventarioTipologias o;
            EndesaEntity.contratacion.ComplementosContrato oo;
            if (inventario.TryGetValue(cups20, out o))
                return o.dic_complementos.TryGetValue(complemento, out oo);
            else
                return false;                    

        }

        public string GetProducto(string cups20)
        {
            string producto = "";
            EndesaEntity.facturacion.InventarioTipologias o;
            EndesaEntity.contratacion.ComplementosContrato oo;
            if (inventario.TryGetValue(cups20, out o))
            {
                if (o.dic_complementos.Count > 0)
                {
                    List<EndesaEntity.contratacion.ComplementosContrato> lista =
                        o.dic_complementos.Values.ToList();

                    producto = lista[0].producto;
                }
                    
            }

            return producto;
        }

        private string CF(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "null";
            else
                return "'" + t.Substring(0,4) 
                    + "-" + t.Substring(4, 2)
                    + "-" + t.Substring(6, 2) + "'";
        }

        private string CN(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "'null'";
            else
                return "'" + t.Trim() + "'";
        }

        public string CDouble(String t)
        {
            t = t.Trim();
            if (t == "")
            {
                return "null";
            }
            else
            {
                t = t.Replace(" ", "");
                t = t.Replace("+", string.Empty);
                t = t.Replace("----------", string.Empty);
                t = t.Replace(".", string.Empty);
                t = t.Replace(",", ".");

                if (t == "")
                {
                    t = "null";
                }
            }

            return t;
        }

    }
}
