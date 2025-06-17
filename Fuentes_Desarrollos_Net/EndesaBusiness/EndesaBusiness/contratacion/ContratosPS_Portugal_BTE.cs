using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace EndesaBusiness.contratacion
{
    public class ContratosPS_Portugal_BTE
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ContratosPS_Portugal_BTE");
        utilidades.Param param;


        public Dictionary<string, EndesaEntity.facturacion.InventarioTipologias> inventario { get; set; }

        // Para el control de la ejecucion
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        public ContratosPS_Portugal_BTE()
        {
            param = new utilidades.Param("contratos_ps_bte_param", servidores.MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            inventario = new Dictionary<string, EndesaEntity.facturacion.InventarioTipologias>();

        }


        public void CargaInventarioBTE()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                strSql = "SELECT b.EMPRESA_TITULAR, b.TX_CUPS_INT, b.TX_DISTRIBUIDORA, b.TX_CONTRATO_ATR,"
                    + " b.TX_VERSION_CONTATR, b.FH_FECHA_PREV_ALTA, b.FH_FECHA_ACTIVACION, b.FH_INI_VERSION,"
                    + " b.FH_PREV_BAJA, b.TX_ESTADO, b.TX_TARIFA_ACCESO, b.TX_TENSION, b.CAMPO13, b.TX_CCONTRATO_PS,"
                    + " b.TX_VERSION_PS, b.TX_POTENCIA_1, b.TX_POTENCIA_2, b.TX_POTENCIA_3, b.TX_POTENCIA_4, b.TX_POTENCIA_5,"
                    + " b.TX_POTENCIA_6, b.CAMPO22, b.CAMPO23, b.TX_CONT_COMER, b.CAMPO25, b.FH_BAJA, b.TX_GP, b.TX_MUNICIPIO,"
                    + " b.TX_LOCALIDADE, b.TX_NIF, b.TX_APELLIDO, b.TX_CPE, b.CAMPO33, b.CAMPO34, b.CAMPO35, b.CAMPO36, b.CAMPO37,"
                    + " b.CAMPO38, b.CAMPO39, b.CAMPO40, b.last_update_date, e.Descripcion AS desc_estado_contrato,"
                    + " ps.CCOMPOBJ, ps.CTARIFA as producto,"
                    + " agr.fh_periodo AS PERIODO_PENDIENTE, agr.cd_estado, de.de_estado ,agr.cd_subestado, ds.de_subestado"
                    + " FROM cont.contratos_ps_bte b"
                    + " LEFT OUTER JOIN cont.contratos_ps_complementos_bte ps ON"
                    + " ps.CUPS = b.TX_CUPS_INT"
                    + " INNER JOIN cont.cont_estadoscontrato e ON"
                    + " e.Cod_Estado = b.TX_ESTADO"
                    + " LEFT OUTER JOIN fact.t_ed_h_sap_pendiente_facturar_agrupado agr ON"
                    + " agr.cd_cups = b.TX_CUPS_INT AND agr.fh_envio = (SELECT MAX(fh_envio) FROM fact.t_ed_h_sap_pendiente_facturar_agrupado)"
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

                    if (r["TX_CUPS_INT"] != System.DBNull.Value)
                        c.cups13 = r["TX_CUPS_INT"].ToString();

                    if (r["TX_CPE"] != System.DBNull.Value)
                        c.cups22 = r["TX_CPE"].ToString();

                    if (r["TX_NIF"] != System.DBNull.Value)
                        c.cif = r["TX_NIF"].ToString();

                    if (r["TX_APELLIDO"] != System.DBNull.Value)
                        c.razon_social = r["TX_APELLIDO"].ToString();

                    c.convertido = false;

                    c.empresa = "BTE";

                    if (r["TX_TARIFA_ACCESO"] != System.DBNull.Value)
                        c.tarifa = r["TX_TARIFA_ACCESO"].ToString();

                    c.tipo_punto_suministro = "";

                    //c.provincia = r["TX_MUNICIPIO"].ToString();

                    if (r["TX_CONTRATO_ATR"] != System.DBNull.Value)
                        c.contrato_ext = r["TX_CONTRATO_ATR"].ToString();

                    if (r["TX_VERSION_PS"] != System.DBNull.Value)
                        c.version = Convert.ToInt32(r["TX_VERSION_PS"]);

                    if (r["FH_FECHA_ACTIVACION"] != System.DBNull.Value)
                    {
                        if (r["FH_FECHA_ACTIVACION"].ToString() != "0000-00-00")
                            c.fecha_puesta_servicio = Convert.ToDateTime(r["FH_FECHA_ACTIVACION"]);
                    }

                    if (r["desc_estado_contrato"] != System.DBNull.Value)
                        c.estado_contrato = r["desc_estado_contrato"].ToString();

                    if (r["TX_LOCALIDADE"] != System.DBNull.Value)
                        c.provincia = r["TX_LOCALIDADE"].ToString();

                    if (r["TX_MUNICIPIO"] != System.DBNull.Value)
                        c.territorio = r["TX_MUNICIPIO"].ToString();

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
                    if (!inventario.TryGetValue(c.cups13, out o))
                    {
                        if (r["CCOMPOBJ"] != System.DBNull.Value)
                            c.dic_complementos.Add(r["CCOMPOBJ"].ToString(), cc);

                        inventario.Add(c.cups13, c);

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
            catch (Exception ex)
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
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE");

                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {
                    

                    string archivo = param.GetValue("ruta_entrada")
                        + param.GetValue("prefijo_archivo").Replace("*",ff.UltimoDiaHabil().ToString("MMdd"))
                        + ".txt";

                    DescargaContratosBTE(ff.UltimoDiaHabil());

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_contratosPS_BTE")))
                        {

                            ImportarArchivoContratos(archivo);

                            param.code = "MD5_archivo_contratosPS_BTE";
                            param.from_date = new DateTime(2022, 07, 23);
                            param.to_date = new DateTime(4999, 12, 31);
                            param.value = md5;
                            param.Save();
                        }
                        else
                        {
                            ss_pp.Update_Comentario("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE",
                            "El archivo " + fileInfo.Name
                                + " no se ha actualizado.");
                        }

                    }
                    else
                    {
                        ss_pp.Update_Comentario("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE",
                        "El archivo " + fileInfo.Name
                        + " no existe.");
                        fileInfo.Delete();
                    }
                        
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.AddError(ex.Message);
                ss_pp.Update_Comentario("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE",
                      ex.Message);
            }
        }

        public void ImportacionContratosComplementos()
        {
            EndesaBusiness.utilidades.Fechas ff = new utilidades.Fechas();
            StringBuilder sb = new StringBuilder();
            string md5 = "";

            DateTime ultima_ejecucion = new DateTime();

            try
            {
                Console.WriteLine("Última Ejecución: "
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_BTE", "2_Importar_Contratos_BTE").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_BTE", "2_Importar_Contratos_BTE");

                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {


                    string archivo = param.GetValue("ruta_entrada")
                        + param.GetValue("prefijo_archivo_complementos");

                    DescargaContratosComplementosBTE();

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_contratos_complementos_PS_BTE")))
                        {


                            ImportarArchivoContratosComplementos(archivo);

                            param.code = "MD5_archivo_contratos_complementos_PS_BTE";
                            param.from_date = new DateTime(2022, 07, 23);
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

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_complementos_PS_Portugal_BTE", "2_Importar_Contratos_Complementos_BTE");

                strSql = "delete from contratos_ps_complementos_bte";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando contratos BTE");
                Console.WriteLine("========================");

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
                        sb.Append("replace into contratos_ps_complementos_bte");
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

                    sb.Append(",'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(campos[c])).Append("',"); c++; // CNIFDNIC
                    sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(campos[c])).Append("'"); // NOMBRE
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

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_complementos_PS_Portugal_BTE", "2_Importar_Contratos_Complementos_BTE");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void ImportarArchivoContratos(string archivo)
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

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE");

                strSql = "delete from contratos_ps_bte";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando contratos BTE");
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
                        sb.Append("replace into contratos_ps_bte");
                        sb.Append(" (EMPRESA_TITULAR, TX_CUPS_INT, TX_DISTRIBUIDORA, TX_CONTRATO_ATR, TX_VERSION_CONTATR,");
                        sb.Append(" FH_FECHA_PREV_ALTA, FH_FECHA_ACTIVACION, FH_INI_VERSION, FH_PREV_BAJA,");
                        sb.Append(" TX_ESTADO, TX_TARIFA_ACCESO, TX_TENSION, CAMPO13, TX_CCONTRATO_PS, TX_VERSION_PS,");
                        sb.Append(" TX_POTENCIA_1, TX_POTENCIA_2, TX_POTENCIA_3, TX_POTENCIA_4,");
                        sb.Append(" TX_POTENCIA_5, TX_POTENCIA_6, CAMPO22, CAMPO23, TX_CONT_COMER, CAMPO25,");
                        sb.Append(" FH_BAJA, TX_GP, TX_MUNICIPIO, TX_LOCALIDADE, TX_NIF, TX_APELLIDO, TX_CPE,");
                        sb.Append(" CAMPO33, CAMPO34, CAMPO35, CAMPO36, CAMPO37, CAMPO38, CAMPO39, CAMPO40");
                        sb.Append(" ) values ");
                        firstOnly = false;
                    }

                    sb.Append("(").Append(campos[c]).Append(","); c++; // EMPRESA_TITULAR
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // TX_CUPS_INT                  
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TX_DISTRIBUIDORA      
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_CONTRATO_ATR
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_VERSION_CONTATR

                    sb.Append(CF(campos[c])).Append(","); c++; // FH_FECHA_PREV_ALTA
                    sb.Append(CF(campos[c])).Append(","); c++; // FH_FECHA_ACTIVACION
                    sb.Append(CF(campos[c])).Append(","); c++; // FH_INI_VERSION
                    sb.Append(CF(campos[c])).Append(","); c++; // FH_PREV_BAJA

                    sb.Append(CN(campos[c])).Append(","); c++; // TX_ESTADO
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // TX_TARIFA_ACCESO     
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_TENSION                    
                    sb.Append(CN(campos[c])).Append(","); c++; // CAMPO13
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_CCONTRATO_PS
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_VERSION_PS

                    sb.Append(CN(campos[c])).Append(","); c++; // TX_POTENCIA_1                    
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_POTENCIA_2
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_POTENCIA_3
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_POTENCIA_4

                    sb.Append(CN(campos[c])).Append(","); c++; // TX_POTENCIA_5
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_POTENCIA_6
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CAMPO22    
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CAMPO23    
                    sb.Append(CN(campos[c])).Append(","); c++; // TX_CONT_COMER
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // CAMPO25    

                    sb.Append(CF(campos[c])).Append(","); c++; // FH_BAJA
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TX_GP
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TX_MUNICIPIO
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TX_LOCALIDADE
                    sb.Append("'").Append(campos[c].Trim()).Append("',"); c++; // TX_NIF
                    sb.Append("'").Append(utilidades.FuncionesTexto.RT(campos[c])).Append("',"); c++; // TX_APELLIDO
                    sb.Append("'").Append(campos[c].Trim()).Append("'"); c++; // TX_CPE

                    // CAMPO33 al CAMPO40
                    for (int j = 1; j <= 8; j++)
                    {
                        sb.Append(",'").Append(campos[c].Trim()).Append("'");
                        c++;
                    }
                   
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

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_PS_Portugal_BTE", "2_Importar_Contratos_BTE");
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("ImportarArchivoContratos: " + ex.Message);
                Console.WriteLine("ImportarArchivoContratos: " + ex.Message);
            }
        }

        private void DescargaContratosBTE(DateTime fecha)
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("script_contratos_ps_bte", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_PS_Portugal_BTE", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("script_contratos_ps_bte", DateTime.Now, DateTime.Now), fecha.ToString("MMdd"));

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_PS_Portugal_BTE", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
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

        private void DescargaContratosComplementosBTE()
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("script_contratos_complementos_ps_bte", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_complementos_PS_Portugal_BTE", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("script_contratos_complementos_ps_bte"));

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_complementos_PS_Portugal_BTE", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
            }


        }

        private string CF(string t)
        {
            if (t.Trim() == "00000000" || t.Trim() == "")
                return "null";
            else
                return "'" + t.Substring(0, 4)
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
