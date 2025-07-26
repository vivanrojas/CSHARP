using EndesaBusiness.medida;
using EndesaBusiness.servidores;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class ContratosPS_Portugal_MT
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ContratosPS_Portugal_MT");
        utilidades.Param param;

        public Dictionary<string, EndesaEntity.facturacion.InventarioTipologias> inventario { get; set; }

        // Para el control de la ejecucion
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;

        public ContratosPS_Portugal_MT()
        {
            param = new utilidades.Param("contratos_ps_mt_param", servidores.MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
            inventario = new Dictionary<string, EndesaEntity.facturacion.InventarioTipologias>();
        }


        public void CargaInventarioMT()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                strSql = "SELECT b.CUPS, b.CONTREXT, b.VERSION, b.TESTCONT, b.FALTACON, b.FPSERCON,"
                    + " b.FPREVBAJ, b.FBAJACON, b.CEMPTITU, b.CLINNEG, b.CSEGMERC, b.CCONTRPS, b.CCLIENTE,"
                    + " b.CCALENPO, b.FSIGFACT, b.FFINVESU, b.CTARIFA, b.CUPSREE, b.CNIFDNIC, b.DAPERSOC, b.f_ult_mod,"
                    + " e.Descripcion AS desc_estado_contrato, ps.CCOMPOBJ, ps.CTARIFA as producto,"
                     + " agr.fh_periodo AS PERIODO_PENDIENTE, agr.cd_estado, de.de_estado ,agr.cd_subestado, ds.de_subestado"
                    + " FROM cont.contratos_ps_mt b"
                     + " LEFT OUTER JOIN cont.contratos_ps_complementos_mt ps ON"
                    + " ps.CCOUNIPS = b.CUPS"
                    + " INNER JOIN cont.cont_estadoscontrato e ON"
                    + " e.Cod_Estado = b.TESTCONT"
                    + " LEFT OUTER JOIN fact.t_ed_h_sap_pendiente_facturar_agrupado agr ON"
                    + " agr.cd_cups = b.CUPS AND agr.fh_envio = (SELECT MAX(fh_envio) FROM fact.t_ed_h_sap_pendiente_facturar_agrupado)"
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

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups13 = r["CUPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cups22 = r["CUPSREE"].ToString();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.cif = r["CNIFDNIC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.razon_social = r["DAPERSOC"].ToString();

                    if (r["VERSION"] != System.DBNull.Value)
                        c.version = Convert.ToInt32(r["VERSION"]);

                    c.convertido = false;

                    c.empresa = "MT";

                    if (r["CTARIFA"] != System.DBNull.Value)
                        c.tarifa = r["CTARIFA"].ToString();

                    c.tipo_punto_suministro = "";

                    //c.provincia = r["TX_MUNICIPIO"].ToString();

                    if (r["CONTREXT"] != System.DBNull.Value)
                        c.contrato_ext = r["CONTREXT"].ToString();

                    

                    if (r["FALTACON"] != System.DBNull.Value)
                    {
                        if (r["FALTACON"].ToString() != "0000-00-00")
                            c.fecha_puesta_servicio = Convert.ToDateTime(r["FALTACON"]);
                    }

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
                            if (!o.dic_complementos.TryGetValue(r["CCOMPOBJ"].ToString(), out oo))
                                o.dic_complementos.Add(r["CCOMPOBJ"].ToString(), cc);
                        }

                    }

                }
                db.CloseConnection();
            }
            catch (Exception ex)
            {
                ficheroLog.addError("CargaInventarioMT: " + ex.Message);
            }
        }

        public void CargaInventarioMT_ESP()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                strSql = "SELECT ps.EMPRESA, ps.IDU AS CUPS, ps.CUPS22 AS CUPSREE, ps.NIF AS CNIFDNIC,"
                    + " ps.Cliente AS DAPERSOC, ps.Version AS VERSION, ps.TARIFA AS CTARIFA,"
                    + " ps.CONTREXT, ps.fAltaCont AS FALTACON, e.Descripcion AS desc_estado_contrato,"
                    + " psc.CCOMPOBJ, psc.VPARAM02, null as producto"
                    + " FROM PS_AT ps"
                    + " LEFT OUTER JOIN contratos_ps_complementos_mt psc ON"
                    + " psc.CCOUNIPS = ps.IDU"
                    + " INNER JOIN cont_estadoscontrato e ON"
                    + " e.Cod_Estado = ps.estadoCont";



                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                Console.WriteLine("Consultando BBDD ...");

                while (r.Read())
                {
                    EndesaEntity.facturacion.InventarioTipologias c = new EndesaEntity.facturacion.InventarioTipologias();

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups13 = r["CUPS"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cups22 = r["CUPSREE"].ToString();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.cif = r["CNIFDNIC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.razon_social = r["DAPERSOC"].ToString();

                    if (r["VERSION"] != System.DBNull.Value)
                        c.version = Convert.ToInt32(r["VERSION"]);

                    c.convertido = false;

                    if (r["EMPRESA"] != System.DBNull.Value)
                        c.empresa = r["EMPRESA"].ToString();

                    if (r["CTARIFA"] != System.DBNull.Value)
                        c.tarifa = r["CTARIFA"].ToString();

                    c.tipo_punto_suministro = "";

                    //c.provincia = r["TX_MUNICIPIO"].ToString();

                    if (r["CONTREXT"] != System.DBNull.Value)
                        c.contrato_ext = r["CONTREXT"].ToString();


                    //if (r["FALTACON"] != System.DBNull.Value)
                    //{
                    //    if (r["FALTACON"].ToString() != "0000-00-00")
                    //        c.fecha_puesta_servicio = Convert.ToDateTime(r["FALTACON"]);
                    //}

                    if (r["desc_estado_contrato"] != System.DBNull.Value)
                        c.estado_contrato = r["desc_estado_contrato"].ToString();

                    EndesaEntity.contratacion.ComplementosContrato cc = new EndesaEntity.contratacion.ComplementosContrato();

                    if (r["CCOMPOBJ"] != System.DBNull.Value)
                        cc.ccompobj = r["CCOMPOBJ"].ToString();

                    if (r["producto"] != System.DBNull.Value)
                        cc.producto = r["producto"].ToString();

                    if (r["VPARAM02"] != System.DBNull.Value)
                        cc.vparam[1] = Convert.ToDouble(r["VPARAM02"]);                    

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
                            if (!o.dic_complementos.TryGetValue(r["CCOMPOBJ"].ToString(), out oo))
                                o.dic_complementos.Add(r["CCOMPOBJ"].ToString(), cc);
                        }

                    }

                }
                db.CloseConnection();
            }
            catch (Exception ex)
            {
                ficheroLog.addError("CargaInventarioMT: " + ex.Message);
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
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_MT", "2_Importar_Contratos_Complementos_MT").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Contratación", "contratos_complementos_PS_Portugal_MT", "2_Importar_Contratos_Complementos_MT");

                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {


                    string archivo = param.GetValue("ruta_entrada")
                        + param.GetValue("prefijo_archivo_complementos");

                    DescargaContratosComplementosMT();

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_contratos_complementos_PS_MT")))
                        {


                            ImportarArchivoContratosComplementos(archivo);

                            param.code = "MD5_archivo_contratos_complementos_PS_MT";
                            param.from_date = new DateTime(2018, 07, 23);
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

        private void DescargaContratosComplementosMT()
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("script_contratos_complementos_ps_mt", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_complementos_PS_Portugal_MT", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("script_contratos_complementos_ps_mt"));

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_complementos_PS_Portugal_MT", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
            }


        }

        private void ImportarArchivoContratosComplementos(string archivo)
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

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_complementos_PS_Portugal_MT", "2_Importar_Contratos_Complementos_MT");

                strSql = "delete from contratos_ps_complementos_mt";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando contratos complementos MT");
                Console.WriteLine("====================================");

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
                        sb.Append("replace into contratos_ps_complementos_mt");
                        sb.Append(" (CCOUNIPS, CONTREXT, TESTCONT, FALTACON,");
                        sb.Append(" CEMPTITU, CCONTRPS, CNUMSCPS, CLINNEG, CCLIENTE, CCALENPO,");
                        sb.Append(" CSEGMERC, CTARIFA, CCOMPOBJ, VNSEGHOR,");
                        sb.Append(" VPARAM01, VPARAM02, VPARAM03, VPARAM04, VPARAM05");
                        sb.Append(" ) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(campos[c]).Append("',"); c++; // CCOUNIPS
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CONTREXT                  
                    sb.Append(CN(campos[c])).Append(","); c++; // TESTCONT
                    sb.Append(CF(campos[c])).Append(","); c++; // FALTACON
                    //sb.Append(CF(campos[c])).Append(","); c++; // FPSERCON
                    //sb.Append(CF(campos[c])).Append(","); c++; // FPREVBAJ
                    //sb.Append(CF(campos[c])).Append(","); c++; // FBAJACON

                    sb.Append(CN(campos[c])).Append(","); c++; // CEMPTITU
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CCONTRPS
                    sb.Append(CN(campos[c])).Append(","); c++; // CNUMSCPS
                    // sb.Append(CN(campos[c])).Append(","); c++; // VERSION                    
                    sb.Append(CN(campos[c])).Append(","); c++; // CLINNEG
                    sb.Append(CN(campos[c])).Append(","); c++; // CCLIENTE
                    sb.Append(CN(campos[c])).Append(","); c++; // CCALENPO
                                        
                    sb.Append("'").Append(campos[c]).Append("',"); c++; // CSEGMERC                    
                    sb.Append(CN(campos[c])).Append(","); c++; // CTARIFA                                        
                    sb.Append(CN(campos[c])).Append(","); c++; // CCOMPOBJ
                    sb.Append(CN(campos[c])); c++; // VNSEGHOR


                    // VPARAM0n
                    for (int j = 1; j <= 5; j++)
                    {
                        sb.Append(",").Append(CDouble(campos[c]));
                        c++;
                    }

                    //sb.Append(",'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(campos[c])).Append("',"); c++; // CNIFDNIC
                    //sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(campos[c])).Append("'"); // NOMBRE
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

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_complementos_PS_Portugal_MT", "2_Importar_Contratos_Complementos_MT");
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("ImportarArchivoContratosComplementos: " + ex.Message);
                Console.WriteLine(ex.Message);
            }
        }

        public void DescargaContratosMT(DateTime fecha)
        {
            try
            {

                ficheroLog.Add("Ejecutando extractor: " + param.GetValue("script_contratos_ps_mt", DateTime.Now, DateTime.Now));

                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_PS_Portugal_MT", "1_Ejecución Extractor");

                utilidades.Fichero.EjecutaComando(param.GetValue("script_contratos_ps_mt"), fecha.ToString("MMdd"));

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_PS_Portugal_MT", "1_Ejecución Extractor");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Descarga --> " + e.Message);
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
                    + ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_MT", "2_Importar_Contratos_MT").ToString("dd/MM/yyyy"));

                ultima_ejecucion =
                ss_pp.GetFecha_FinProceso("Contratación", "contratos_PS_Portugal_MT", "2_Importar_Contratos_MT");

                Console.WriteLine(ultima_ejecucion.ToString("dd/MM/yyyy") + " > " + DateTime.Now.Date.ToString("dd/MM/yyyy"));

                if (ultima_ejecucion < DateTime.Now.Date)
                {

                    string archivo = param.GetValue("ruta_entrada")
                       + param.GetValue("prefijo_archivo");

                    DescargaContratosMT(ff.UltimoDiaHabil());

                    FileInfo fileInfo = new FileInfo(archivo);
                    if (fileInfo.Exists)
                    {

                        md5 = utilidades.Fichero.checkMD5(archivo).ToString();
                        if (fileInfo.Length > 0 && (md5 != param.GetValue("MD5_archivo_contratosPS_MT")))
                        {

                            ImportarArchivoContratos(archivo);

                            param.code = "MD5_archivo_contratosPS_MT";
                            param.from_date = new DateTime(2018, 07, 23);
                            param.to_date = new DateTime(4999, 12, 31);
                            param.value = md5;
                            param.Save();


                        }

                    }
                    else
                    {
                        ficheroLog.Add("md5 fichero = md5 guardado " +
                            md5 + " = " + param.GetValue("MD5_archivo_contratosPS_MT"));
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
                ss_pp.Update_Fecha_Inicio("Contratación", "contratos_PS_Portugal_MT", "2_Importar_Contratos_MT");

                strSql = "delete from contratos_ps_mt";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Importando contratos MT");
                Console.WriteLine("========================");

                ficheroLog.Add("Importando archivo " + archivo);

                System.IO.StreamReader file = new System.IO.StreamReader(archivo, System.Text.Encoding.GetEncoding(1252));
                while ((line = file.ReadLine()) != null)
                {
                    campos = line.Split(';');
                    
                    c = 1;
                    total_registros++;

                    if(total_registros > 1)
                    {

                        i++;
                        if (firstOnly)
                        {
                            sb.Append("replace into contratos_ps_mt");
                            sb.Append(" (CUPS, CONTREXT, VERSION, TESTCONT, FALTACON,");
                            sb.Append(" FPSERCON, FPREVBAJ, FBAJACON, CEMPTITU,");
                            sb.Append(" CLINNEG, CSEGMERC, CCONTRPS, CCLIENTE,");
                            sb.Append(" CCALENPO, FSIGFACT, FFINVESU, CTARIFA,");
                            sb.Append(" CUPSREE, CNIFDNIC, DAPERSOC");                            
                            sb.Append(") values ");
                            firstOnly = false;
                        }

                        sb.Append("('").Append(campos[c]).Append("',"); c++; // CUPS
                        sb.Append("'").Append(campos[c]).Append("',"); c++; // CONTREXT                        
                        sb.Append(CN(campos[c])).Append(","); c++; // VERSION
                        sb.Append(CN(campos[c])).Append(","); c++; // TESTCONT
                        sb.Append(CF(campos[c])).Append(","); c++; // FALTACON

                        sb.Append(CF(campos[c])).Append(","); c++; // FPSERCON
                        sb.Append(CF(campos[c])).Append(","); c++; // FPREVBAJ
                        sb.Append(CF(campos[c])).Append(","); c++; // FBAJACON
                        sb.Append(CN(campos[c])).Append(","); c++; // CEMPTITU

                        sb.Append(CN(campos[c])).Append(","); c++; // CLINNEG
                        sb.Append("'").Append(campos[c]).Append("',"); c++; // CSEGMERC
                        sb.Append("'").Append(campos[c]).Append("',"); c++; // CCONTRPS
                        sb.Append("'").Append(campos[c]).Append("',"); c++; // CCLIENTE

                        sb.Append(CN(campos[c])).Append(","); c++; // CCALENPO
                        sb.Append(CF(campos[c])).Append(","); c++; // FSIGFACT
                        sb.Append(CF(campos[c])).Append(","); c++; // FFINVESU
                        sb.Append("'").Append(campos[c]).Append("',"); c++; // CTARIFA

                        sb.Append("'").Append(campos[c]).Append("',"); c++; // CUPSREE
                        sb.Append("'").Append(utilidades.FuncionesTexto.RT(campos[c])).Append("',"); c++; // CNIFDNIC
                        sb.Append("'").Append(utilidades.FuncionesTexto.RT(campos[c])).Append("'"); c++; // DAPERSOC

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

                ss_pp.Update_Fecha_Fin("Contratación", "contratos_PS_Portugal_MT", "2_Importar_Contratos_MT");
            }
            catch (Exception ex)
            {
                ficheroLog.AddError("ImportarArchivoContratos: " + ex.Message);
                Console.WriteLine("ImportarArchivoContratos: " + ex.Message);
            }
        }

        private void ImportarArchivo_old(string archivo)
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "replace into contratos_ps_mt_hist select * from contratos_ps_mt";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "delete from contratos_ps_mt";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                ficheroLog.Add("Importando " + archivo);
                strSql = "LOAD DATA LOCAL INFILE '" + archivo.Replace(@"\", "\\\\")
                    + "' REPLACE INTO TABLE contratos_ps_mt"
                    + " FIELDS TERMINATED BY ';' LINES TERMINATED BY '\\n'"
                    + " (@K, @cups13, @ccontext, @testcon, @faltacon, @fprebaja, @fbaja, @empresa,"
                    + " @copns, @ccontrps, @cnumscps, @linea, @ccliente, @calenda, @dias, @fsigfactu,"
                    + " @ffinversu, @tarifa, @cups20, @ccompobj, @vnseghor, @valor1, @valor2, @valor3, @valor4) SET"
                    + " cups13 = @cups13,"
                    + " ccontext = @ccontext,"
                    + " testcon = @testcon,"
                    + " faltacon = concat(substr(@faltacon, 1, 4), '-', substr(@faltacon, 5, 2), '-', substr(@faltacon, 7, 2)),"
                    + " fprebaja = concat(substr(@fprebaja, 1, 4), '-', substr(@fprebaja, 5, 2), '-', substr(@fprebaja, 7, 2)),"
                    + " fbaja = concat(substr(@fbaja, 1, 4), '-', substr(@fbaja, 5, 2), '-', substr(@fbaja, 7, 2)),"
                    + " empresa = @empresa,"
                    + " copns = @copns,"
                    + " ccontrps = @ccontrps,"
                    + " cnumscps = @cnumscps,"
                    + " linea = @linea,"
                    + " ccliente = @ccliente,"
                    + " calenda = @calenda,"
                    + " dias = @dias,"
                    + " fsigfactu = concat(substr(@fsigfactu, 1, 4), '-', substr(@fsigfactu, 5, 2), '-', substr(@fsigfactu, 7, 2)),"
                    + " ffinversu = concat(substr(@ffinversu, 1, 4), '-', substr(@ffinversu, 5, 2), '-', substr(@ffinversu, 7, 2)),"
                    + " tarifa = @tarifa,"
                    + " cups20 = @cups20,"
                    + " ccompobj = @ccompobj,"
                    + " vnseghor = @vnseghor,"
                    + " valor1 = @valor1,"
                    + " valor2 = @valor2,"
                    + " valor3 = @valor3,"
                    + " valor4 = @valor4;";

                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "update contratos_ps_mt set ffinversu = null where ffinversu = '0000-00-00'";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                ficheroLog.Add("Fin Importación");

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Importar: " + e.Message);
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

        public string GetCatalogo(string cups20)
        {
            string catalogo = "N/A";
            EndesaEntity.facturacion.InventarioTipologias o;
            EndesaEntity.contratacion.ComplementosContrato oo;
            if (inventario.TryGetValue(cups20, out o))
            {
                if (o.dic_complementos.Count > 0)
                {
                    List<EndesaEntity.contratacion.ComplementosContrato> lista =
                        o.dic_complementos.Values.ToList();

                    bool tiene_L77 = lista.Exists(z => z.ccompobj == "L77");
                    bool tiene_A01 = lista.Exists(z => z.ccompobj == "A01");

                    EndesaEntity.contratacion.ComplementosContrato c =
                        lista.Find(z => z.ccompobj == "L01");

                    if (c != null)
                    {
                        if (c.vparam[1] > 0)
                            catalogo = "Catálogo MT";
                        else if (c.vparam[1] == 0 && !tiene_L77)
                            catalogo = "Personalizado con revisión específica";
                        else if (c.vparam[1] == 0 && tiene_L77)
                            catalogo = "Pesonalizado con revisión estándar";
                    } else if (tiene_A01)
                        catalogo = "Flexible";
                }

            }

            return catalogo;
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

