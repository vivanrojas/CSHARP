using EndesaBusiness.contratacion.eexxi;
using EndesaBusiness.servidores;
using EndesaEntity.factoring;
using Microsoft.Graph.SecurityNamespace;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.factoring
{
    public class Prevision
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "mes13_prevision");
        utilidades.Param param;
        CalculoEstimacion cal_estimacion;
        public Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic { get; set; }
        public Prevision() 
        {
            dic = new Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>>();
            param = new utilidades.Param("ff_param", MySQLDB.Esquemas.FAC);
        }

        public void CreaPrevision(string fichero)
        {
            double percent = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            cal_estimacion = new CalculoEstimacion(dic);
            pb.progressBar.Minimum = 0;
            pb.progressBar.Maximum = 100;
            pb.progressBar.Step = 1;
            

            pb.Show();
            foreach (KeyValuePair<string, List<EndesaEntity.factoring.CalendarioFactoring>> p in dic)
            {

                pb.Text = "Calculando previsión " + p.Key;

                for (int i = 0; i < p.Value.Count; i++)
                {
                    percent = ((i + 1) / Convert.ToDouble(p.Value.Count)) * 100;
                    pb.txtDescripcion.Text = "Calculando bloque: "
                        + string.Format("{0}", (i).ToString("#,##0"))
                        + " de "
                        + string.Format("{0}", 3);
                    pb.progressBar.Value = Convert.ToInt32(percent);
                    pb.progressBar.Value = Convert.ToInt32(percent) - 1;
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();


                    BuscarFacturas(p.Key, p.Value[i].bloque,
                        p.Value[i].facturas_desde, p.Value[i].facturas_hasta,
                        p.Value[i].consumos_desde, p.Value[i].consumos_hasta,
                        p.Value[i].importe_min_factura, p.Value[i].importe_min_factura_agrupada);


                    cal_estimacion.GuardaImporteHidrocarburo(Convert.ToInt32(p.Key), p.Value[i].bloque, "INDIVIDUALES");
                    cal_estimacion.GuardaImporteHidrocarburo(Convert.ToInt32(p.Key), p.Value[i].bloque, "AGRUPADAS");
                }

                // utilizamos el bloque 1 como base para establecer que clientes con facturación tanto en individuales como en agrupadas
                // tiene un importe mayor
                //if (!sacar_solo_excel)
                //    sf.Quita_Grupo_Menor(factoring);

                 cal_estimacion.CalculaEstimacion();

                // [GUS 19/01/2024] Añadimos nueva llamada a esta nueva función que compara los importes de las estimaciones agrupadas con las individuales
                // eliminando las individuales de un NIF en caso de sumar menor importe que la agrupada
                SeleccionaAgrupadasVSIndividuales();
               

                Informes inf = new Informes();
                inf.Informe_Estimacion(fichero, p.Key, DateTime.Now, cal_estimacion.dic_est_ind,
                    cal_estimacion.dic_est_agr, cal_estimacion.lista_negra.dic, p.Value);

                pb.Close();
            }

        }
        // [GUS 19/01/2024] Nueva función que compara los importes de las estimaciones agrupadas con las individuales
        // eliminando las individuales con un NIF en caso de sumar menor importe que la agrupada
        private void SeleccionaAgrupadasVSIndividuales()
        {
            double numVeces;
            double importe_estimacion_agrupada;
            double importe_estimacion_individuales;
            double importe_estimacion_individuales_tmp;
            bool existe_nif_individuales;
            //Creamos un diccionario donde guardamos las key de las estimaciones individuales para borrar si procedemos
            Dictionary<string, string> dic_key_enc_dic_inv = new Dictionary<string, string>();

            foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> p in cal_estimacion.dic_est_agr)
            {
                //Borramos claves y valores
                dic_key_enc_dic_inv.Clear();

                importe_estimacion_agrupada = 0;
                //Para cada registro de agrupadas calculamos el importe de la estimación agrupada
                numVeces = 0;
                for (int i = 0; i < 4; i++)
                {
                    if (p.Value.ifactura[i] > 0)
                    {

                        if (i == 0)
                        {
                            if (p.Value.ifactura[i] > 0)
                                numVeces = 2;
                            importe_estimacion_agrupada = importe_estimacion_agrupada + (p.Value.ifactura[i] + p.Value.ifactura[i]);
                        }
                        else
                        {
                            if (p.Value.ifactura[i] > 0)
                                numVeces++;
                            importe_estimacion_agrupada = importe_estimacion_agrupada + p.Value.ifactura[i];
                        }

                    }
                }

                if (numVeces > 0)
                    importe_estimacion_agrupada = Math.Round((importe_estimacion_agrupada / numVeces), 2);

                // Ya tenemos el importe total de la estimación para la fatura agrupada en curso en importe_estimacion_agrupada
                // Ahora procedemos a buscar el NIf en las individuales y calcular su suma  
                existe_nif_individuales = false;
                importe_estimacion_individuales = 0;
                
                // Para cada registro en individuales buscamos NIF y calculamos suma estimaciones individuales
                foreach (KeyValuePair<string, EndesaEntity.factoring.Estimacion> pp in cal_estimacion.dic_est_ind)
                {

                    if (pp.Value.nif == p.Value.nif)
                    {
                        existe_nif_individuales = true;
                        importe_estimacion_individuales_tmp = 0;
                        //Si encontramos el NIF actualizamos importe_estimacion_individuales
                        // Guardamos key en diccionario
                        dic_key_enc_dic_inv.Add(pp.Key, pp.Key);

                        numVeces = 0;
                        for (int i = 0; i < 4; i++)
                        {
                            if (pp.Value.ifactura[i] > 0)
                            {

                                if (i == 0)
                                {
                                    if (pp.Value.ifactura[i] > 0)
                                        numVeces = 2;
                                    importe_estimacion_individuales_tmp = importe_estimacion_individuales_tmp + (pp.Value.ifactura[i] + pp.Value.ifactura[i]);
                                }
                                else
                                {
                                    if (pp.Value.ifactura[i] > 0)
                                        numVeces++;
                                    importe_estimacion_individuales_tmp = importe_estimacion_individuales_tmp + pp.Value.ifactura[i];
                                }

                            }
                        }

                        if (numVeces > 0)
                            importe_estimacion_individuales = importe_estimacion_individuales + Math.Round((importe_estimacion_individuales_tmp / numVeces), 2);

                    }
                }

                //Hemos encontrado el NIf de agrupada en individuales?
                //SI --> comparamos importes para determinar si nos quedamos con las agrupadas y borramos las individuales,
                //en caso contrario el proceso actual descartará la agrupada (ver EndesaBusiness.factoring.Informes líneas 408-416)
                if (existe_nif_individuales)
                {
                    if (importe_estimacion_agrupada > importe_estimacion_individuales)
                    {
                        foreach (KeyValuePair<string, string> p_enc in dic_key_enc_dic_inv)
                        {
                            cal_estimacion.dic_est_ind.Remove(p_enc.Value);
                        }
                    }
                }

            }

        }

        private void BuscarFacturas(string factoring, int bloque,
            DateTime f_factura_des, DateTime f_factura_has,
            DateTime ffactdes, DateTime ffacthas,
            double importeMinimoFactura, double importeMinimoFacturaAgrupada)
        {
            AnexionDatos(factoring, bloque, f_factura_des, f_factura_has, ffactdes, ffacthas);
            MarcaCisternas(factoring, bloque);
            //Mantenemos este borrado para las facturas procedentes de SCE con antigua codificación
            BorradoFacturas(factoring, bloque, "A");
            BorradoFacturas(factoring, bloque, "S");

            // [05/02/2025] GUS: Añadimos nueva función: eliminar facturas que tengan en su código fiscal los caracteres 'FC' en posición 4 y 5 respectivamente
            BorrarFacturasPorCodigoFiscal(factoring, bloque);


            AgruparFacturas_PerdiodosPartidos(factoring, bloque, ffactdes, ffacthas);
            BorrarFacturasPorImporte(factoring, bloque, importeMinimoFactura);
            BorrarFacturasAgrupadasPorImporte(factoring, bloque, importeMinimoFacturaAgrupada);
            
            UltimaFactura(factoring, bloque);
            MoverFacturasDeTmp(factoring, bloque);
        }

        private void AnexionDatos(string factoring, int bloque,
            DateTime f_factura_des, DateTime f_factura_has,
            DateTime ffactdes, DateTime ffacthas)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                // Borramos tablas porque se trata de la primera vez                
                if (bloque == 0)
                {
                    // Borramos la tabla de facturas tmp del factoring
                    Console.WriteLine("Borramos la tabla de facturas temporal");
                    strSql = "delete from ff_facturas_all_tmp where" +
                        " factoring = " + factoring ;
                    ficheroLog.Add(strSql);
                    Console.WriteLine(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    // Borramos la tabla de facturas
                    Console.WriteLine("Borramos la tabla de facturas");
                    strSql = "delete from ff_facturas_all where" +
                        " factoring = " + factoring;
                    ficheroLog.Add(strSql);
                    Console.WriteLine(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    // Borramos la tabla de facturas
                    Console.WriteLine("Borramos la tabla de facturas estimacion");
                    strSql = "delete from ff_estimacion where" +
                        " factoring = " + factoring;
                    ficheroLog.Add(strSql);
                    Console.WriteLine(strSql);
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                }
                // Borramos la tabla de facturas tmp
                Console.WriteLine("Borramos la tabla de facturas temporal");
                strSql = "delete from ff_facturas_all_tmp where" +
                    " factoring = " + factoring + " and" +
                    " bloque = " + bloque;
                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //10/10/2024 GUS: Anexamos facturas individuales de SIGAME 
                //06/02/2025 GUS: Movemos la obtención de las facturas individuales de GAS al primer lugar por si en SAP no está alguna y si lo está que sobreescriba datos
                //22/05/2025 GUS: Ajustamos la fecha hasta de la factura (f_factura_has) añadiéndole los días especificados por el parámetro dias_dilatacion
                Console.WriteLine("Inicio anexión facturas individuales clientes SIGAME - " + " Factoring " + factoring + " Bloque " + bloque);
                AnexaFacturasSIGAME(factoring, bloque, f_factura_des, f_factura_has.AddDays(Convert.ToInt32(param.GetValue("dias_dilatacion", DateTime.Now, DateTime.Now))), ffactdes, ffacthas);
                Console.WriteLine("Fin anexión facturas individuales clientes SIGAME -" + " Factoring " + factoring + " Bloque " + bloque);


                //24/05/2024 GUS: Primero Anexamos facturas individuales  y agrupadas de SAP - CLIENTES MIGRADOS
                //22/05/2025 GUS: Ajustamos la fecha hasta de la factura (f_factura_has) añadiéndole los días especificados por el parámetro dias_dilatacion
                Console.WriteLine("Inicio anexión facturas individuales clientes SAP - " + " Factoring " + factoring + " Bloque " + bloque);
                AnexaFacturasSAP(factoring, bloque, f_factura_des, f_factura_has.AddDays(Convert.ToInt32(param.GetValue("dias_dilatacion", DateTime.Now, DateTime.Now))), ffactdes, ffacthas, "IND");
                Console.WriteLine("Fin anexión facturas individuales clientes SAP - " + " Factoring " + factoring + " Bloque " + bloque);


                //15/10/2024 GUS: Deshabilitamos la obtención de facturas agrupadas SD de la tabla sap_tfactura_n0 (vemos que se incluyen en sap_facts) -
                // Volvemos a habilitar al no encontrar todos los clientes, ubicamos en primer lugar para que posteriormente prevalezca los datos de t_ed_h_sap_facts
                Console.WriteLine("Inicio anexión facturas agrupadas clientes SAP N0 SD - " + " Factoring " + factoring + " Bloque " + bloque);
                AnexaFacturasSAP(factoring, bloque, f_factura_des, f_factura_has, ffactdes, ffacthas, "AGR_N0_SD");
                Console.WriteLine("Fin anexión facturas agrupadas clientes SAP N0 SD - " + " Factoring " + factoring + " Bloque " + bloque);

                Console.WriteLine("Inicio anexión facturas agrupadas clientes SAP - " + " Factoring " + factoring + " Bloque " + bloque);
                AnexaFacturasSAP(factoring, bloque, f_factura_des, f_factura_has, ffactdes, ffacthas, "AGR");
                Console.WriteLine("Fin anexión facturas agrupadas clientes SAP - " + " Factoring " + factoring + " Bloque " + bloque);

                Console.WriteLine("Inicio anexión facturas agrupadas clientes SAP - sin cruce con PS y solo España " + " Factoring " + factoring + " Bloque " + bloque);
                AnexaFacturasSAP(factoring, bloque, f_factura_des, f_factura_has, ffactdes, ffacthas, "AGR_no_PS");
                Console.WriteLine("Fin anexión facturas agrupadas clientes SAP - - sin cruce con PS y solo España" + " Factoring " + factoring + " Bloque " + bloque);

                

               
                

                //24/05/2024 GUS: Ahora incluimos las de SCE que no estén en la tabla ff_facturas_all_tmp por código factura
                Console.WriteLine("Facturas con consumos desde el "
                    + ffactdes.ToString("dd/MM/yyyy") + " hasta el "
                    + ffacthas.ToString("dd/MM/yyyy"));

                //strSql = "replace into ff_facturas_all_tmp select"
                //    + " " + factoring + ", " + bloque + ","
                //    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                //    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                //    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                //    + " from fo where"
                //    + " (FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                //    + " and FFACTURA <= '" +
                //    f_factura_has.AddDays(Convert.ToInt32(param.GetValue("dias_dilatacion", DateTime.Now, DateTime.Now))).ToString("yyyy-MM-dd") + "')"
                //    + " and TFACTURA in (1,2);";
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "replace into ff_facturas_all_tmp select"
                //    + " " + factoring + ", " + bloque + ","
                //    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                //    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                //    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                //    + " from fo where"
                //    + " (FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                //    + " and FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "')"
                //    + " and TFACTURA in (5,6);";
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();


                strSql = "replace into ff_facturas_all_tmp select"
                    + " " + factoring + ", " + bloque + ","
                    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                    + " from fo where"
                    + " (FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "'"
                    + " and FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "') AND"
                    + " (FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                    + " and FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "')"
                    + " and TFACTURA in (1,2) and CFACTURA NOT IN (select t.CFACTURA from ff_facturas_all_tmp t where t.factoring="+ factoring +" and t.CFACTURA IS NOT NULL);";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

               

                strSql = "replace into ff_facturas_all_tmp select"
                    + " " + factoring + ", " + bloque + ","
                    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                    + " from fo where"
                    + " (FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                    + " and FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "')"
                    + " and TFACTURA in (5,6) and CFACTURA NOT IN (select t.CFACTURA from ff_facturas_all_tmp t where t.factoring=" + factoring + " and t.CFACTURA IS NOT NULL);";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                // Adjuntamos las facturas de Holanda

                //strSql = "replace into ff_facturas_all_tmp select"
                //   + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                //   + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                //   + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                //   + " from fo f"
                //   + " where"
                //   + " f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and"
                //   + " f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "' and"
                //   + " f.CNIFDNIC like 'NL%'";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //// Adjuntamos facturas SAP SPAIN
                //Console.WriteLine("Adjuntamos facturas emitidas En SAP");
                //strSql = "replace into ff_facturas_all_tmp select "
                //    + " " + factoring + ", " + bloque + ","
                //    + " NULL as CCOUNIPS, 20, f.cd_cuenta_contr AS CREFEREN, 0 AS SECFACTU,"
                //    + " f.id_fact AS CFACTURA,"
                //    + " f.fh_fact AS FFACTURA, f.fh_ini_fact AS FFACTDES, f.fh_fin_fact AS FFACTHAS,"
                //    + " f.im_factdo_con_iva AS IFACTURA, f.nm_total_iva AS IVA,"
                //    + " f.im_impuesto_2 AS IIMPUES2, f.im_impuesto_3 AS IIMPUES3,"
                //    + " NULL AS IBASEISE, NULL AS ISE, 1 as TFACTURA,"
                //    + " f.cd_est_fact AS TESTFACT, c.tx_apell_cli AS DAPERSOC, c.cd_nif_cif_cli AS CNIFDNIC,"
                //    + " 'OPERACIONES B2B' AS INDEMPRE, f.cd_cups_ext AS CUPSREE, NULL AS CFACTREC,"
                //    + " 1 AS CLINNEG, 'N' AS CSEGMERC, 0 AS NUMLABOR, 'L' AS TIPONEGOCIO, null"
                //    + " FROM fact.t_ed_h_sap_facts f"
                //    + " INNER JOIN cont.t_ed_h_ps c ON c.cups20 = SUBSTR(f.cd_cups_ext,1,20)"                    
                //    + " WHERE f.id_fact IS NOT NULL AND"                     
                //     + " (f.fh_fact >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and"
                //    + " f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "') and"
                //    + " (f.fh_ini_fact >= '" + ffactdes.ToString("yyyy-MM-dd") + "' and"
                //    + " f.fh_fin_fact <= '" + ffacthas.ToString("yyyy-MM-dd") + "')";
                //Console.WriteLine(strSql);
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.GBL);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //// Adjuntamos facturas SAP PORTUGAL
                //Console.WriteLine("Adjuntamos facturas emitidas En SAP");
                //strSql = "replace into ff_facturas_all_tmp select "
                //    + " " + factoring + ", " + bloque + ","
                //    + " NULL as CCOUNIPS, 20, f.cd_cuenta_contr AS CREFEREN, 0 AS SECFACTU,"
                //    + " f.id_fact AS CFACTURA,"
                //    + " f.fh_fact AS FFACTURA, f.fh_ini_fact AS FFACTDES, f.fh_fin_fact AS FFACTHAS,"
                //    + " f.im_factdo_con_iva AS IFACTURA, f.nm_total_iva AS IVA,"
                //    + " f.im_impuesto_2 AS IIMPUES2, f.im_impuesto_3 AS IIMPUES3,"
                //    + " NULL AS IBASEISE, NULL AS ISE, 1 as TFACTURA,"
                //    + " f.cd_est_fact AS TESTFACT, c.tx_apell_cli AS DAPERSOC, c.cd_nif_cif_cli AS CNIFDNIC,"
                //    + " 'OPERACIONES B2B' AS INDEMPRE, f.cd_cups_ext AS CUPSREE, NULL AS CFACTREC,"
                //    + " 1 AS CLINNEG, 'N' AS CSEGMERC, 0 AS NUMLABOR, 'L' AS TIPONEGOCIO, null"
                //    + " FROM fact.t_ed_h_sap_facts f"
                //    + " INNER JOIN cont.t_ed_h_ps_pt c ON c.cups20 = f.cd_cups_ext"
                //    + " WHERE f.id_fact IS NOT NULL AND"
                //     + " (f.fh_fact >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and"
                //    + " f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "') and"
                //    + " (f.fh_ini_fact >= '" + ffactdes.ToString("yyyy-MM-dd") + "' and"
                //    + " f.fh_fin_fact <= '" + ffacthas.ToString("yyyy-MM-dd") + "')"
                //    + " AND c.cd_tp_tension IN ('MT')";
                //Console.WriteLine(strSql);
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.GBL);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();




                // Adjuntamos facturas emitidas el mismo mes del periodo de consumo.
                Console.WriteLine("Adjuntamos facturas emitidas el mismo mes del periodo de consumo.");
                strSql = "replace into ff_facturas_all_tmp select "
                    + " " + factoring + ", " + bloque + ","
                    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                    + " from fo f where"
                    + " (f.FFACTURA >= '" + ffactdes.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + ffacthas.ToString("yyyy-MM-dd") + "') and"
                    + " (f.FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "') and"
                    + " f.TFACTURA in (1,2,5,6) and CFACTURA NOT IN (select t.CFACTURA from ff_facturas_all_tmp t where t.factoring=" + factoring + " and t.CFACTURA IS NOT NULL);";
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                // Adjuntamos facturas de GAS
                // Modificado 2023-07-12
                //Console.WriteLine("Adjuntamos facturas de GAS");
                //strSql = "replace into ff_facturas_all_tmp select"
                //    + " " + factoring + ", " + bloque + ","
                //    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                //    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                //    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                //    + " from fo f"
                //    + " where f.CEMPTITU <> 70 and"
                //    + " f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and"
                //    + " f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "' and"
                //    + " f.TIPONEGOCIO = 'G';";
                //Console.WriteLine(strSql);
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();


                //Adjuntamos facturas de GAS
                // Modificado 2023 - 07 - 12
                Console.WriteLine("Adjuntamos facturas de GAS");
                strSql = "replace into ff_facturas_all_tmp select"
                    + " " + factoring + ", " + bloque + ","
                    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,"
                    + " IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                    + " from fo f"
                    + " where f.CEMPTITU <> 70 and"
                    + " (FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "'"
                    + " and FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "') and"
                    + " f.TIPONEGOCIO = 'G' and CFACTURA NOT IN (select t.CFACTURA from ff_facturas_all_tmp t where t.factoring=" + factoring + " and t.CFACTURA IS NOT NULL);";
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                // Modificado 2023-07-12
                //Console.WriteLine("Adjuntamos facturas de GAS");
                //strSql = "replace into ff_facturas_all_tmp select"
                //    + " " + factoring + ", " + bloque + ","
                //    + " CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA," 
                //    + " if(MONTH(FFACTURA) = MONTH(FFACTDES), DATE_ADD(FFACTURA, INTERVAL 1 month), FFACTURA ) FFACTURA, FFACTDES, FFACTHAS,"
                //    + " SUM(IFACTURA) as IFACTURA, SUM(IVA) AS IVA,"
                //    + " SUM(IIMPUES2) as IIMPUES2, SUM(IIMPUES3) IIMPUES3, SUM(IBASEISE) IBASEISE, SUM(ISE) AS ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,"
                //    + " CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, null"
                //    + " from fo f"
                //    + " where f.CEMPTITU <> 70 and"
                //    + " f.FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "' and"
                //    + " f.FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "' and"
                //    + " f.TIPONEGOCIO = 'G'"
                //    + " GROUP BY CUPSREE, FFACTDES";
                //Console.WriteLine(strSql);
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();



                // Cambiamos el tipo de las facturas de GAS por tipo 1
                Console.WriteLine("Cambiamos el tipo de las facturas de GAS por tipo 1");
                strSql = "update ff_facturas_all_tmp set tfactura = 1 where TIPONEGOCIO = 'G';";
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //// Quitamos las facturas de consumos del mismo mes de emisión de facturas
                //strSql = "replace into ff_facturas_all_tmp_eliminadas select *, " +
                //    "'Fecha Factura = Periodo' from ff_facturas_all_tmp where" +
                //    " FFACTDES >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and" +
                //    " factoring = " + factoring + " and bloque = " + bloque;
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();


                //Console.WriteLine("Quitamos las facturas de consumos del mismo mes de emisión de facturas");
                //strSql = "delete from ff_facturas_all_tmp where" +
                //    " FFACTDES >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and" +
                //    " factoring = " + factoring + " and bloque = " + bloque;
                //db = new MySQLDB(MySQLDB.Esquemas.FAC); ;
                //Console.WriteLine(strSql);
                //ficheroLog.Add(strSql);
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                // Quitamos las facturas de EEXXI
                Console.WriteLine("Quitamos las facturas de EEXXI");
                strSql = "delete from ff_facturas_all_tmp where CEMPTITU = 70 and" +
                    " factoring = " + factoring + " and bloque = " + bloque; ;
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //Quitamos las facturas de BTN
                Console.WriteLine("Quitamos las facturas de BTN");
                strSql = "delete from ff_facturas_all_tmp where CEMPTITU = 20 and INDEMPRE = 'CEFACO' and" +
                    " factoring = " + factoring + " and bloque = " + bloque;
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de BTE
                Console.WriteLine("Quitamos las facturas de BTE");
                strSql = "delete from ff_facturas_all_tmp where CEMPTITU = 80 and INDEMPRE = 'CEFACO' and" +
                    " factoring = " + factoring + " and bloque = " + bloque;
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Borramos las facturas de consumos que no correspondan con las fechas que queramos
                Console.WriteLine("Borramos las facturas de consumos que no correspondan con las fechas que queramos");
                strSql = "delete from ff_facturas_all_tmp where ffacthas < '" + ffactdes.ToString("yyyy-MM-dd") + "' and" +
                    " factoring = " + factoring + " and bloque = " + bloque;
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SeleccionaFacturas.AnexionDatos " + e.Message);
            }
        }

        private void MarcaCisternas(string factoring, int bloque)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                strSql = "update ff_facturas_all_tmp f"
                    + " left outer join fo_s s on"
                    + " s.CREFEREN = f.CREFEREN and"
                    + " s.SECFACTU = f.SECFACTU"
                    + " left outer join cm_inventario_gas i on"
                    + " i.ID_PS = s.ID_PS and"
                    + " (i.FINICIO <= f.FFACTDES and (i.FFIN >= f.FFACTHAS or i.FFIN is null)) and"
                    + " i.Grupo = 'Cisterna'"
                    + " set f.CUPSREE = concat('Cisterna_', i.ID_PS),"
                    + " f.CCOUNIPS = concat('Cisterna_', i.ID_PS)"
                    + " where"
                    + " f.factoring = " + factoring + " and"
                    + " f.bloque = " + bloque + " and"
                    + " f.TIPONEGOCIO = 'G' and"
                    + " (f.CUPSREE is null or LENGTH(f.CUPSREE) < 20);";


                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                strSql = "update ff_facturas_all_tmp f"
                    + " left outer join fo s on"
                    + " s.CEMPTITU = f.CEMPTITU and"
                    + " s.CREFEREN = f.CREFEREN and"
                    + " s.SECFACTU = f.SECFACTU"
                    + " set f.CUPSREE = s.CUPSREE,"
                    + " f.CCOUNIPS = s.CCOUNIPS"
                    + " where"
                    + " f.factoring = " + factoring + " and"
                    + " f.bloque = " + bloque + " and"
                    + " f.TIPONEGOCIO = 'G' and"
                    + " (f.CUPSREE is null or LENGTH(f.CUPSREE) < 20);";

                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SeleccionaFacturas.MarcaCisternas " + e.Message);
            }
        }

        public void BorradoFacturas(string factoring, int bloque, string testfact)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                Console.WriteLine("Borrando facturas de tipo " + testfact);
                strSql = "delete from fact.ff_facturas_all_tmp where"
                    + " factoring = " + factoring + " and"
                    + " bloque = " + bloque + " and"
                    + " TESTFACT = '" + testfact + "';";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SeleccionaFacturas.BorradoFacturas " + e.Message);
            }
        }

        private void AgruparFacturas_PerdiodosPartidos(string factoring, int bloque, DateTime desdeFechaConsumo, DateTime hastaFechaConsumo)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            string ccounips = "";

            List<EndesaEntity.factoring.Factura> lista_facturas_parte1 = new List<EndesaEntity.factoring.Factura>();
            List<EndesaEntity.factoring.Factura> lista_facturas_parte2 = new List<EndesaEntity.factoring.Factura>();

            List<EndesaEntity.factoring.Factura> lista_facturas = new List<EndesaEntity.factoring.Factura>();
            List<string> lista_facturas_borrar = new List<string>();

            int i = 0;

            try
            {

                Console.WriteLine("Borramos la tabla ff_facturas_pp_tmp");
                strSql = "delete from ff_facturas_pp_tmp;";
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                lista_facturas_parte1 = FacturasParticionadas_bloque1(factoring, bloque, desdeFechaConsumo, hastaFechaConsumo);
                lista_facturas_parte2 = FacturasParticionadas_bloque2(factoring, bloque, desdeFechaConsumo, hastaFechaConsumo);


                foreach (EndesaEntity.factoring.Factura p in lista_facturas_parte1)
                {
                    EndesaEntity.factoring.Factura factura_tmp = lista_facturas_parte2.Find(z => z.cups20 == p.cups20);
                    if (factura_tmp != null)
                    {
                        EndesaEntity.factoring.Factura c = p;
                        c.iva = c.iva + factura_tmp.iva;
                        c.iimpue2 = c.iimpue2 + factura_tmp.iimpue2;
                        c.iimpue3 = c.iimpue3 + factura_tmp.iimpue3;
                        c.ibaseise = c.ibaseise + factura_tmp.ibaseise;
                        c.ise = c.ise + factura_tmp.ise;
                        c.ifactura = c.ifactura + factura_tmp.ifactura;
                        c.ffacthas = factura_tmp.ffacthas;

                        lista_facturas.Add(c);
                        lista_facturas_borrar.Add(factura_tmp.cfactura);
                    }
                    else
                    {
                        //Console.WriteLine("Factura no encontrada");
                    }

                }

                BorraGrupoFacturas(factoring, lista_facturas_borrar);
                GuardaFactura(factoring, bloque, lista_facturas);


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SeleccionaFacturas.AgruparFacturas_PerdiodosPartidos " + ccounips +
                    " " + e.Message);
            }
        }

        private void BorrarFacturasPorImporte(string factoring, int bloque, double importe)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            strSql = "replace into ff_facturas_all_tmp_eliminadas select f.*, 'Importe inferior a " + importe + "'"
                + " from ff_facturas_all_tmp f where f.IFACTURA < " + importe + " and"
                + " f.factoring = " + factoring + " and"
                + " f.bloque = " + bloque;
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
            Console.WriteLine("Borrando facturas de ff_facturas_all_tmp inferiores a " + string.Format("{0:#,##0}", importe));
            strSql = "delete from ff_facturas_all_tmp where IFACTURA < " + importe + " and"
                + " factoring = " + factoring + " and"
                + " bloque = " + bloque;
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        private void BorrarFacturasAgrupadasPorImporte(string factoring, int bloque, double importe)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;


            Console.WriteLine("Borrando facturas de clientes que no superen los " + string.Format("{0:#,##0}", importe));
            Console.WriteLine("delete from ff_nif_tmp");
            strSql = "delete from ff_nif_tmp;";
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "replace into ff_nif_tmp select f.CNIFDNIC, f.TIPONEGOCIO, sum(f.IFACTURA) as total"
                + " from ff_facturas_all_tmp f where"
                + " f.factoring = " + factoring + " and"
                + " f.bloque = " + bloque
                // Ahora el importe total por NIF se tiene en cuenta con la suma de GAS y ELECTRICIDAD
                //+ " group by f.CNIFDNIC, f.TIPONEGOCIO having total < " + importe;
                + " group by f.CNIFDNIC < " + importe;
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "delete f"
                + " from ff_facturas_all_tmp f inner join"
                + " ff_nif_tmp r on"
                + " r.CNIFDNIC = f.CNIFDNIC where"
                + " f.factoring = " + factoring + " and"
                + " f.bloque = " + bloque;
            ficheroLog.Add(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        // Función para eliminar las facturas que incluyen el literal FC en las posiciones 4 y 5 del código fiscal (CFACTURA)
        // Petición de Alvaro Carbo a 04/02/2025
        private void BorrarFacturasPorCodigoFiscal(string factoring, int bloque)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            try
            {
                strSql = "replace into ff_facturas_all_tmp_eliminadas select f.*, 'Código fiscal incluye FC en posición 4 y 5'"
                    + " from ff_facturas_all_tmp f where SUBSTRING(f.CFACTURA,4,2) = 'FC' and"
                    + " f.factoring = " + factoring + " and"
                    + " f.bloque = " + bloque;
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                Console.WriteLine("Borrando facturas de ff_facturas_all_tmp cuyo código fiscal incluye FC en posición 4 y 5");
                strSql = "delete from ff_facturas_all_tmp where SUBSTRING(CFACTURA,4,2) = 'FC' and"
                    + " factoring = " + factoring + " and"
                    + " bloque = " + bloque;
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "BorradoFacturasPorCodigoFiscal",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }
        private void MoverFacturasDeTmp(string factoring, int bloque)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            // Movemos las facturas de trabajo a la tabla principal

            Console.WriteLine("Volcamos las facturas INDIVIDUALES de ff_facturas_all_tmp a la tabla ff_facturas_all");
            strSql = "replace into ff_facturas_all select "
                + factoring + " as factoring,"
                + bloque + " as bloque,"
                + "'INDIVIDUALES' as tipo,"
                + " f.CCOUNIPS, f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                + " f.IFACTURA, f.IVA, f.IIMPUES2, f.IIMPUES3, f.IBASEISE, f.ISE, f.TFACTURA, f.TESTFACT, f.DAPERSOC,"
                + " f.CNIFDNIC, f.INDEMPRE, f.CUPSREE, f.CFACTREC, f.CLINNEG, f.CSEGMERC, f.NUMLABOR, f.TIPONEGOCIO,"
                + " f.COMENTARIOS, null from ff_facturas_all_tmp f where (f.TFACTURA <> 5 and f.TFACTURA <> 6) and"
                + " f.factoring = " + factoring + " and"
                + " f.bloque = " + bloque;
            Console.WriteLine(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            Console.WriteLine("Volcamos las facturas AGRUPADAS de ff_facturas_all_tmp a la tabla ff_facturas_all");
            strSql = "replace into ff_facturas_all select "
                + factoring + " as factoring,"
                + bloque + " as bloque,"
                + "'AGRUPADAS' as tipo,"
                + " f.CCOUNIPS, f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                + " f.IFACTURA, f.IVA, f.IIMPUES2, f.IIMPUES3, f.IBASEISE, f.ISE, f.TFACTURA, f.TESTFACT, f.DAPERSOC,"
                + " f.CNIFDNIC, f.INDEMPRE, f.CUPSREE, f.CFACTREC, f.CLINNEG, f.CSEGMERC, f.NUMLABOR, f.TIPONEGOCIO,"
                + " f.COMENTARIOS, null from ff_facturas_all_tmp f where f.TFACTURA in (5,6) and"
                + " f.factoring = " + factoring + " and"
                + " f.bloque = " + bloque;
            Console.WriteLine(strSql);
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


        }

        private void GuardaFactura(string factoring, int bloque, MySqlDataReader r1, MySqlDataReader r2)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            double iva = 0;
            double importe = 0;

            try
            {


                #region Cabecera replace
                strSql = "replace into ff_facturas_all_tmp (factoring, bloque, CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS,"
                    + " IFACTURA, IVA, IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC,"
                    + " INDEMPRE, CUPSREE, CFACTREC, CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, COMENTARIOS) values";

                #endregion

                #region Campos

                strSql += "(" + factoring + "," + bloque + ",";
                if (r2["CCOUNIPS"] != System.DBNull.Value)
                    strSql += "'" + r2["CCOUNIPS"].ToString() + "', ";
                else
                    strSql += " (null,";

                if (r2["CEMPTITU"] != System.DBNull.Value)
                    strSql += r2["CEMPTITU"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["CREFEREN"] != System.DBNull.Value)
                    strSql += r2["CREFEREN"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["SECFACTU"] != System.DBNull.Value)
                    strSql += r2["SECFACTU"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["CFACTURA"] != System.DBNull.Value)
                    strSql += "'" + r2["CFACTURA"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["FFACTURA"] != System.DBNull.Value)
                    strSql += "'" + Convert.ToDateTime(r2["FFACTURA"]).ToString("yyyy-MM-dd") + "', ";
                else
                    strSql += " null,";

                if (r1["FFACTDES"] != System.DBNull.Value)
                    strSql += "'" + Convert.ToDateTime(r1["FFACTDES"]).ToString("yyyy-MM-dd") + "', ";
                else
                    strSql += " null,";

                if (r2["FFACTHAS"] != System.DBNull.Value)
                    strSql += "'" + Convert.ToDateTime(r2["FFACTHAS"]).ToString("yyyy-MM-dd") + "', ";
                else
                    strSql += " null,";

                if (r2["IFACTURA"] != System.DBNull.Value)
                {
                    double totalIFactura =
                        (r1["IFACTURA"] != System.DBNull.Value ? Convert.ToDouble(r1["IFACTURA"]) : 0) + Convert.ToDouble(r2["IFACTURA"]);
                    strSql += totalIFactura.ToString().Replace(",", ".") + ", ";
                }
                else
                    strSql += " null,";

                if (r2["IVA"] != System.DBNull.Value)
                {
                    iva = (r1["IVA"] != System.DBNull.Value ? Convert.ToDouble(r1["IVA"]) : 0) + Convert.ToDouble(r2["IVA"]);
                    strSql += iva.ToString().Replace(",", ".") + ", ";
                }
                else
                    strSql += " null,";

                if (r2["IIMPUES2"] != System.DBNull.Value)
                {
                    importe = (r1["IIMPUES2"] != System.DBNull.Value ? Convert.ToDouble(r1["IIMPUES2"]) : 0) + Convert.ToDouble(r2["IIMPUES2"]);
                    strSql += importe.ToString().Replace(",", ".") + ", ";
                }

                else
                    strSql += " null,";

                if (r2["IIMPUES3"] != System.DBNull.Value)
                {
                    importe = (r1["IIMPUES3"] != System.DBNull.Value ? Convert.ToDouble(r1["IIMPUES3"]) : 0) + Convert.ToDouble(r2["IIMPUES3"]);
                    strSql += importe.ToString().Replace(",", ".") + ", ";
                }
                else
                    strSql += " null,";

                if (r2["IBASEISE"] != System.DBNull.Value)
                {
                    importe = (r1["IBASEISE"] != System.DBNull.Value ? Convert.ToDouble(r1["IBASEISE"]) : 0) + Convert.ToDouble(r2["IBASEISE"]);
                    strSql += importe.ToString().Replace(",", ".") + ", ";
                }

                else
                    strSql += " null,";

                if (r2["ISE"] != System.DBNull.Value)
                {
                    importe = r1["ISE"] != System.DBNull.Value ? Convert.ToDouble(r1["ISE"]) : 0 + Convert.ToDouble(r2["ISE"]);
                    strSql += importe.ToString().Replace(",", ".") + ", ";
                }
                else
                    strSql += " null,";

                if (r2["TFACTURA"] != System.DBNull.Value)
                    strSql += r2["TFACTURA"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["TESTFACT"] != System.DBNull.Value)
                    strSql += "'" + r2["TESTFACT"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["DAPERSOC"] != System.DBNull.Value)
                    strSql += "'" + r2["DAPERSOC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CNIFDNIC"] != System.DBNull.Value)
                    strSql += "'" + r2["CNIFDNIC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["INDEMPRE"] != System.DBNull.Value)
                    strSql += "'" + r2["INDEMPRE"].ToString() + "', ";
                else
                    strSql += " null,";

                #endregion

                #region RestoCampos                

                if (r2["CUPSREE"] != System.DBNull.Value)
                    strSql += "'" + (r2["CUPSREE"].ToString().Length > 20 ? r2["CUPSREE"].ToString().Substring(0, 20) : r2["CUPSREE"].ToString()) + "', ";
                else
                    strSql += " null,";

                if (r2["CFACTREC"] != System.DBNull.Value)
                    strSql += "'" + r2["CFACTREC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["CLINNEG"] != System.DBNull.Value)
                    strSql += r2["CLINNEG"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["CSEGMERC"] != System.DBNull.Value)
                    strSql += "'" + r2["CSEGMERC"].ToString() + "', ";
                else
                    strSql += " null,";

                if (r2["NUMLABOR"] != System.DBNull.Value)
                    strSql += r2["NUMLABOR"].ToString() + ", ";
                else
                    strSql += " null,";

                if (r2["TIPONEGOCIO"] != System.DBNull.Value)
                    strSql += "'" + r2["TIPONEGOCIO"].ToString() + "', ";
                else
                    strSql += " null,";

                strSql += "'" + "FACTURA SUMADA POR PERIODOS PARTIDOS" + "');";

                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("GuardaFactura " + e.Message);
            }

        }

        private void GuardaFactura(string factoring, int bloque, List<EndesaEntity.factoring.Factura> lista_facturas)
        {
            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            try
            {
                totalRegistros = lista_facturas.Count();

                foreach (EndesaEntity.factoring.Factura p in lista_facturas)
                {

                    j++;
                    k++;

                    if (firstOnly)
                    {
                        sb = null;
                        sb = new StringBuilder();
                        sb.Append("REPLACE INTO ff_facturas_all_tmp");
                        sb.Append(" (factoring, bloque, CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS,");
                        sb.Append(" IFACTURA, IVA, IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC,");
                        sb.Append(" INDEMPRE, CUPSREE, CFACTREC, CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, COMENTARIOS) values");
                        firstOnly = false;
                    }

                    #region Campos
                    sb.Append("(").Append(factoring).Append(",");
                    sb.Append(bloque).Append(",");

                    if (p.cups13 != null)
                        sb.Append("'").Append(p.cups13).Append("',");
                    else
                        sb.Append("null,");

                    if (p.cemptitu != 0)
                        sb.Append(p.cemptitu).Append(",");
                    else
                        sb.Append("null,");

                    if (p.creferen != 0)
                        sb.Append(p.creferen).Append(",");
                    else
                        sb.Append("null,");

                    if (p.secfactu != 0)
                        sb.Append(p.secfactu).Append(",");
                    else
                        sb.Append("null,");

                    if (p.cfactura != "")
                        sb.Append("'").Append(p.cfactura).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ffactura > DateTime.MinValue)
                        sb.Append("'").Append(p.ffactura.ToString("yyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ffactdes > DateTime.MinValue)
                        sb.Append("'").Append(p.ffactdes.ToString("yyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ffacthas > DateTime.MinValue)
                        sb.Append("'").Append(p.ffacthas.ToString("yyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.ifactura != 0)
                        sb.Append(p.ifactura.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.iva != 0)
                        sb.Append(p.iva.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.iimpue2 != 0)
                        sb.Append(p.iimpue2.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.iimpue3 != 0)
                        sb.Append(p.iimpue3.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.ibaseise != 0)
                        sb.Append(p.ibaseise.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.ise != 0)
                        sb.Append(p.ise.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    sb.Append(p.tfactura).Append(",");
                    sb.Append("'").Append(p.testfact).Append("',");
                    sb.Append("'").Append(p.cliente).Append("',");
                    sb.Append("'").Append(p.nif).Append("',");

                    if (p.indempre != "")
                        sb.Append("'").Append(p.indempre).Append("',");
                    else
                        sb.Append("null,");

                    if (p.cups20 != "")
                        sb.Append("'").Append(p.cups20).Append("',");
                    else
                        sb.Append("null,");

                    if (p.cfactrec != "")
                        sb.Append("'").Append(p.cfactrec).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append(p.clinneg).Append(",");

                    if (p.csegmerc != "")
                        sb.Append("'").Append(p.csegmerc).Append("',");
                    else
                        sb.Append("null,");

                    if (p.numlabor != 0)
                        sb.Append(p.numlabor).Append(",");
                    else
                        sb.Append("null,");

                    if (p.tiponegocio != "")
                        sb.Append("'").Append(p.tiponegocio).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append("FACTURA SUMADA POR PERIODOS PARTIDOS").Append("'),");
                    #endregion


                    if (j == 250)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;
                    }

                }

                if (j > 0)
                {
                    Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                    firstOnly = true;
                    db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.AddError("GuardaFactura " + ex.Message);
            }
        }

        private void BorraGrupoFacturas(long creferen, int secfactu, int i, double importe)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                //Console.WriteLine(i + " Borrando factura con CREFEREN: " + creferen + " y SECFACTU: " + secfactu);
                strSql = "delete from ff_facturas_all_tmp where"
                    + " creferen = " + creferen + " and"
                    + " secfactu = " + secfactu + " and"
                    + " ifactura = " + importe.ToString().Replace(",", ".");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private void BorraGrupoFacturas(string factoring, List<string> lista_facturas)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                //Console.WriteLine(i + " Borrando factura con CREFEREN: " + creferen + " y SECFACTU: " + secfactu);
                strSql = "delete from ff_facturas_all_tmp where"
                   + " factoring = '" + factoring + "' AND"
                   + " cfactura in ('" + lista_facturas[0] + "'";

                for (int i = 1; i < lista_facturas.Count; i++)
                    strSql += ", '" + lista_facturas[i] + "'";


                strSql += ")";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }
     
        private void UltimaFactura(string factoring, int bloque)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                Console.WriteLine("Proceso para quedarnos con la última factura emitida");
                Console.WriteLine("Borrado de la tabla ff_refacturaciones");
                strSql = "delete from ff_refacturaciones;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Buscamos las refecturaciones.");
                strSql = "replace into ff_refacturaciones select f.CFACTREC"
                      + " from ff_facturas_all_tmp f where (f.CFACTREC is not null and f.CFACTREC <> '') and"
                      + " f.factoring = " + factoring + " and"
                      + " f.bloque = " + bloque;

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Borramos las facturas");
                strSql = "delete f"
                    + " from ff_facturas_all_tmp f"
                    + " inner join ff_refacturaciones f1 on"
                    + " f.CFACTURA = f1.CFACTREC;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message,
                //   "UltimaFactura",
                //   MessageBoxButtons.OK,
                //   MessageBoxIcon.Error);
            }
        }

        private List<EndesaEntity.factoring.Factura> FacturasParticionadas_bloque1(string factoring, int bloque, DateTime fechaConsumoDesde, DateTime fechaConsumoHasta)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.factoring.Factura> lista = new List<EndesaEntity.factoring.Factura>();

            try
            {

                strSql = "select f.factoring, f.bloque, f.CCOUNIPS, f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.IFACTURA, f.IVA, f.IIMPUES2, f.IIMPUES3, f.IBASEISE, f.ISE, f.TFACTURA, f.TESTFACT, f.DAPERSOC, f.CNIFDNIC, f.INDEMPRE,"
                    + " f.CUPSREE, f.CFACTREC, f.CLINNEG, f.CSEGMERC, f.NUMLABOR, f.TIPONEGOCIO, f.COMENTARIOS"
                    + " from ff_facturas_all_tmp f where"
                    + " f.FFACTHAS < '" + fechaConsumoHasta.ToString("yyyy-MM-dd") + "' and"
                    + " f.factoring = " + factoring + " and"
                    + " f.bloque = " + bloque;
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.factoring.Factura c = new EndesaEntity.factoring.Factura();
                    c.factoring = Convert.ToInt32(factoring);
                    c.bloque = bloque;

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.cups13 = r["CCOUNIPS"].ToString();

                    c.cemptitu = Convert.ToInt32(r["CEMPTITU"]);
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    c.cfactura = r["CFACTURA"].ToString();
                    c.ffactura = Convert.ToDateTime(r["FFACTURA"]);
                    c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);


                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    if (r["IVA"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["IVA"]);

                    if (r["IIMPUES2"] != System.DBNull.Value)
                        c.iimpue2 = Convert.ToDouble(r["IIMPUES2"]);

                    if (r["IIMPUES3"] != System.DBNull.Value)
                        c.iimpue3 = Convert.ToDouble(r["IIMPUES3"]);

                    if (r["IBASEISE"] != System.DBNull.Value)
                        c.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                    if (r["ISE"] != System.DBNull.Value)
                        c.ise = Convert.ToDouble(r["ISE"]);

                    c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.cliente = r["DAPERSOC"].ToString();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.nif = r["CNIFDNIC"].ToString();

                    if (r["INDEMPRE"] != System.DBNull.Value)
                        c.indempre = r["INDEMPRE"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cups20 = r["CUPSREE"].ToString();

                    if (r["CFACTREC"] != System.DBNull.Value)
                        c.cfactrec = r["CFACTREC"].ToString();

                    if (r["CLINNEG"] != System.DBNull.Value)
                        c.clinneg = Convert.ToInt32(r["CLINNEG"]);

                    if (r["CSEGMERC"] != System.DBNull.Value)
                        c.csegmerc = r["CSEGMERC"].ToString();

                    if (r["NUMLABOR"] != System.DBNull.Value)
                        c.numlabor = Convert.ToInt32(r["NUMLABOR"]);

                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.tiponegocio = r["TIPONEGOCIO"].ToString();

                    if (r["COMENTARIOS"] != System.DBNull.Value)
                        c.comentarios = r["COMENTARIOS"].ToString();

                    lista.Add(c);


                }
                db.CloseConnection();

                return lista;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

        }

        private List<EndesaEntity.factoring.Factura> FacturasParticionadas_bloque2(string factoring, int bloque, DateTime fechaConsumoDesde, DateTime fechaConsumoHasta)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.factoring.Factura> lista = new List<EndesaEntity.factoring.Factura>();

            try
            {

                strSql = "select f.factoring, f.bloque, f.CCOUNIPS, f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + " f.IFACTURA, f.IVA, f.IIMPUES2, f.IIMPUES3, f.IBASEISE, f.ISE, f.TFACTURA, f.TESTFACT, f.DAPERSOC, f.CNIFDNIC, f.INDEMPRE,"
                    + " f.CUPSREE, f.CFACTREC, f.CLINNEG, f.CSEGMERC, f.NUMLABOR, f.TIPONEGOCIO, f.COMENTARIOS"
                    + " from ff_facturas_all_tmp f where"
                    + " f.FFACTDES > '" + fechaConsumoDesde.ToString("yyyy-MM-dd") + "' and"
                    + " f.factoring = " + factoring + " and"
                    + " f.bloque = " + bloque;
                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.factoring.Factura c = new EndesaEntity.factoring.Factura();
                    c.factoring = Convert.ToInt32(factoring);
                    c.bloque = bloque;

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.cups13 = r["CCOUNIPS"].ToString();

                    c.cemptitu = Convert.ToInt32(r["CEMPTITU"]);
                    c.creferen = Convert.ToInt64(r["CREFEREN"]);
                    c.secfactu = Convert.ToInt32(r["SECFACTU"]);
                    c.cfactura = r["CFACTURA"].ToString();
                    c.ffactura = Convert.ToDateTime(r["FFACTURA"]);
                    c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);


                    if (r["IFACTURA"] != System.DBNull.Value)
                        c.ifactura = Convert.ToDouble(r["IFACTURA"]);

                    if (r["IVA"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["IVA"]);

                    if (r["IIMPUES2"] != System.DBNull.Value)
                        c.iimpue2 = Convert.ToDouble(r["IIMPUES2"]);

                    if (r["IIMPUES3"] != System.DBNull.Value)
                        c.iimpue3 = Convert.ToDouble(r["IIMPUES3"]);

                    if (r["IBASEISE"] != System.DBNull.Value)
                        c.ibaseise = Convert.ToDouble(r["IBASEISE"]);

                    if (r["ISE"] != System.DBNull.Value)
                        c.ise = Convert.ToDouble(r["ISE"]);

                    c.tfactura = Convert.ToInt32(r["TFACTURA"]);

                    if (r["TESTFACT"] != System.DBNull.Value)
                        c.testfact = r["TESTFACT"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.cliente = r["DAPERSOC"].ToString();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.nif = r["CNIFDNIC"].ToString();

                    if (r["INDEMPRE"] != System.DBNull.Value)
                        c.indempre = r["INDEMPRE"].ToString();

                    if (r["CUPSREE"] != System.DBNull.Value)
                        c.cups20 = r["CUPSREE"].ToString();

                    if (r["CFACTREC"] != System.DBNull.Value)
                        c.cfactrec = r["CFACTREC"].ToString();

                    if (r["CLINNEG"] != System.DBNull.Value)
                        c.clinneg = Convert.ToInt32(r["CLINNEG"]);

                    if (r["CSEGMERC"] != System.DBNull.Value)
                        c.csegmerc = r["CSEGMERC"].ToString();

                    if (r["NUMLABOR"] != System.DBNull.Value)
                        c.numlabor = Convert.ToInt32(r["NUMLABOR"]);

                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        c.tiponegocio = r["TIPONEGOCIO"].ToString();

                    if (r["COMENTARIOS"] != System.DBNull.Value)
                        c.comentarios = r["COMENTARIOS"].ToString();

                    lista.Add(c);


                }
                db.CloseConnection();

                return lista;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

        }

        // Función para anexar facturas a la tabla ff_facturas_all_tmp de los clientes migrados a SAP - tipo="IND" --> INDIVIDUALES o tipo="AGR" --> AGRUPADAS
        // f_factura_* --> Fecha factura
        // ffact* --> Fecha consumo 
        // Además añadimos las obligaciones (flimpago y fptacobr) a nueva tabla 13_obligaciones_sap
        public void AnexaFacturasSAP(string factoring, int bloque, DateTime f_factura_des, DateTime f_factura_has,
            DateTime ffactdes, DateTime ffacthas, string tipo)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sb_oblig = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            string strTMP = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;
        
            try
            {

                switch(tipo)
                {
                    case "IND":
                        //Obtenemos las facturas individuales de SAP
                        strSql = QueryFactIndSAP(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        break;
                    case "AGR":
                        //Obtenemos las facturas agrupadas de SAP
                        strSql = QueryFactAgrSAP(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        break;
                    case "AGR_no_PS":
                        //Obtenemos las facturas SD agrupadas de SAP, sin cruce con PS (sin CUPS)
                        strSql = QueryFactAgrSAP_no_PS(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        break;
                    case "AGR_N0_SD":
                        //Obtenemos las facturas SD agrupadas de SAP, obtenidas de la tabla ed_owner.sap_tfactura_n0
                        strSql = QueryFactAgrSAP_N0_SD(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        break; 
                    default:
                        //Por defecto obtenemos las facturas individuales de SAP
                        strSql = QueryFactIndSAP(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        break;
                }

                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                command.CommandTimeout = 2000;
                r = command.ExecuteReader();
                while (r.Read())
                {

                    j++;
                    k++;

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO ff_facturas_all_tmp");
                        sb.Append(" (factoring, bloque, CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,");
                        sb.Append(" IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,");
                        sb.Append(" CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, COMENTARIOS) values ");

                        sb_oblig.Append("REPLACE INTO 13_obligaciones_sap");
                        sb_oblig.Append(" (id_fact, flimpago, fptacobr) values  ");

                        firstOnly = false;
                    }

                    #region Campos
                    //factoring (INT)
                    sb.Append("('").Append(factoring).Append("',");
                    //bloque (TINYINT)
                    sb.Append("'").Append(bloque).Append("',");
                    //CCOUNIPS (CHAR)
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CCOUNIPS"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CEMPTITU (TINYINT)
                    if (r["CEMPTITU"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CEMPTITU"].ToString()).Append("',");
                    else
                        sb.Append("'0',");
                    //CREFEREN (BIGINT)
                    if (r["CREFEREN"] != System.DBNull.Value)
                    {
                        //sb.Append("'").Append(r["CREFEREN"].ToString()).Append("',");
                        strTMP = (r["CREFEREN"].ToString());
                        sb.Append("'").Append(strTMP.Where(c => char.IsDigit(c)).ToArray()).Append("',");
                    }
                    else
                        sb.Append("'0',");
                    //SECFACTU (SMALLINT)
                    if (r["SECFACTU"] != System.DBNull.Value)
                        sb.Append("'").Append(r["SECFACTU"].ToString()).Append("',");
                    else
                        sb.Append("'0',");
                    //CFACTURA (CHAR)
                    if (r["CFACTURA"] != System.DBNull.Value)
                    {
                        sb.Append("'").Append(r["CFACTURA"].ToString()).Append("',");
                        sb_oblig.Append("('").Append(r["CFACTURA"].ToString()).Append("',");
                    }
                    else //no debería venir nunca el codigo factura vacío, pero añadimos 0 en cualquier caso 
                    {
                        sb.Append("null,");
                        sb_oblig.Append("('0',");
                    }
                    //FFACTURA (DATE)
                    if (r["FFACTURA"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["FFACTURA"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");
                    //FFACTDES (DATE)
                    if (r["FFACTDES"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["FFACTDES"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");
                    //FFACTHAS (DATE)
                    if (r["FFACTHAS"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["FFACTHAS"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");
                    //IFACTURA (DECIMAL)
                    if (r["IFACTURA"] != System.DBNull.Value)
                        sb.Append(r["IFACTURA"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IVA (DECIMAL)
                    if (r["IVA"] != System.DBNull.Value)
                        sb.Append(r["IVA"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IIMPUES2 (DECIMAL)
                    if (r["IIMPUES2"] != System.DBNull.Value)
                        sb.Append(r["IIMPUES2"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IIMPUES3 (DECIMAL)
                    if (r["IIMPUES3"] != System.DBNull.Value)
                        sb.Append(r["IIMPUES3"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IBASEISE (DECIMAL)
                    if (r["IBASEISE"] != System.DBNull.Value)
                        sb.Append(r["IBASEISE"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //ISE (DECIMAL)
                    if (r["ISE"] != System.DBNull.Value)
                        sb.Append(r["ISE"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //TFACTURA (TINYINT)
                    if (r["TFACTURA"] != System.DBNull.Value)
                        sb.Append("'").Append(r["TFACTURA"].ToString()).Append("',");
                    else
                        sb.Append("'0',");
                    //TESTFACT (CHAR)
                    if (r["TESTFACT"] != System.DBNull.Value)
                        sb.Append("'").Append(r["TESTFACT"].ToString()).Append("',");
                    else
                        sb.Append("'N',");
                    //DAPERSOC (CHAR)
                    if (r["DAPERSOC"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["DAPERSOC"].ToString())).Append("',");
                        //sb.Append("'").Append(r["DAPERSOC"].ToString().Replace("'"," ")).Append("',");
                        //sb.Append("'").Append(r["DAPERSOC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CNIFDNIC (CHAR)
                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CNIFDNIC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //INDEMPRE (CHAR)
                    if (r["INDEMPRE"] != System.DBNull.Value)
                        sb.Append("'").Append(r["INDEMPRE"].ToString()).Append("',");
                    else
                        sb.Append("'OPERACIONES B2B',");
                    //CUPSREE (CHAR)
                    if (r["CUPSREE"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CUPSREE"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CFACTREC (CHAR)
                    if (r["CFACTREC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CFACTREC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CLINNEG (TINYINT)
                    if (r["CLINNEG"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CLINNEG"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CSEGMERC (CHAR)
                    if (r["CSEGMERC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CSEGMERC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //NUMLABOR (TINYINT)
                    if (r["NUMLABOR"] != System.DBNull.Value)
                        sb.Append("'").Append(r["NUMLABOR"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //TIPONEGOCIO (ENUM 'L','G')
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        sb.Append("'").Append(r["TIPONEGOCIO"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //COMENTARIOS (CHAR)
                    if (r["COMENTARIOS"] != System.DBNull.Value)
                        sb.Append("'").Append(r["COMENTARIOS"].ToString()).Append("'),");
                    else
                        sb.Append("null),");
                    //FLIMPAGO (DATE)
                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        sb_oblig.Append("'").Append(Convert.ToDateTime(r["FLIMPAGO"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb_oblig.Append("null,");
                    //FPTACOBR (DATE)
                    if (r["FPTACOBR"] != System.DBNull.Value)
                        sb_oblig.Append("'").Append(Convert.ToDateTime(r["FPTACOBR"]).ToString("yyyy-MM-dd")).Append("'),");
                    else
                        sb_oblig.Append("null),");
                    #endregion

                    if (j == 50)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();

                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb_oblig.ToString().Substring(0, sb_oblig.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb_oblig = null;
                        sb_oblig = new StringBuilder();

                        j = 0;
                    }

                }
                db.CloseConnection();

                if (j > 0)
                {
                    Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                    firstOnly = true;
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();

                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    commandmy = new MySqlCommand(sb_oblig.ToString().Substring(0, sb_oblig.Length - 1), dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                    sb_oblig = null;
                    sb_oblig = new StringBuilder();

                    j = 0;
                }
       
            }
            catch (Exception ex)
            {
                ficheroLog.addError("AnexaFacturasSAP: " + ex.Message);
                Console.WriteLine("AnexaFacturasSAP: " + ex.Message);
            }
        }
        
        public void AnexaFacturasSIGAME(string factoring, int bloque, DateTime f_factura_des, DateTime f_factura_has,
           DateTime ffactdes, DateTime ffacthas)
        {
            StringBuilder sb = new StringBuilder();
            

            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";
            string strTMP = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            try
            {

               
                //Obtenemos las facturas individuales de SIGAME
                strSql = QueryFactIndSIGAME(f_factura_des, f_factura_has, ffactdes, ffacthas);
                      

                ficheroLog.Add(strSql);
                Console.WriteLine(strSql);
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);
                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    j++;
                    k++;

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO ff_facturas_all_tmp");
                        sb.Append(" (factoring, bloque, CCOUNIPS, CEMPTITU, CREFEREN, SECFACTU, CFACTURA, FFACTURA, FFACTDES, FFACTHAS, IFACTURA, IVA,");
                        sb.Append(" IIMPUES2, IIMPUES3, IBASEISE, ISE, TFACTURA, TESTFACT, DAPERSOC, CNIFDNIC, INDEMPRE, CUPSREE, CFACTREC,");
                        sb.Append(" CLINNEG, CSEGMERC, NUMLABOR, TIPONEGOCIO, COMENTARIOS) values ");

                        

                        firstOnly = false;
                    }

                    #region Campos
                    //factoring (INT)
                    sb.Append("('").Append(factoring).Append("',");
                    //bloque (TINYINT)
                    sb.Append("'").Append(bloque).Append("',");
                    //CCOUNIPS (CHAR)
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CCOUNIPS"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CEMPTITU (TINYINT)
                    if (r["CEMPTITU"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CEMPTITU"].ToString()).Append("',");
                    else
                        sb.Append("'0',");
                    //CREFEREN (BIGINT)
                    if (r["CREFEREN"] != System.DBNull.Value)
                    {
                        //sb.Append("'").Append(r["CREFEREN"].ToString()).Append("',");
                        strTMP = (r["CREFEREN"].ToString());
                        sb.Append("'").Append(strTMP.Where(c => char.IsDigit(c)).ToArray()).Append("',");
                    }
                    else
                        sb.Append("'0',");
                    //SECFACTU (SMALLINT)
                    if (r["SECFACTU"] != System.DBNull.Value)
                        sb.Append("'").Append(r["SECFACTU"].ToString()).Append("',");
                    else
                        sb.Append("'0',");
                    //CFACTURA (CHAR)
                    if (r["CFACTURA"] != System.DBNull.Value)
                    {
                        sb.Append("'").Append(r["CFACTURA"].ToString()).Append("',");
                        
                    }
                    else //no debería venir nunca el codigo factura vacío, pero añadimos 0 en cualquier caso 
                    {
                        sb.Append("null,");
                        
                    }
                    //FFACTURA (DATE)
                    if (r["FFACTURA"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["FFACTURA"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");
                    //FFACTDES (DATE)
                    if (r["FFACTDES"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["FFACTDES"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");
                    //FFACTHAS (DATE)
                    if (r["FFACTHAS"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["FFACTHAS"]).ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");
                    //IFACTURA (DECIMAL)
                    if (r["IFACTURA"] != System.DBNull.Value)
                        sb.Append(r["IFACTURA"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IVA (DECIMAL)
                    if (r["IVA"] != System.DBNull.Value)
                        sb.Append(r["IVA"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IIMPUES2 (DECIMAL)
                    if (r["IIMPUES2"] != System.DBNull.Value)
                        sb.Append(r["IIMPUES2"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IIMPUES3 (DECIMAL)
                    if (r["IIMPUES3"] != System.DBNull.Value)
                        sb.Append(r["IIMPUES3"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //IBASEISE (DECIMAL)
                    if (r["IBASEISE"] != System.DBNull.Value)
                        sb.Append(r["IBASEISE"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //ISE (DECIMAL)
                    if (r["ISE"] != System.DBNull.Value)
                        sb.Append(r["ISE"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //TFACTURA (TINYINT)
                    if (r["TFACTURA"] != System.DBNull.Value)
                        sb.Append("'").Append(r["TFACTURA"].ToString()).Append("',");
                    else
                        sb.Append("'0',");
                    //TESTFACT (CHAR)
                    if (r["TESTFACT"] != System.DBNull.Value)
                        sb.Append("'").Append(r["TESTFACT"].ToString()).Append("',");
                    else
                        sb.Append("'N',");
                    //DAPERSOC (CHAR)
                    if (r["DAPERSOC"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["DAPERSOC"].ToString())).Append("',");
                    //sb.Append("'").Append(r["DAPERSOC"].ToString().Replace("'"," ")).Append("',");
                    //sb.Append("'").Append(r["DAPERSOC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CNIFDNIC (CHAR)
                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CNIFDNIC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //INDEMPRE (CHAR)
                    if (r["INDEMPRE"] != System.DBNull.Value)
                        sb.Append("'").Append(r["INDEMPRE"].ToString()).Append("',");
                    else
                        sb.Append("'OPERACIONES B2B',");
                    //CUPSREE (CHAR)
                    if (r["CUPSREE"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CUPSREE"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CFACTREC (CHAR)
                    if (r["CFACTREC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CFACTREC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CLINNEG (TINYINT)
                    if (r["CLINNEG"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CLINNEG"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //CSEGMERC (CHAR)
                    if (r["CSEGMERC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CSEGMERC"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //NUMLABOR (TINYINT)
                    if (r["NUMLABOR"] != System.DBNull.Value)
                        sb.Append("'").Append(r["NUMLABOR"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //TIPONEGOCIO (ENUM 'L','G')
                    if (r["TIPONEGOCIO"] != System.DBNull.Value)
                        sb.Append("'").Append(r["TIPONEGOCIO"].ToString()).Append("',");
                    else
                        sb.Append("null,");
                    //COMENTARIOS (CHAR)
                    if (r["COMENTARIOS"] != System.DBNull.Value)
                        sb.Append("'").Append(r["COMENTARIOS"].ToString()).Append("'),");
                    else
                        sb.Append("null),");
                    //FLIMPAGO (DATE)
                        
                    //FPTACOBR (DATE)
 
                    #endregion

                    if (j == 50)
                    {
                        Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                        firstOnly = true;
                        dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                        commandmy.ExecuteNonQuery();
                        dbmy.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();

                        

                        j = 0;
                    }

                }
                db.CloseConnection();

                if (j > 0)
                {
                    Console.WriteLine("Anexamos " + String.Format("{0:N0}", k) + " / " + String.Format("{0:N0}", totalRegistros));
                    firstOnly = true;
                    dbmy = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    commandmy = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), dbmy.con);
                    commandmy.ExecuteNonQuery();
                    dbmy.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();

                    

                    j = 0;
                }

            }
            catch (Exception ex)
            {
                ficheroLog.addError("AnexaFacturasSAP: " + ex.Message);
                Console.WriteLine("AnexaFacturasSAP: " + ex.Message);
            }
        }

        // Función de devuelve la consulta SQL de redshift para obtener las facturas individuales de los clientes de SAP (ed_owner.t_ed_h_sap_facts)
        private string QueryFactIndSAP(DateTime f_factura_des, DateTime f_factura_has,
            DateTime ffactdes, DateTime ffacthas)
        {
            string strSql = "";

            strSql = "select 'SAP' as CCOUNIPS, case dc.cd_empr when 'ES21' then '20' when 'ES22' then '70' when 'PT1Q' then '80' else '0' end as CEMPTITU,"
                + " f.cd_di as CREFEREN, 0 as SECFACTU, f.id_fact as CFACTURA, f.fh_fact as FFACTURA,"
                + " f.fh_ini_fact as FFACTDES, f.fh_fin_fact as FFACTHAS, f.im_factdo_con_iva as IFACTURA, case when f.nm_total_iva is null then (f.im_factdo_con_iva-f.im_factdo_sin_iva) else f.nm_total_iva end as IVA, f.im_impuesto_2 as IIMPUES2, f.im_impuesto_3 as IIMPUES3,"
                + " (f.im_factdo_sin_iva - f.im_impto_elect) as IBASEISE, f.im_impto_elect as ISE,"
                + " case f.cd_tp_fact when 'AG' then '6' when 'SD' then '5' when 'CO' then '1'	when 'MC' then '2' else '1'end as TFACTURA,"
                //+ " case f.cd_est_fact when 'A' then 'Y' when 'X' then 'A' else f.cd_est_fact end as TESTFACT,"
                + " f.cd_est_fact as TESTFACT,"
                + " c.tx_apell_cli as DAPERSOC, c.cd_nif_cif_cli as CNIFDNIC, 'OPERACIONES' as INDEMPRE, dc.cd_cups_ext as CUPSREE,"
                + " null as CFACTREC, case dc.cd_linea_negocio when '001' then '1' when '002' then '2' else '1' end as CLINNEG, f.cd_tp_cli as CSEGMERC,"
                + " 0 as NUMLABOR, case dc.cd_linea_negocio when '001' then 'L' when '002' then 'G' else 'L' end as TIPONEGOCIO, null as COMENTARIOS,  o.fh_ini_vencimiento as FLIMPAGO, o.fh_puesta_cobro as FPTACOBR "
                + " from ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc dc on f.cd_di=dc.cd_di"
                //[2025-05-13 GUS]: Vamos a recortar a cups20
                //+ " inner join ed_owner.t_ed_h_ps ps on ps.cd_cups_ext = dc.cd_cups_ext"
                + " inner join ed_owner.t_ed_h_ps ps on SUBSTRING(ps.cd_cups_ext,1,20) = SUBSTRING(dc.cd_cups_ext,1,20)"
                + "	left join ed_owner.t_ed_f_clis c on f.cl_cli = c.cl_cli"
                + "	left join ed_owner.t_ed_h_obligaciones o on f.id_fact = o.cd_fiscal_fact"
                + " where f.im_factdo_con_iva > 0 and f.fh_ini_fact >='" + ffactdes.ToString("yyyy-MM-dd") + "' and f.fh_fin_fact <= '" + ffacthas.ToString("yyyy-MM-dd") + "'"
                + " and f.fh_fact >='" + f_factura_des.ToString("yyyy-MM-dd") + "' and f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " and f.cl_empr = '006' and ( (dc.cd_empr ='PT1Q' and ps.cd_tp_tension in ('MT','AT','MAT')) or (dc.cd_empr='ES21') )"
                + " and f.cd_est_fact not in ('A','X','S')"
                + " and f.id_fact is not null";
            // + " and ps.lg_migrado_sap ='S'";

            return strSql;
        }

        // Función de devuelve la consulta SQL de redshift para obtener las facturas agrupadas de los clientes migrados a SAP
        private string QueryFactAgrSAP(DateTime f_factura_des, DateTime f_factura_has,
           DateTime ffactdes, DateTime ffacthas)
        {
            string strSql = "";

            strSql = "select 'SAP' as CCOUNIPS, case dc.cd_empr when 'ES21' then '20' when 'ES22' then '70' when 'PT1Q' then '80' else '0' end as CEMPTITU,"
                + " f.cd_di as CREFEREN, 0 as SECFACTU, f.id_fact as CFACTURA, f.fh_fact as FFACTURA,"
                + " f.fh_ini_fact as FFACTDES, f.fh_fin_fact as FFACTHAS, f.im_factdo_con_iva as IFACTURA, case when f.nm_total_iva is null then (f.im_factdo_con_iva-f.im_factdo_sin_iva) else f.nm_total_iva end as IVA, f.im_impuesto_2 as IIMPUES2, f.im_impuesto_3 as IIMPUES3,"
                + " (f.im_factdo_sin_iva - f.im_impto_elect) as IBASEISE, f.im_impto_elect as ISE,"
                // + " case f.cd_tp_fact when 'AG' then '6' when 'SD' then '5' when 'CO' then '1'	when 'MC' then '2' else '1'end as TFACTURA,"
                // Añadimos los tipos de agrupadas SD: 'ZCLI','ZCOM','ZDEV','ZGAM','ZGAP','ZINT','ZLIQ','ZACL','ZAGA'
                + " case f.cd_tp_fact when 'AG' then '6' when 'SD' then '5' when 'ZCLI' then '5' when 'ZCOM' then '5' when 'ZDEV' then '5' when 'ZGAM' then '5' when 'ZGAP' then '1' when 'ZINT' then '5' when 'ZLIQ' then '5' when 'ZACL' then '5' when 'ZAGA' then '5'"
                + " when 'CO' then '1' when 'MC' then '2' else '1' end as TFACTURA,"
                //+ " case f.cd_est_fact when 'A' then 'Y' when 'X' then 'A' else f.cd_est_fact end as TESTFACT,"
                + " f.cd_est_fact as TESTFACT,"
                + " c.tx_apell_cli as DAPERSOC, c.cd_nif_cif_cli as CNIFDNIC, 'OPERACIONES' as INDEMPRE, null as CUPSREE,"
                + " null as CFACTREC, case dc.cd_linea_negocio when '001' then '1' when '002' then '2' else '1' end as CLINNEG, f.cd_tp_cli as CSEGMERC,"
                + " 0 as NUMLABOR, case dc.cd_linea_negocio when '001' then 'L' when '002' then 'G' else 'L' end as TIPONEGOCIO, null as COMENTARIOS, o.fh_ini_vencimiento as FLIMPAGO, o.fh_puesta_cobro as FPTACOBR "
                + " from ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc dc on f.cd_di=dc.cd_di"
                //[2025-05-13 GUS]: Vamos a recortar a cups20
                // + " inner join ed_owner.t_ed_h_ps ps on ps.cd_cups_ext = dc.cd_cups_ext"
                + " inner join ed_owner.t_ed_h_ps ps on SUBSTRING(ps.cd_cups_ext,1,20) = SUBSTRING(dc.cd_cups_ext,1,20)"
                + "	left join ed_owner.t_ed_f_clis c on f.cl_cli = c.cl_cli"
                + "	left join ed_owner.t_ed_h_obligaciones o on f.id_fact = o.cd_fiscal_fact"
                + " where f.im_factdo_con_iva > 0 and f.fh_ini_fact is null and f.fh_fin_fact is null"
                + " and f.fh_fact >='" + f_factura_des.ToString("yyyy-MM-dd") + "' and f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " and f.cl_empr = '006' and ( (dc.cd_empr ='PT1Q' and ps.cd_tp_tension in ('MT','AT','MAT')) or (dc.cd_empr='ES21') )"
                + " and f.cd_est_fact not in ('A','X','S')"
                + " and f.id_fact is not null"
                //+ " and ps.lg_migrado_sap ='S'"
                + " group by ccounips,cemptitu,creferen,secfactu,cfactura,ffactura,ffactdes,ffacthas,ifactura,iva,iimpues2,iimpues3,ibaseise,ise,tfactura,testfact,dapersoc,cnifdnic,indempre,cupsree,cfactrec,clinneg,csegmerc,numlabor,tiponegocio,flimpago,fptacobr";

            return strSql;
        }

        // Función de devuelve la consulta SQL de redshift para obtener las facturas agrupadas de los clientes migrados a SAP, sin cruzar con tabla PS, sin CUPS y solo España
        private string QueryFactAgrSAP_no_PS(DateTime f_factura_des, DateTime f_factura_has,
           DateTime ffactdes, DateTime ffacthas)
        {
            string strSql = "";

            strSql = "select 'SAP_no_PS' as CCOUNIPS, case dc.cd_empr when 'ES21' then '20' when 'ES22' then '70' when 'PT1Q' then '80' else '0' end as CEMPTITU,"
                + " f.cd_di as CREFEREN, 0 as SECFACTU, f.id_fact as CFACTURA, f.fh_fact as FFACTURA,"
                + " f.fh_ini_fact as FFACTDES, f.fh_fin_fact as FFACTHAS, f.im_factdo_con_iva as IFACTURA, case when f.nm_total_iva is null then (f.im_factdo_con_iva-f.im_factdo_sin_iva) else f.nm_total_iva end as IVA, f.im_impuesto_2 as IIMPUES2, f.im_impuesto_3 as IIMPUES3,"
                + " (f.im_factdo_sin_iva - f.im_impto_elect) as IBASEISE, f.im_impto_elect as ISE,"
                // + " case f.cd_tp_fact when 'AG' then '6' when 'SD' then '5' when 'CO' then '1'	when 'MC' then '2' else '1'end as TFACTURA,"
                // Añadimos los tipos de agrupadas SD: 'ZCLI','ZCOM','ZDEV','ZGAM','ZGAP','ZINT','ZLIQ','ZACL','ZAGA'
                + " case f.cd_tp_fact when 'AG' then '6' when 'SD' then '5' when 'ZCLI' then '5' when 'ZCOM' then '5' when 'ZDEV' then '5' when 'ZGAM' then '5' when 'ZGAP' then '1' when 'ZINT' then '5' when 'ZLIQ' then '5' when 'ZACL' then '5' when 'ZAGA' then '5'"
                + " when 'CO' then '1' when 'MC' then '2' else '1' end as TFACTURA,"
                 //+ " case f.cd_est_fact when 'A' then 'Y' when 'X' then 'A' else f.cd_est_fact end as TESTFACT,"
                + " f.cd_est_fact as TESTFACT,"
                + " c.tx_apell_cli as DAPERSOC, c.cd_nif_cif_cli as CNIFDNIC, 'OPERACIONES' as INDEMPRE, null as CUPSREE,"
                + " null as CFACTREC, case dc.cd_linea_negocio when '001' then '1' when '002' then '2' else '1' end as CLINNEG, f.cd_tp_cli as CSEGMERC,"
                + " 0 as NUMLABOR, case dc.cd_linea_negocio when '001' then 'L' when '002' then 'G' else 'L' end as TIPONEGOCIO, null as COMENTARIOS, o.fh_ini_vencimiento as FLIMPAGO, o.fh_puesta_cobro as FPTACOBR "
                + " from ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc dc on f.cd_di=dc.cd_di"
                //+ " inner join ed_owner.t_ed_h_ps ps on ps.cd_cups_ext = dc.cd_cups_ext"
                + "	left join ed_owner.t_ed_f_clis c on f.cl_cli = c.cl_cli"
                + "	left join ed_owner.t_ed_h_obligaciones o on f.id_fact = o.cd_fiscal_fact"
                + " where f.im_factdo_con_iva > 0 and f.fh_ini_fact is null and f.fh_fin_fact is null"
                + " and f.fh_fact >='" + f_factura_des.ToString("yyyy-MM-dd") + "' and f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " and f.cl_empr = '006' and dc.cd_empr='ES21'"
                + " and f.cd_est_fact not in ('A','X','S')"
                + " and f.id_fact is not null"
                //+ " and ps.lg_migrado_sap ='S'"
                + " group by ccounips,cemptitu,creferen,secfactu,cfactura,ffactura,ffactdes,ffacthas,ifactura,iva,iimpues2,iimpues3,ibaseise,ise,tfactura,testfact,dapersoc,cnifdnic,indempre,cupsree,cfactrec,clinneg,csegmerc,numlabor,tiponegocio,flimpago,fptacobr";

            return strSql;
        }
        // Función de devuelve la consulta SQL de redshift para obtener las facturas SD agrupadas de SAP, obtenidas de la tabla ed_owner.sap_tfactura_n0
        // Según nos indican son las del tipo tp_fact in ('ZCLI','ZCOM','ZDEV','ZGAM','ZINT','ZLIQ','ZACL','ZAGA')
        private string QueryFactAgrSAP_N0_SD(DateTime f_factura_des, DateTime f_factura_has,
           DateTime ffactdes, DateTime ffacthas)
        {
            string strSql = "";

            strSql = "select 'SAP_N0_SD' as CCOUNIPS, case f.cd_sociedad when 'ES21' then '20' when 'ES22' then '70' when 'PT1Q' then '80' else '0' end as CEMPTITU,"
                + " f.cd_di as CREFEREN, 0 as SECFACTU, f.nm_doc_oficial as CFACTURA, f.fh_contab as FFACTURA,"
                + " f.fh_desde_fact as FFACTDES, f.fh_hasta_fact as FFACTHAS, f.im_factura as IFACTURA, f.im_impuesto1 as IVA, f.im_impuesto2 as IIMPUES2, f.im_impuesto3 as IIMPUES3,"
                + " f.nm_base_ise as IBASEISE, f.im_base_ise as ISE,"
                + " '5' as TFACTURA,"
                //+ " case f.cd_estado_fact when 'A' then 'Y' when 'X' then 'A' else f.cd_estado_fact end as TESTFACT,"
                + " f.cd_estado_fact as TESTFACT,"
                + " case when c.tx_apell_cli is null then f.de_nom_apell else c.tx_apell_cli end as DAPERSOC, f.nm_id as CNIFDNIC, f.cd_ind_cef_ope as INDEMPRE, null as CUPSREE,"
                + " null as CFACTREC, '001' as CLINNEG, f1.cd_tp_cli as CSEGMERC,"
                + " 0 as NUMLABOR, 'L' as TIPONEGOCIO, null as COMENTARIOS, o.fh_ini_vencimiento as FLIMPAGO, o.fh_puesta_cobro as FPTACOBR "

                + " from ed_owner.sap_tfactura_n0 f"
                + " inner join ed_owner.t_ed_h_sap_facts f1 on f.cd_di=f1.cd_di"
                // + " inner join ed_owner.t_ed_h_ps ps on ps.cd_cups_ext = dc.cd_cups_ext"
                + "	left join ed_owner.t_ed_f_clis c on f1.cl_cli = c.cl_cli"
                + "	left join ed_owner.t_ed_h_obligaciones o on f.nm_doc_oficial = o.cd_fiscal_fact"
                + " where f.im_factura > 0 and f.fh_desde_fact is null and f.fh_hasta_fact is null"
                + " and f.fh_contab >='" + f_factura_des.ToString("yyyy-MM-dd") + "' and f.fh_contab <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " and f.tp_fact in ('ZCLI','ZCOM','ZDEV','ZGAM','ZINT','ZLIQ','ZACL','ZAGA')" //[07/03/2025] GUS: Eliminamos del filtro las ZGAP que corresponden a GAS
                + " and f.cd_estado_fact not in ('A','X','S')"
                + " and f.nm_doc_oficial is not null"
                + " and f.cd_sector <> '2'" //[07/03/2025] GUS: Descartamos las de GAS
                + " group by ccounips,cemptitu,creferen,secfactu,cfactura,ffactura,ffactdes,ffacthas,ifactura,iva,iimpues2,iimpues3,ibaseise,ise,tfactura,testfact,dapersoc,cnifdnic,indempre,cupsree,cfactrec,clinneg,csegmerc,numlabor,tiponegocio,flimpago,fptacobr";

            return strSql;
        }
        // Función de devuelve la consulta SQL de SQLServer (SIGAME) para obtener las facturas individuales de los clientes de GAS
        private string QueryFactIndSIGAME(DateTime f_factura_des, DateTime f_factura_has,
            DateTime ffactdes, DateTime ffacthas)
        {
            string strSql = "";

            strSql = "select 'SIGAME' as CCOUNIPS, case T_SGM_G_PS.CD_PAIS when 'ESP' then '20' when 'POR' then '80' else '0' end as CEMPTITU,"
                + " CONCAT(T_SGM_M_FACTURAS_REALES_PS.ID_FACTURA,T_SGM_M_FACTURAS_REALES_PS.CD_NFACTURA_REALES_PS) as CREFEREN, '0' as SECFACTU, T_SGM_M_FACTURAS_REALES_PS.CD_NFACTURA_REALES_PS as CFACTURA, T_SGM_M_FACTURAS_REALES_PS.FH_FACTURA as FFACTURA,"
                + " T_SGM_M_FACTURAS_REALES_PS.FH_INI_FACTURACION as FFACTDES, T_SGM_M_FACTURAS_REALES_PS.FH_FIN_FACTURACION as FFACTHAS, T_SGM_M_FACTURAS_REALES_PS.NM_IMPORTE_BRUTO as IFACTURA, T_SGM_M_FACTURAS_REALES_PS.NM_IEH_IMPORTE + T_SGM_M_FACTURAS_REALES_PS.NM_IMPORTE_IVA as IVA, '0' as IIMPUES2, '0' as IIMPUES3,"
                + " '0' as IBASEISE, '0' as ISE,"
                + " '1' as TFACTURA,"
                + " T_SGM_M_FACTURAS_REALES_PS.CD_TESTFACT as TESTFACT,"
                + " T_SGM_M_CLIENTES.DE_NOMBRE_CLIENTE as DAPERSOC, T_SGM_M_CLIENTES.CD_CIF as CNIFDNIC, 'OPERACIONES' as INDEMPRE, T_SGM_G_PS.CD_CUPS as CUPSREE,"
                + " null as CFACTREC, '2' as CLINNEG, 'N' as CSEGMERC,"
                + " 0 as NUMLABOR, 'G' as TIPONEGOCIO, null as COMENTARIOS, null as FLIMPAGO, null as FPTACOBR "
                + " from T_SGM_G_PS"
                + " inner join T_SGM_G_CONTRATOS_PS"
                + " on T_SGM_G_PS.ID_PS = T_SGM_G_CONTRATOS_PS.ID_PS"
                + " inner"
                + " join T_SGM_M_CLIENTES"
                + " on T_SGM_G_PS.ID_CLIENTE = T_SGM_M_CLIENTES.ID_CLIENTE"
                + " inner"
                + " join dbo.T_SGM_M_FACTURAS_REALES_PS"
                + " on T_SGM_G_CONTRATOS_PS.ID_CTO_PS = T_SGM_M_FACTURAS_REALES_PS.ID_CTO_PS"
                + " WHERE T_SGM_M_FACTURAS_REALES_PS.NM_IMPORTE_BRUTO > 0 "
                + " AND T_SGM_M_FACTURAS_REALES_PS.FH_INI_FACTURACION >= '" + ffactdes.ToString("yyyy-MM-dd") + "'"
                + " AND T_SGM_M_FACTURAS_REALES_PS.FH_FIN_FACTURACION <= '" + ffacthas.ToString("yyyy-MM-dd") + "'"
                + " AND T_SGM_M_FACTURAS_REALES_PS.FH_FACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                + " AND T_SGM_M_FACTURAS_REALES_PS.FH_FACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " AND T_SGM_M_FACTURAS_REALES_PS.CD_NFACTURA_REALES_PS IS NOT null";

            return strSql;
        }
    }
}
