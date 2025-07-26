using EndesaBusiness.servidores;
using EndesaEntity.xml;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.factoring
{
    public class SeguimientoMes13
    {

        private EndesaBusiness.utilidades.Param param;

        public SeguimientoMes13()
        {

        }

        public SeguimientoMes13(int seguimiento, DateTime ffactura_des, DateTime ffactura_has, DateTime ffactdes, DateTime ffacthas)
        {

            param = new utilidades.Param("13_param", servidores.MySQLDB.Esquemas.FAC);

            AnexionDatosIndividuales(ffactura_des, ffactura_has, ffactdes, ffacthas);
            MarcaCisternas();
            BorradoFacturasIndividuales("A");
            BorradoFacturasIndividuales("S");
            UltimaFacturaIndividual();
            GuardaImporteHidrocarburo();
            AnexaSeguimientoIndividuales(seguimiento, ffactdes, ffacthas);
            InformaFechaLimitePagoIndividuales(seguimiento);
        }

        public void SeguimientoMes13Agrupadas(int seguimiento, DateTime ffactura_des, DateTime ffactura_has)
        {
            AnexionDatosAgrupadas(ffactura_des, ffactura_has);
            BorradoFacturasAgrupadas("A");
            BorradoFacturasAgrupadas("S");
            // [05/02/2025] GUS: Añadimos nueva función: eliminar facturas que tengan en su código fiscal los caracteres 'FC' en posición 4 y 5 respectivamente
            BorrarFacturasAgrupadasPorCodigoFiscal();

            UltimaFacturaAgrupada();
            AnexaSeguimientoAgrupadas(seguimiento);
            InformaFechaLimitePagoAgrupadas(seguimiento);
        }


        private void InformaFechaLimitePagoIndividuales(int seguimiento)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "UPDATE 13_seg_individuales_detalle f"
                + " left outer join 13_obligaciones_sap s on s.id_fact = f.cfactura"
                + " left outer join deuda_obligaciones_original d on"
                + " d.CEMPTITU = f.CEMPTITU and"
                + " d.CREFEREN = f.CREFEREN and"
                + " d.SECFACTU = f.SECFACTU"
                + " left outer join 13_oblig_cob o on"
                + " o.CEMPTITU = f.CEMPTITU and"
                + " o.CREFEREN = f.CREFEREN and"
                + " o.SECFACTU = f.SECFACTU"
                + " SET f.flimpago = if (s.flimpago is null or s.flimpago = '0000-00-00',if (o.FLIMPAGO is null or o.FLIMPAGO = '0000-00-00', d.FLIMPAGO, o.FLIMPAGO) , s.flimpago),"
                + " f.fptacobr = if(s.fptacobr is null, o.fptacobr, s.fptacobr),"
                + " f.seguimiento_codigo = CONCAT(" + seguimiento
                + ", if (o.FLIMPAGO is not NULL or d.FLIMPAGO IS not NULL or s.flimpago is not NULL,'1','0'))";


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        private void InformaFechaLimitePagoAgrupadas(int seguimiento)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "UPDATE 13_seg_agrupadas_detalle f"
               + " left outer join 13_obligaciones_sap s on s.id_fact = f.cfactura"
               + " left outer join deuda_obligaciones_original d on"
               + " d.CEMPTITU = f.CEMPTITU and"
               + " d.CREFEREN = f.CREFEREN and"
               + " d.SECFACTU = f.SECFACTU"
               + " left outer join 13_oblig_cob o on"
               + " o.CEMPTITU = f.CEMPTITU and"
               + " o.CREFEREN = f.CREFEREN and"
               + " o.SECFACTU = f.SECFACTU"
               + " SET f.flimpago = if (s.flimpago is null or s.flimpago = '0000-00-00',if (o.FLIMPAGO is null or o.FLIMPAGO = '0000-00-00', d.FLIMPAGO, o.FLIMPAGO) , s.flimpago),"
               + " f.fptacobr = if(s.fptacobr is null, o.fptacobr, s.fptacobr),"
               + " f.seguimiento_codigo = CONCAT(" + seguimiento
                + ", if (o.FLIMPAGO is not NULL or d.FLIMPAGO IS not NULL or s.flimpago is not NULL,'1','0'))";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void AnexionDatosIndividuales(DateTime f_factura_des, DateTime f_factura_has, DateTime ffactdes, DateTime ffacthas)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                // Borramos la tabla de facturas de seguimiento
                Console.WriteLine("Borramos la tabla de facturas de seguimiento");
                strSql = "delete from 13_seg_facturas_individuales";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //30/04/2025 GUS: Primero Anexamos las facturas individuales de GAS de SAP
                Console.WriteLine("Inicio anexión facturas individuales clientes SAP de GAS");
                AnexaFacturasSAP(f_factura_des, f_factura_has, ffactdes, ffacthas, "IND_GAS");
                Console.WriteLine("Fin anexión facturas individuales clientes SAP de GAS ");
                
                // //06/02/2025 GUS: Después anexamos las facturas de gas desde SIGAME, por si en SAP no está volcada alguna
                // Console.WriteLine("Inicio anexión facturas individuales clientes SIGAME ");
                // AnexaFacturasSIGAME(f_factura_des, f_factura_has, ffactdes, ffacthas);
                // Console.WriteLine("Fin anexión facturas individuales clientes SIGAME ");

              
                //Individuales de LUZ
                Console.WriteLine("Inicio anexión facturas individuales clientes SAP ");
                AnexaFacturasSAP(f_factura_des, f_factura_has, ffactdes, ffacthas, "IND");
                Console.WriteLine("Fin anexión facturas individuales clientes SAP ");

                


                Console.WriteLine("Facturas emitidas desde el "
                    + f_factura_des.ToString("dd/MM/yyyy") + " hasta el "
                    + f_factura_has.ToString("dd/MM/yyyy"));

                strSql = "replace into fact.13_seg_facturas_individuales"
                    + " select f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.TESTFACT, f.TFACTURA,"
                    + " f.CNIFDNIC, f.CCOUNIPS, substr(f.CUPSREE,1,20) as CUPSREE, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.IFACTURA,"
                    + " f.IVA, f.IIMPUES2, f.IIMPUES3, f.ISE, null as hidrocarburos, f.CFACTREC, f.COMENTARIO_REFACT, f.TIPONEGOCIO, f.INDEMPRE"
                    + " from fact.13_seg_individuales s inner join"
                    + " fact.fo f on"
                    + " f.cnifdnic = s.nif"
                    + " left join 13_seg_facturas_individuales t on t.cfactura = f.CFACTURA"
                    + " where "
                    + " (f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                    + " and f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "')"
                    + " and (f.FFACTDES >= '" + ffactdes.ToString("yyyy-MM-dd") + "'"
                    + " and f.FFACTHAS <= '" + ffacthas.ToString("yyyy-MM-dd") + "')"
                    + " and f.TFACTURA in (1,2) and f.IFACTURA >= 0 and t.cfactura is null;";
                // + " and f.TFACTURA in (1,2) and f.IFACTURA >= 0 and f.CFACTURA NOT IN (select t.cfactura from 13_seg_facturas_individuales t);";


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

               

                // Adjuntamos facturas de GAS
                Console.WriteLine("Adjuntamos facturas de GAS");
                strSql = "replace into fact.13_seg_facturas_individuales"
                    + " select f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.TESTFACT, f.TFACTURA,"
                    + " f.CNIFDNIC, f.CCOUNIPS, f.CUPSREE, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.IFACTURA,"
                    + " f.IVA, f.IIMPUES2, f.IIMPUES3, f.ISE, null as hidrocarburos, f.CFACTREC, f.COMENTARIO_REFACT, f.TIPONEGOCIO, f.INDEMPRE"
                    + " from fo f"
                    + " left join 13_seg_facturas_individuales t on t.cfactura = f.CFACTURA"
                    + " where"
                    + " f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "' and"
                    + " f.TIPONEGOCIO = 'G' and f.IFACTURA >= 0  and t.cfactura is null;";
                //  + " f.TIPONEGOCIO = 'G' and f.IFACTURA >= 0  and f.CFACTURA NOT IN (select t.cfactura from 13_seg_facturas_individuales t);";
                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Cambiamos el tipo de las facturas de GAS por tipo 1
                Console.WriteLine("Cambiamos el tipo de las facturas de GAS por tipo 1");
                strSql = "update fact.13_seg_facturas_individuales set tfactura = 1 where TIPONEGOCIO = 'G';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de consumos del mismo mes de emisión de facturas
                Console.WriteLine("Quitamos las facturas de consumos del mismo mes de emisión de facturas");
                strSql = "delete from fact.13_seg_facturas_individuales where FFACTDES >= '" + f_factura_des.ToString("yyyy-MM-dd") + "';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de EEXXI
                Console.WriteLine("Quitamos las facturas de EEXXI");
                strSql = "delete from fact.13_seg_facturas_individuales where CEMPTITU = 70;";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //TrataHolanda(f_factura_des, f_factura_has, ffactdes);

                // Quitamos las facturas de BTN
                Console.WriteLine("Quitamos las facturas de BTN");
                strSql = "delete from fact.13_seg_facturas_individuales where CEMPTITU = 20 and INDEMPRE = 'CEFACO';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Quitamos las facturas de BTE
                Console.WriteLine("Quitamos las facturas de BTE");
                strSql = "delete from fact.13_seg_facturas_individuales where CEMPTITU = 80 and INDEMPRE = 'CEFACO';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Borramos las facturas de consumos que no correspondan con las fechas que queramos
                Console.WriteLine("Borramos las facturas de consumos que no correspondan con las fechas que queramos");
                strSql = "delete from fact.13_seg_facturas_individuales where ffacthas < '" + ffactdes.ToString("yyyy-MM-dd") + "';";
                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // En este momento y antes de seguir borrando algunas facturas procedemos a guardar las facturas
                // en otra tabla para realizar búsquedas.


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "AnexionDatos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void AnexionDatosAgrupadas(DateTime f_factura_des, DateTime f_factura_has)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                // Borramos la tabla de facturas de seguimiento
                Console.WriteLine("Borramos la tabla de facturas de seguimiento");
                strSql = "delete from 13_seg_facturas_agrupadas;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                //28-10-2024 GUS: ubicamos en primer lugar para que posteriormente prevalezca los datos de t_ed_h_sap_facts
                Console.WriteLine("Inicio anexión facturas agrupadas clientes SAP N0 SD ");
                AnexaFacturasSAP(f_factura_des, f_factura_has, f_factura_des, f_factura_has, "AGR_N0_SD");
                Console.WriteLine("Fin anexión facturas agrupadas clientes SAP N0 SD ");

                //05/06/2024 GUS: Anexamos facturas agrupadas de SAP - CLIENTES MIGRADOS
                Console.WriteLine("Inicio anexión facturas agrupadas clientes SAP ");
                AnexaFacturasSAP(f_factura_des, f_factura_has, f_factura_des, f_factura_has, "AGR");
                Console.WriteLine("Fin anexión facturas agrupadas clientes SAP ");

                //28-10-2024 GUS
                Console.WriteLine("Inicio anexión facturas agrupadas clientes SAP - sin cruce con PS y solo España ");
                AnexaFacturasSAP(f_factura_des, f_factura_has, f_factura_des, f_factura_has, "AGR_no_PS");
                Console.WriteLine("Fin anexión facturas agrupadas clientes SAP - - sin cruce con PS y solo España ");

                //Console.WriteLine("Inicio anexión facturas agrupadas N0 SD ");
                //AnexaFacturasSAP(f_factura_des, f_factura_has, f_factura_des, f_factura_has, "AGR_N0_SD");
                //Console.WriteLine("Fin anexión facturas agrupadas N0 SD ");


                // Facturas emitidas en "Cambiar por el mes que corresponda"
                Console.WriteLine("Facturas emitidas desde el "
                    + f_factura_des.ToString("dd/MM/yyyy") + " hasta el "
                    + f_factura_has.ToString("dd/MM/yyyy"));

                strSql = "replace into fact.13_seg_facturas_agrupadas"
                    + " select f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.TESTFACT, f.TFACTURA,"
                    + " f.CNIFDNIC, f.CCOUNIPS, f.CUPSREE, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.IFACTURA,"
                    + " f.IVA, f.IIMPUES2, f.IIMPUES3, f.ISE, null as hidrocarburos, f.CFACTREC, f.COMENTARIO_REFACT, f.TIPONEGOCIO, f.INDEMPRE"
                    + " from fact.13_seg_agrupadas s inner join"
                    + " fact.fo f on"
                    + " f.CEMPTITU = s.empresa_titular and"
                    + " f.CNIFDNIC = s.nif"
                    + " left join 13_seg_facturas_agrupadas t on t.cfactura = f.CFACTURA"
                    + " where f.CEMPTITU <> 70 and"
                    + " (f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd") + "'"
                    + " and f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd") + "')"
                    + " and (f.TFACTURA = 6 or f.TFACTURA = 5) and f.IFACTURA > 0 and t.cfactura is null;";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "AnexionDatos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void MarcaCisternas()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                strSql = "update fact.13_seg_facturas_individuales f"
                    + " left outer join fo_s s on"
                    + " s.CREFEREN = f.CREFEREN and"
                    + " s.SECFACTU = f.SECFACTU"
                    + " left outer join cm_inventario_gas i on"
                    + " i.ID_PS = s.ID_PS and"
                    + " (i.FINICIO <= f.FFACTDES and(i.FFIN >= f.FFACTHAS or i.FFIN is null)) and"
                    + " i.Grupo = 'Cisterna'"
                    + " set f.CUPSREE = concat('CISTERNA_', i.ID_PS)"
                    + " where"
                    + " f.TIPONEGOCIO = 'G' and"
                    + " (f.CUPSREE is null or trim(f.CUPSREE = ''));";

                Console.WriteLine(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "MarcaCisternas",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        private void BorradoFacturasIndividuales(string testfact)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                Console.WriteLine("Borrando facturas de tipo " + testfact);
                strSql = "delete from 13_seg_facturas_individuales where testfact = '" + testfact + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "BorradoFacturas",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        private void BorradoFacturasAgrupadas(string testfact)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                Console.WriteLine("Borrando facturas de tipo " + testfact);
                strSql = "delete from 13_seg_facturas_agrupadas where testfact = '" + testfact + "'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "BorradoFacturas",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        // Función para eliminar las facturas agrupadas que incluyen el literal FC en las posiciones 4 y 5 del código fiscal (cfactura)
        // Petición de Alvaro Carbo a 04/02/2025
        private void BorrarFacturasAgrupadasPorCodigoFiscal()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                Console.WriteLine("Borrando facturas agrupadas que incluyen en su código fiscal el literal FC en la posición 4 y 5");
                strSql = "delete from 13_seg_facturas_agrupadas where SUBSTRING(cfactura,4,2) = 'FC'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "BorradoFacturasAgrupadasPorCodigoFiscal",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        private void UltimaFacturaIndividual()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                Console.WriteLine("Proceso para quedarnos con la última factura emitida");
                Console.WriteLine("Borrado de la tabla 13_refacturaciones");
                strSql = "delete from 13_refacturaciones;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Buscamos las refecturaciones.");
                strSql = "replace into 13_refacturaciones select f.CFACTREC"
                      + " from 13_seg_facturas_individuales f where (f.CFACTREC is not null and f.CFACTREC <> '')";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Borramos las facturas");
                strSql = "delete f"
                    + " from 13_seg_facturas_individuales f"
                    + " inner join 13_refacturaciones f1 on"
                    + " f.CFACTURA = f1.CFACTREC;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "UltimaFactura",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        private void UltimaFacturaAgrupada()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                Console.WriteLine("Proceso para quedarnos con la última factura emitida");
                Console.WriteLine("Borrado de la tabla 13_refacturaciones");
                strSql = "delete from 13_refacturaciones;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Buscamos las refecturaciones.");
                strSql = "replace into 13_refacturaciones select f.CFACTREC"
                      + " from 13_seg_facturas_agrupadas f where (f.CFACTREC is not null and f.CFACTREC <> '')";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("Borramos las facturas");
                strSql = "delete f"
                    + " from 13_seg_facturas_agrupadas f"
                    + " inner join 13_refacturaciones f1 on"
                    + " f.CFACTURA = f1.CFACTREC;";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "UltimaFactura",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }

        private void GuardaImporteHidrocarburo()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            strSql = "update 13_seg_facturas_individuales set hidrocarburos = (select sum(t.ICONFAC) from fo_tcon t where"
                + " t.CREFEREN = 13_seg_facturas_individuales.CREFEREN and"
                + " t.SECFACTU = 13_seg_facturas_individuales.SECFACTU and"
                + " t.TESTFACT = 13_seg_facturas_individuales.TESTFACT and"
                + " t.TCONFAC in (1100, 1101, 1102, 1103, 1104, 1105, 1106, 1107, 1504, 1505, 1512, 1633, 1634))";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void AnexaSeguimientoAgrupadas(int seguimiento)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "replace into 13_seg_agrupadas_detalle"
                + " select " + seguimiento + " as seguimiento, null, f.cnifdnic, f.cemptitu,  f.creferen, f.secfactu, f.cfactura,"
                + " f.ffactura, f.ffactdes, f.ffacthas,"
                + " if (os.flimpago is null or os.flimpago = '0000-00-00',if (o.FLIMPAGO is null or o.FLIMPAGO = '0000-00-00', d.FLIMPAGO, o.FLIMPAGO) , os.flimpago) as flimpago,"
                + " if(os.fptacobr is null, o.fptacobr, os.fptacobr),"
                + " f.ifactura, (if (f.iva is null,0,f.iva) + if (f.iimpues2 is null, 0, f.iimpues2) + if (f.iimpues3 is null, 0, f.iimpues3)"
                + " + if (f.ise is null, 0, f.ise) +"
                + " if (f.hidrocarburos is null, 0, f.hidrocarburos)) as impuestos, null as comentarios"
                + " from 13_seg_agrupadas s inner join"
                + " 13_seg_facturas_agrupadas f on"
                + " f.cemptitu = s.empresa_titular and"
                + " f.cnifdnic = s.nif"
                + " left outer join 13_obligaciones_sap os on os.id_fact = f.cfactura"
                + " left outer join deuda_obligaciones_original d on"
                + " d.CEMPTITU = f.CEMPTITU and"
                + " d.CREFEREN = f.CREFEREN and"
                + " d.SECFACTU = f.SECFACTU"
                + " left outer join 13_oblig_cob o on"
                + " o.CEMPTITU = f.CEMPTITU and"
                + " o.CREFEREN = f.CREFEREN and"
                + " o.SECFACTU = f.SECFACTU";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }
        private void AnexaSeguimientoIndividuales(int seguimiento, DateTime ffactdes, DateTime ffacthas)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            // Cargamos las facturas individuales ya publicadas
            Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>> dic_publicadas
                = CargaSeguimientoFacturasIndividuales();


            strSql = "select cupsree, ffactdes, ffacthas from 13_seg_facturas_individuales";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["cupsree"] != System.DBNull.Value)
                {
                    List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> o;
                    if (dic_publicadas.TryGetValue(r["cupsree"].ToString().ToUpper(), out o))
                    {
                        for (int i = 0; i < o.Count; i++)
                        {
                            // La factura publicada está completa y no necesitamos volver a publicarla
                            if (o[i].ffactdes == ffactdes && o[i].ffacthas == ffacthas)
                                BorrarFacturaPorCUPSREE(o[i].cupsree);
                            // Alguna partición de la factura se ha vuelto a refacturar
                            else if (o[i].ffactdes == Convert.ToDateTime(r["ffactdes"]) || o[i].ffacthas == Convert.ToDateTime(r["ffacthas"]))
                                BorrarFacturaPorCUPSREE(o[i].cupsree);
                        }
                    }
                }


            }
            db.CloseConnection();

            // Finalmente anexamos las facturas que quedan

            strSql = "replace into 13_seg_individuales_detalle"
                + " select " + seguimiento + " as seguimiento, null as seguimiento_codigo,"
                + " f.cupsree, f.cemptitu,  f.creferen, f.secfactu, f.cfactura,"
                + " f.ffactura, f.ffactdes, f.ffacthas,"
                + " null as flimpago, null as fptacobr,"
                + " f.ifactura, (if (f.iva is null,0,f.iva) + if (f.iimpues2 is null, 0, f.iimpues2) + if (f.iimpues3 is null, 0, f.iimpues3)"
                + " + if (f.ise is null, 0, f.ise) +"
                + " if (f.hidrocarburos is null, 0, f.hidrocarburos)) as impuestos, null as comentarios, f.cnifdnic"
                + " from 13_seg_individuales s inner join"
                + " 13_seg_facturas_individuales f on"
                + " f.cemptitu = s.empresa_titular and"
                + " f.cnifdnic = s.nif and"
                + " f.cupsree = s.cupsree";

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();




        }

        private void BorrarFacturaPorCUPSREE(string cupsree)
        {
            MySQLDB db;
            MySqlCommand command;

            string strSql = "";

            strSql = "delete from 13_seg_facturas_individuales where cupsree = '" + cupsree + "';";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        private void TrataHolanda(DateTime f_factura_des, DateTime f_factura_has, DateTime ffactdes)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            try
            {
                Console.WriteLine("Tratando facturas de Holanda.");

                strSql = "replace into fact.13_seg_facturas_individuales"
                    + " select f.CEMPTITU, f.CREFEREN, f.SECFACTU, f.TESTFACT, f.TFACTURA,"
                    + " f.CNIFDNIC, f.CCOUNIPS, f.CUPSREE, f.CFACTURA, f.FFACTURA, f.FFACTDES, f.FFACTHAS, f.IFACTURA,"
                    + " f.IVA, f.IIMPUES2, f.IIMPUES3, f.ISE, null as hidrocarburos, f.CFACTREC, f.COMENTARIO_REFACT, f.TIPONEGOCIO, f.INDEMPRE"
                    + " from fo f where"
                    + " f.FFACTURA >= '" + f_factura_des.ToString("yyyy-MM-dd")
                    + "' and f.FFACTURA <= '" + f_factura_has.ToString("yyyy-MM-dd")
                    + "' and f.CNIFDNIC like 'NL%' and"
                    + " f.COMENTARIO_REFACT like '%" + ffactdes.ToString("yyyyMM") + "%'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                Console.WriteLine("     Actualizando CUPSREE a través del comentario.");
                strSql = "update fact.13_seg_facturas_individuales set CUPSREE = concat(substr(COMENTARIO_REFACT, 3, 2), substr(COMENTARIO_REFACT, 8, 16))"
                    + " where CNIFDNIC like 'NL%';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "update fact.13_seg_facturas_individuales set ffactdes = '" + f_factura_des.ToString("yyyy-MM-dd") + "',"
                    + " ffacthas = '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                   + " where CNIFDNIC like 'NL%';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "TrataHolanda",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }
        public void SeguimientoExcel(string fichero, bool formato_fecha_1904, Mes12 mes12)
        {
            int f = 1;
            int c = 1;
            bool firstOnly = true;

            //Variables para datos pestaña Resumen
            int num_fact_estimadas_ind = 0;
            int num_fact_estimadas_agr = 0;
            int num_fact_ind = 0;
            int num_fact_agr = 0;
            double importe_estimado_ind = 0.0;
            double importe_facturado_ind = 0.0;
            double importe_estimado_agr = 0.0;
            double importe_facturado_agr = 0.0;
            int sin_flimpago_ind = 0;
            int sin_flimpago_agr = 0;


            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();

            

            List<EndesaEntity.facturacion.mes13.Seguimiento> lista = CargaSeguimientoIndividuales();
            Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>> dic = CargaSeguimientoFacturasIndividuales();

            FileInfo plantillaExcel =
                        new FileInfo(System.Environment.CurrentDirectory +
                        param.GetValue("plantilla_conversion"));


            
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(plantillaExcel);

            var workSheet = excelPackage.Workbook.Worksheets["INDIVIDUALES"];

            var headerCells = workSheet.Cells[1, 1, 1, 28];
            var headerFont = headerCells.Style.Font;


            f = 1;

            headerFont.Bold = true;
            try
            {
                #region Cabecera_Excel

                workSheet.Cells[f, c].Value = "ENTIDAD";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "LN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "EMPRESA TITULAR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "CLIENTE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "CCOUNIPS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "CUPSREE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "REFERENCIA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 203, 173)); c++;
                // workSheet.Cells[f, c].Value = "REFERENCIA_1"; c++;
                workSheet.Cells[f, c].Value = "SEC";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 203, 173)); c++;

                workSheet.Cells[f, c].Value = "CONTROL";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 203, 173)); c++;

                workSheet.Cells[f, c].Value = "Estimación Importe";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                workSheet.Cells[f, c].Value = "Estimación Base";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                workSheet.Cells[f, c].Value = "Estimación Impuestos";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                // workSheet.Cells[f, c].Value = "REF"; c++;
                workSheet.Cells[f, c].Value = "DiaF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                workSheet.Cells[f, c].Value = "Fecha vto estimada";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                // workSheet.Cells[f, c].Value = "DiaV(F+)"; c++;
                workSheet.Cells[f, c].Value = "seguimiento";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "CREFEREN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "SECFACTU";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "CFACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "fechaFact";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "fechaDes";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "fechaHas";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "FLIMPAGO";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "importeFact";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "Impuestos";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "FPTACOBR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                //workSheet.Cells[f, c].Value = "MES12";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                // workSheet.Cells[f, c].Value = "Comentarios"; c++;
                // workSheet.Cells[f, c].Value = "NIF_FACTURA"; c++;
                #endregion

                for (int i = 0; i < lista.Count; i++)
                {
                    c = 1;
                    f++;

                    #region Seguimiento_cabecera

                    
                    // workSheet.Cells[f, c].Value = lista[i].segmento; c++;
                    workSheet.Cells[f, c].Value = lista[i].entidad; c++;
                    workSheet.Cells[f, c].Value = lista[i].linea_negocio; c++;
                    workSheet.Cells[f, c].Value = lista[i].empresa_titular; c++;
                    workSheet.Cells[f, c].Value = lista[i].nif; c++;
                    workSheet.Cells[f, c].Value = lista[i].nombre_cliente; c++;
                    workSheet.Cells[f, c].Value = lista[i].cups13; c++;
                    workSheet.Cells[f, c].Value = lista[i].cups20; c++;
                    workSheet.Cells[f, c].Value = lista[i].referencia; c++;
                    // workSheet.Cells[f, c].Value = lista[i].nif_referencia; c++;
                    workSheet.Cells[f, c].Value = lista[i].secuencial; c++;
                    workSheet.Cells[f, c].Value = lista[i].control; c++;
                    //Importe estimado individual
                    workSheet.Cells[f, c].Value = lista[i].estimacion_importe; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    importe_estimado_ind += lista[i].estimacion_importe;
                    //Facturas estimadas individual
                    num_fact_estimadas_ind++;
                   
                    workSheet.Cells[f, c].Value = lista[i].estimacion_base; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = lista[i].estimacion_impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    // workSheet.Cells[f, c].Value = lista[i].nif_referencia; c++;

                    if (lista[i].diaf > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].diaf.AddDays(-1462) : lista[i].diaf;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;
                    if (lista[i].fecha_vto_estimada > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].fecha_vto_estimada.AddDays(-1462) : lista[i].fecha_vto_estimada;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    // workSheet.Cells[f, c].Value = lista[i].diav; c++;
                    #endregion

                    List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> o;
                    if (dic.TryGetValue(lista[i].cups20.ToUpper(), out o))
                    {
                        firstOnly = true;
                        for (int j = 0; j < o.Count; j++)
                        {
                            if (firstOnly)
                            {
                                #region Detalle_Factura

                                workSheet.Cells[f, c].Value = o[j].seguimiento_codigo; c++;
                                if (o[j].seguimiento_codigo == 10)
                                {
                                    num_fact_ind++;
                                    sin_flimpago_ind++;
                                }
                                else if (o[j].seguimiento_codigo == 11)
                                {
                                    num_fact_ind++;
                                }
                                workSheet.Cells[f, c].Value = o[j].creferen;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = o[j].secfactu; c++;
                                workSheet.Cells[f, c].Value = o[j].cfactura; c++;

                                if (o[j].ffactura > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactura.AddDays(-1462) : o[j].ffactura;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffactdes > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactdes.AddDays(-1462) : o[j].ffactdes;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffacthas > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffacthas.AddDays(-1462) : o[j].ffacthas;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].flimpago > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].flimpago.AddDays(-1462) : o[j].flimpago;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;


                                //Importe facturado individuales
                                workSheet.Cells[f, c].Value = o[j].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                                importe_facturado_ind += o[j].ifactura;
                                workSheet.Cells[f, c].Value = o[j].impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                                if (o[j].fptacobr > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].fptacobr.AddDays(-1462) : o[j].fptacobr;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;


                                //if (mes12.ExisteFacturaEnMes12(o[j].cfactura))
                                //    workSheet.Cells[f, c].Value = "Existe en Mes12"; c++;

                                //workSheet.Cells[f, c].Value = o[j].comentarios; c++;
                                //workSheet.Cells[f, c].Value = o[j].nif_factura; c++;

                                #endregion
                                firstOnly = false;
                            }
                            else
                            {
                                #region Seguimiento_cabecera_periodo_partido
                                c = 1;
                                f++;

                                //workSheet.Cells[f, c].Value = lista[i].segmento; c++;
                                workSheet.Cells[f, c].Value = lista[i].entidad; c++;
                                workSheet.Cells[f, c].Value = lista[i].linea_negocio; c++;
                                workSheet.Cells[f, c].Value = lista[i].empresa_titular; c++;
                                workSheet.Cells[f, c].Value = lista[i].nif; c++;
                                workSheet.Cells[f, c].Value = lista[i].nombre_cliente; c++;
                                workSheet.Cells[f, c].Value = lista[i].cups13; c++;
                                workSheet.Cells[f, c].Value = lista[i].cups20; c++;
                                workSheet.Cells[f, c].Value = lista[i].referencia; c++;
                                workSheet.Cells[f, c].Value = j + 1; c++;
                                workSheet.Cells[f, c].Value = lista[i].control; c++;
                                c++; // estimacion_importe
                                c++; // estimacion_base
                                c++; // estimacion_impuestos                           


                                if (lista[i].diaf > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].diaf.AddDays(-1462) : lista[i].diaf;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;
                                if (lista[i].fecha_vto_estimada > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].fecha_vto_estimada.AddDays(-1462) : lista[i].fecha_vto_estimada;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                // workSheet.Cells[f, c].Value = lista[i].diav; c++;
                                #endregion

                                #region Detalle_Factura

                                workSheet.Cells[f, c].Value = o[j].seguimiento_codigo; c++;

                                //Facturadas individuales
                                //Si código seguimiento 10: factura sin datos flimpago, 11: factura con datos flimpago
                                if (o[j].seguimiento_codigo == 10)
                                {
                                    num_fact_ind++;
                                    sin_flimpago_ind++;
                                }
                                else if (o[j].seguimiento_codigo == 11)
                                {
                                    num_fact_ind++;
                                }
                                
                                workSheet.Cells[f, c].Value = o[j].creferen;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = o[j].secfactu; c++;
                                workSheet.Cells[f, c].Value = o[j].cfactura; c++;

                                if (o[j].ffactura > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactura.AddDays(-1462) : o[j].ffactura;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffactdes > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactdes.AddDays(-1462) : o[j].ffactdes;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffacthas > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffacthas.AddDays(-1462) : o[j].ffacthas;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].flimpago > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].flimpago.AddDays(-1462) : o[j].flimpago;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                //Importe facturado individuales
                                workSheet.Cells[f, c].Value = o[j].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                                importe_facturado_ind += o[j].ifactura;
                               
                                workSheet.Cells[f, c].Value = o[j].impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                                if (o[j].fptacobr > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].fptacobr.AddDays(-1462) : o[j].fptacobr;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                //if (mes12.ExisteFacturaEnMes12(o[j].cfactura))
                                //    workSheet.Cells[f, c].Value = "Existe en Mes12"; c++;


                                // workSheet.Cells[f, c].Value = o[j].comentarios; c++;
                                // workSheet.Cells[f, c].Value = o[j].nif_factura; c++;

                                #endregion
                            }


                        }
                    }
                    else
                    {

                    }
                }

                var allCells = workSheet.Cells[1, 1, f, 28];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:Z1"].AutoFilter = true;
                allCells.AutoFitColumns();

                workSheet = excelPackage.Workbook.Worksheets["AGRUPADAS"];
                headerCells = workSheet.Cells[1, 1, 1, 31];
                headerFont = headerCells.Style.Font;

                lista = CargaSeguimientoAgrupadas();
                dic = CargaSeguimientoFacturasAgrupadas();
                f = 1;
                c = 1;

                #region Cabecera_Excel
                workSheet.Cells[f, c].Value = "ENTIDAD";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "LN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "EMPRESA TITULAR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "CLIENTE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "CCOUNIPS";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "CUPSREE";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); c++;

                workSheet.Cells[f, c].Value = "REFERENCIA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 203, 173)); c++;

                workSheet.Cells[f, c].Value = "SEC";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 203, 173)); c++;

                workSheet.Cells[f, c].Value = "CONTROL";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 203, 173)); c++;

                workSheet.Cells[f, c].Value = "Estimación Importe";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                workSheet.Cells[f, c].Value = "Estimación Base";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                workSheet.Cells[f, c].Value = "Estimación Impuestos";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                // workSheet.Cells[f, c].Value = "REF"; c++;
                workSheet.Cells[f, c].Value = "DiaF";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                workSheet.Cells[f, c].Value = "Fecha vto estimada";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(180, 198, 231)); c++;

                // workSheet.Cells[f, c].Value = "DiaV(F+)"; c++;
                workSheet.Cells[f, c].Value = "seguimiento";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "CREFEREN";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "SECFACTU";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "CFACTURA";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "fechaFact";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "fechaDes";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "fechaHas";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "FLIMPAGO";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "importeFact";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "Impuestos";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                workSheet.Cells[f, c].Value = "FPTACOBR";
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;

                //workSheet.Cells[f, c].Value = "MES12";
                //workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 176, 240)); c++;


                // workSheet.Cells[f, c].Value = "Comentarios"; c++;
                // workSheet.Cells[f, c].Value = "NIF_FACTURA"; c++;
                #endregion

                for (int i = 0; i < lista.Count; i++)
                {
                    c = 1;
                    f++;

                    #region Seguimiento_cabecera

                    // workSheet.Cells[f, c].Value = lista[i].segmento; c++;
                    workSheet.Cells[f, c].Value = lista[i].entidad; c++;
                    workSheet.Cells[f, c].Value = lista[i].linea_negocio; c++;
                    workSheet.Cells[f, c].Value = lista[i].empresa_titular; c++;
                    workSheet.Cells[f, c].Value = lista[i].nif; c++;
                    workSheet.Cells[f, c].Value = lista[i].nombre_cliente; c++;
                    workSheet.Cells[f, c].Value = lista[i].cups13; c++;
                    workSheet.Cells[f, c].Value = lista[i].cups20; c++;
                    workSheet.Cells[f, c].Value = lista[i].referencia; c++;
                    workSheet.Cells[f, c].Value = lista[i].secuencial; c++;
                    workSheet.Cells[f, c].Value = lista[i].control; c++;

                    //Importe estimado agrupado
                    workSheet.Cells[f, c].Value = lista[i].estimacion_importe;workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    importe_estimado_agr += lista[i].estimacion_importe;
                    //Facturas estimadas agrupadas
                    num_fact_estimadas_agr++; 
                    

                    workSheet.Cells[f, c].Value = lista[i].estimacion_base; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    workSheet.Cells[f, c].Value = lista[i].estimacion_impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                    // workSheet.Cells[f, c].Value = lista[i].nif_referencia; c++;

                    if (lista[i].diaf > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].diaf.AddDays(-1462) : lista[i].diaf;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;
                    if (lista[i].fecha_vto_estimada > DateTime.MinValue)
                    {
                        workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].fecha_vto_estimada.AddDays(-1462) : lista[i].fecha_vto_estimada;
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    }
                    c++;

                    //workSheet.Cells[f, c].Value = lista[i].diav; c++;
                    #endregion

                    List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> o;
                    if (dic.TryGetValue(lista[i].nif, out o))
                    {
                        firstOnly = true;
                        for (int j = 0; j < o.Count; j++)
                        {
                            if (firstOnly)
                            {
                                #region Detalle_Factura

                                workSheet.Cells[f, c].Value = o[j].seguimiento_codigo; c++;

                                //Facturadas agrupadas
                                //Si código seguimiento 10: factura sin datos flimpago, 11: factura con datos flimpago
                                if (o[j].seguimiento_codigo == 10)
                                {
                                    num_fact_agr++;
                                    sin_flimpago_agr++;
                                }
                                else if (o[j].seguimiento_codigo == 11)
                                {
                                    num_fact_agr++;
                                }

                                workSheet.Cells[f, c].Value = o[j].creferen;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = o[j].secfactu; c++;
                                workSheet.Cells[f, c].Value = o[j].cfactura; c++;

                                if (o[j].ffactura > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactura.AddDays(-1462) : o[j].ffactura;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffactdes > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactdes.AddDays(-1462) : o[j].ffactdes;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffacthas > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffacthas.AddDays(-1462) : o[j].ffacthas;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].flimpago > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].flimpago.AddDays(-1462) : o[j].flimpago;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                //Importe facturado agrupadas
                                workSheet.Cells[f, c].Value = o[j].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                                importe_facturado_agr += o[j].ifactura;
                                
                                workSheet.Cells[f, c].Value = o[j].impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                                if (o[j].fptacobr > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].fptacobr.AddDays(-1462) : o[j].fptacobr;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                //if (mes12.ExisteFacturaEnMes12(o[j].cfactura))
                                //    workSheet.Cells[f, c].Value = "Existe en Mes12"; c++;

                                // workSheet.Cells[f, c].Value = o[j].comentarios; c++;

                                #endregion
                                firstOnly = false;
                            }
                            else
                            {
                                #region Seguimiento_cabecera_periodo_partido
                                c = 1;
                                f++;

                                // workSheet.Cells[f, c].Value = lista[i].segmento; c++;
                                workSheet.Cells[f, c].Value = lista[i].entidad; c++;
                                workSheet.Cells[f, c].Value = lista[i].linea_negocio; c++;
                                workSheet.Cells[f, c].Value = lista[i].empresa_titular; c++;
                                workSheet.Cells[f, c].Value = lista[i].nif; c++;
                                workSheet.Cells[f, c].Value = lista[i].nombre_cliente; c++;
                                workSheet.Cells[f, c].Value = lista[i].cups13; c++;
                                workSheet.Cells[f, c].Value = lista[i].cups20; c++;
                                workSheet.Cells[f, c].Value = lista[i].referencia; c++;
                                // workSheet.Cells[f, c].Value = lista[i].nif_referencia; c++;
                                workSheet.Cells[f, c].Value = j + 1; c++;
                                workSheet.Cells[f, c].Value = lista[i].control; c++;
                                c++; // estimacion_importe
                                c++; // estimacion_base
                                c++; // estimacion_impuestos                           
                                // workSheet.Cells[f, c].Value = lista[i].nif_referencia; c++;

                                if (lista[i].diaf > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].diaf.AddDays(-1462) : lista[i].diaf;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;
                                if (lista[i].fecha_vto_estimada > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? lista[i].fecha_vto_estimada.AddDays(-1462) : lista[i].fecha_vto_estimada;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                // workSheet.Cells[f, c].Value = lista[i].diav; c++;
                                #endregion

                                #region Detalle_Factura

                                workSheet.Cells[f, c].Value = o[j].seguimiento_codigo; c++;
                                //Facturadas agrupadas
                                //Si código seguimiento 10: factura sin datos flimpago, 11: factura con datos flimpago
                                if (o[j].seguimiento_codigo == 10)
                                {
                                    num_fact_agr++;
                                    sin_flimpago_agr++;
                                }
                                else if (o[j].seguimiento_codigo == 11)
                                {
                                    num_fact_agr++;
                                }
                                workSheet.Cells[f, c].Value = o[j].creferen;
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                                workSheet.Cells[f, c].Value = o[j].secfactu; c++;
                                workSheet.Cells[f, c].Value = o[j].cfactura; c++;

                                if (o[j].ffactura > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactura.AddDays(-1462) : o[j].ffactura;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffactdes > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffactdes.AddDays(-1462) : o[j].ffactdes;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].ffacthas > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].ffacthas.AddDays(-1462) : o[j].ffacthas;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                if (o[j].flimpago > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].flimpago.AddDays(-1462) : o[j].flimpago;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;


                                //Importe facturado agrupadas
                                workSheet.Cells[f, c].Value = o[j].ifactura; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;
                                importe_facturado_agr += o[j].ifactura;
                                workSheet.Cells[f, c].Value = o[j].impuestos; workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0.00"; c++;

                                if (o[j].fptacobr > DateTime.MinValue)
                                {
                                    workSheet.Cells[f, c].Value = formato_fecha_1904 ? o[j].fptacobr.AddDays(-1462) : o[j].fptacobr;
                                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                }
                                c++;

                                //GUS: comentamos 02-10-2024, no hace falta
                                //if (mes12.ExisteFacturaEnMes12(o[j].cfactura))
                                //    workSheet.Cells[f, c].Value = "Existe en Mes12"; c++;

                                //workSheet.Cells[f, c].Value = o[j].comentarios; c++;

                                #endregion
                            }


                        }
                    }
                    else
                    {

                    }
                }

                allCells = workSheet.Cells[1, 1, f, 28];
                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:Z1"].AutoFilter = true;
                allCells.AutoFitColumns();

                workSheet = excelPackage.Workbook.Worksheets["RESUMEN"];
                
                workSheet.Cells[7, 3].Value = num_fact_estimadas_ind;
                workSheet.Cells[7, 4].Value = num_fact_ind;
                workSheet.Cells[7, 6].Value = importe_facturado_ind;
                workSheet.Cells[7, 8].Value = importe_estimado_ind;

                workSheet.Cells[11, 3].Value = num_fact_estimadas_agr;
                workSheet.Cells[11, 4].Value = num_fact_agr;
                workSheet.Cells[11, 6].Value = importe_facturado_agr;
                workSheet.Cells[11, 8].Value = importe_estimado_agr;

                workSheet.Cells[15, 3].Value = sin_flimpago_ind;
                workSheet.Cells[16, 3].Value = sin_flimpago_agr;
                

                excelPackage.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Genera excel seguimiento: " + ex.Message);
            }



        }

        private List<EndesaEntity.facturacion.mes13.Seguimiento> CargaSeguimientoIndividuales()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            List<EndesaEntity.facturacion.mes13.Seguimiento> lista = new List<EndesaEntity.facturacion.mes13.Seguimiento>();


            strSql = "select segmento, entidad, ln, empresa_titular, nif, cliente, ccounips, cupsree,"
                + " referencia, ref_cif, sec, control, estimacion_importe, estimacion_base, estimacion_impuestos,"
                + " ref, diaf, fecha_vto_estimada, diav"
                + " from fact.13_seg_individuales";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                EndesaEntity.facturacion.mes13.Seguimiento c = new EndesaEntity.facturacion.mes13.Seguimiento();

                if (r["segmento"] != System.DBNull.Value)
                    c.segmento = r["segmento"].ToString();
                if (r["entidad"] != System.DBNull.Value)
                    c.entidad = r["entidad"].ToString();
                if (r["ln"] != System.DBNull.Value)
                    c.linea_negocio = r["ln"].ToString();
                if (r["empresa_titular"] != System.DBNull.Value)
                    c.empresa_titular = r["empresa_titular"].ToString();
                if (r["nif"] != System.DBNull.Value)
                    c.nif = r["nif"].ToString();
                if (r["cliente"] != System.DBNull.Value)
                    c.nombre_cliente = r["cliente"].ToString();
                if (r["ccounips"] != System.DBNull.Value)
                    c.cups13 = r["ccounips"].ToString();
                if (r["cupsree"] != System.DBNull.Value)
                    c.cups20 = r["cupsree"].ToString();
                if (r["referencia"] != System.DBNull.Value)
                    c.referencia = r["referencia"].ToString();
                if (r["ref_cif"] != System.DBNull.Value)
                    c.nif_referencia = r["ref_cif"].ToString();
                if (r["sec"] != System.DBNull.Value)
                    c.secuencial = Convert.ToInt32(r["sec"]);
                if (r["control"] != System.DBNull.Value)
                    //c.control = Convert.ToInt32(r["control"]);
                    c.control = r["control"].ToString();
                if (r["estimacion_importe"] != System.DBNull.Value)
                    c.estimacion_importe = Convert.ToDouble(r["estimacion_importe"]);
                if (r["estimacion_base"] != System.DBNull.Value)
                    c.estimacion_base = Convert.ToDouble(r["estimacion_base"]);
                if (r["estimacion_impuestos"] != System.DBNull.Value)
                    c.estimacion_impuestos = Convert.ToDouble(r["estimacion_impuestos"]);
                if (r["ref"] != System.DBNull.Value)
                    c.nif_referencia = r["ref"].ToString();
                if (r["diaf"] != System.DBNull.Value)
                    c.diaf = Convert.ToDateTime(r["diaf"]);
                if (r["fecha_vto_estimada"] != System.DBNull.Value)
                    c.fecha_vto_estimada = Convert.ToDateTime(r["fecha_vto_estimada"]);
                if (r["diav"] != System.DBNull.Value)
                    c.diav = Convert.ToDateTime(r["diav"]);

                lista.Add(c);
            }
            db.CloseConnection();
            return lista;
        }
        private Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>> CargaSeguimientoFacturasIndividuales()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>> dic = new Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>>();
            strSql = "select seguimiento, seguimiento_codigo, cupsree, creferen, secfactu, cfactura, ffactura, ffactdes, ffacthas,"
                + " flimpago, fptacobr, ifactura, impuestos, comentarios, nif"
                + " from fact.13_seg_individuales_detalle where impuestos>=0 and ifactura > impuestos";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.mes13.FacturasSeguimiento c = new EndesaEntity.facturacion.mes13.FacturasSeguimiento();
                c.seguimiento = Convert.ToInt32(r["seguimiento"]);
                c.seguimiento_codigo = Convert.ToInt32(r["seguimiento_codigo"]);
                if (r["nif"] != System.DBNull.Value)
                    c.nif_factura = r["nif"].ToString();
                if (r["cupsree"] != System.DBNull.Value)
                    c.cupsree = r["cupsree"].ToString().ToUpper();
                if (r["creferen"] != System.DBNull.Value)
                    c.creferen = Convert.ToInt64(r["creferen"]);
                if (r["secfactu"] != System.DBNull.Value)
                    c.secfactu = Convert.ToInt32(r["secfactu"]);
                if (r["cfactura"] != System.DBNull.Value)
                    c.cfactura = r["cfactura"].ToString();
                if (r["ffactura"] != System.DBNull.Value)
                    c.ffactura = Convert.ToDateTime(r["ffactura"]);
                if (c.nif_factura.Substring(0, 2) != "NL")
                {
                    if (r["ffactdes"] != System.DBNull.Value)
                        c.ffactdes = Convert.ToDateTime(r["ffactdes"]);
                    if (r["ffacthas"] != System.DBNull.Value)
                        c.ffacthas = Convert.ToDateTime(r["ffacthas"]);
                }

                if (r["flimpago"] != System.DBNull.Value)
                    c.flimpago = Convert.ToDateTime(r["flimpago"]);
                if (r["fptacobr"] != System.DBNull.Value)
                    c.fptacobr = Convert.ToDateTime(r["fptacobr"]);
                if (r["ifactura"] != System.DBNull.Value)
                    c.ifactura = Convert.ToDouble(r["ifactura"]);
                if (r["impuestos"] != System.DBNull.Value)
                    c.impuestos = Convert.ToDouble(r["impuestos"]);
                if (r["comentarios"] != System.DBNull.Value)
                    c.comentarios = r["comentarios"].ToString();



                List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> o;
                if (!dic.TryGetValue(c.cupsree, out o))
                {
                    List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> lista = new List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>();
                    lista.Add(c);
                    dic.Add(c.cupsree, lista);
                }
                else
                {
                    o.Add(c);
                }

            }
            db.CloseConnection();
            return dic;
        }

        private List<EndesaEntity.facturacion.mes13.Seguimiento> CargaSeguimientoAgrupadas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            List<EndesaEntity.facturacion.mes13.Seguimiento> lista = new List<EndesaEntity.facturacion.mes13.Seguimiento>();


            strSql = "select segmento, entidad, ln, empresa_titular, nif, cliente, ccounips, cupsree,"
                + " referencia, sec, control, estimacion_importe, estimacion_base, estimacion_impuestos,"
                + " ref, diaf, fecha_vto_estimada, diav"
                + " from fact.13_seg_agrupadas";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                EndesaEntity.facturacion.mes13.Seguimiento c = new EndesaEntity.facturacion.mes13.Seguimiento();

                if (r["segmento"] != System.DBNull.Value)
                    c.segmento = r["segmento"].ToString();
                if (r["entidad"] != System.DBNull.Value)
                    c.entidad = r["entidad"].ToString();
                if (r["ln"] != System.DBNull.Value)
                    c.linea_negocio = r["ln"].ToString();
                if (r["empresa_titular"] != System.DBNull.Value)
                    c.empresa_titular = r["empresa_titular"].ToString();
                if (r["nif"] != System.DBNull.Value)
                    c.nif = r["nif"].ToString();

                if (r["cliente"] != System.DBNull.Value)
                    c.nombre_cliente = r["cliente"].ToString();
                if (r["ccounips"] != System.DBNull.Value)
                    c.cups13 = r["ccounips"].ToString().ToUpper();
                if (r["cupsree"] != System.DBNull.Value)
                    c.cups20 = r["cupsree"].ToString().ToUpper();
                if (r["referencia"] != System.DBNull.Value)
                    c.referencia = r["referencia"].ToString();
                if (r["sec"] != System.DBNull.Value)
                    c.secuencial = Convert.ToInt32(r["sec"]);
                if (r["control"] != System.DBNull.Value)
                    //c.control = Convert.ToInt32(r["control"]);
                    c.control = r["control"].ToString();
                if (r["estimacion_importe"] != System.DBNull.Value)
                    c.estimacion_importe = Convert.ToDouble(r["estimacion_importe"]);
                if (r["estimacion_base"] != System.DBNull.Value)
                    c.estimacion_base = Convert.ToDouble(r["estimacion_base"]);
                if (r["estimacion_impuestos"] != System.DBNull.Value)
                    c.estimacion_impuestos = Convert.ToDouble(r["estimacion_impuestos"]);
                if (r["ref"] != System.DBNull.Value)
                    c.nif_referencia = r["ref"].ToString();
                if (r["diaf"] != System.DBNull.Value)
                    c.diaf = Convert.ToDateTime(r["diaf"]);
                if (r["fecha_vto_estimada"] != System.DBNull.Value)
                    c.fecha_vto_estimada = Convert.ToDateTime(r["fecha_vto_estimada"]);
                if (r["diav"] != System.DBNull.Value)
                    c.diav = Convert.ToDateTime(r["diav"]);

                lista.Add(c);
            }
            db.CloseConnection();
            return lista;
        }
        private Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>> CargaSeguimientoFacturasAgrupadas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>> dic = 
                new Dictionary<string, List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>>();
            strSql = "select seguimiento, seguimiento_codigo, nif, creferen, secfactu, cfactura, ffactura, ffactdes, ffacthas,"
                + " seguimiento_codigo, flimpago, fptacobr, ifactura, impuestos, comentarios"
                + " from fact.13_seg_agrupadas_detalle where impuestos>=0 and ifactura>impuestos";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.mes13.FacturasSeguimiento c = new EndesaEntity.facturacion.mes13.FacturasSeguimiento();
                c.seguimiento = Convert.ToInt32(r["seguimiento"]);

                if (r["seguimiento_codigo"] != System.DBNull.Value)
                    c.seguimiento_codigo = Convert.ToInt32(r["seguimiento_codigo"]);

                if (r["nif"] != System.DBNull.Value)
                    c.nif = r["nif"].ToString();

                if (r["creferen"] != System.DBNull.Value)
                    c.creferen = Convert.ToInt64(r["creferen"]);

                if (r["secfactu"] != System.DBNull.Value)
                    c.secfactu = Convert.ToInt32(r["secfactu"]);

                if (r["cfactura"] != System.DBNull.Value)
                    c.cfactura = r["cfactura"].ToString();

                if (r["ffactura"] != System.DBNull.Value)
                    c.ffactura = Convert.ToDateTime(r["ffactura"]);

                if (r["ffactdes"] != System.DBNull.Value)
                    c.ffactdes = Convert.ToDateTime(r["ffactdes"]);

                if (r["ffacthas"] != System.DBNull.Value)
                    c.ffacthas = Convert.ToDateTime(r["ffacthas"]);

                if (r["flimpago"] != System.DBNull.Value)
                    c.flimpago = Convert.ToDateTime(r["flimpago"]);

                if (r["fptacobr"] != System.DBNull.Value)
                    c.fptacobr = Convert.ToDateTime(r["fptacobr"]);

                if (r["ifactura"] != System.DBNull.Value)
                    c.ifactura = Convert.ToDouble(r["ifactura"]);

                if (r["impuestos"] != System.DBNull.Value)
                    c.impuestos = Convert.ToDouble(r["impuestos"]);

                if (r["comentarios"] != System.DBNull.Value)
                    c.comentarios = r["comentarios"].ToString();

                List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> o;
                if (!dic.TryGetValue(c.nif, out o))
                {
                    List<EndesaEntity.facturacion.mes13.FacturasSeguimiento> lista = 
                        new List<EndesaEntity.facturacion.mes13.FacturasSeguimiento>();
                    lista.Add(c);
                    dic.Add(c.nif, lista);
                }
                else
                {
                    o.Add(c);

                }

            }
            db.CloseConnection();
            return dic;
        }

        public string UltimoFactoringCargado()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            string seguimiento = "";

            strSql = "select max(seguimiento_id) seguimiento from 13_seguimientos";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                if (r["seguimiento"] != System.DBNull.Value)
                    seguimiento = Convert.ToInt32(r["seguimiento"]).ToString();

            db.CloseConnection();

            return seguimiento;


        }

        public void BorradoTablasImportacionSeguimiento()
        {
            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "DELETE FROM 13_seguimientos";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM 13_seg_individuales";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM 13_seg_individuales_detalle";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM 13_seg_agrupadas";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "DELETE FROM 13_seg_agrupadas_detalle";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void ImportarAdjudicacion(string factoring, string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            int linea = 0;
            bool firstOnly = true;
            string cabecera = "";
            string strSql = "";
            List<EndesaEntity.facturacion.mes13.Seguimiento> lista = new List<EndesaEntity.facturacion.mes13.Seguimiento>();


            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;

            try
            {
                
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 100000; i++)
                {
                    linea = 1;
                    c = 1;

                    if (firstOnly)
                    {
                        for (int w = 1; w < 11; w++)
                            cabecera += workSheet.Cells[1, w].Value.ToString();


                        if (!EstructuraCorrecta(cabecera))
                        {
                            // this.hayError = true;
                            // this.descripcion_error = "La estructura del archivo excel no es la correcta.";

                            MessageBox.Show("La estructura del archivo no es la correcta",
                                "Estructura Incorrecta Excel",
                                MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    cabecera = "";
                    if (workSheet.Cells[f, 1].Value != null &&
                       workSheet.Cells[f, 2].Value != null &&
                       workSheet.Cells[f, 3].Value != null)
                    {

                        EndesaEntity.facturacion.mes13.Seguimiento s = new EndesaEntity.facturacion.mes13.Seguimiento();

                        // SEGMENTO ENTIDAD LN EMPRESATITULAR NIF CLIENTE CCOUNIPS CUPS REEREFERENCIA SEC
                        // s.segmento = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.entidad = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.linea_negocio = workSheet.Cells[f, c].Value.ToString().Substring(0, 1) == "L" ? "LUZ" : "GAS"; c++;
                        s.empresa_titular = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.nif = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.nombre_cliente = workSheet.Cells[f, c].Value.ToString(); c++;

                        if (workSheet.Cells[f, c].Value != null)
                            if (workSheet.Cells[f, c].Value.ToString().Trim() != "")
                                s.cups13 = workSheet.Cells[f, c].Value.ToString();
                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                        {
                            if (workSheet.Cells[f, c].Value.ToString().Trim() != "")
                                s.cups20 = workSheet.Cells[f, c].Value.ToString();
                        }
                        else if (s.cups13 != "")
                            s.cups20 = s.cups13;

                        c++;

                        s.referencia = workSheet.Cells[f, c].Value.ToString(); c++;
                        //s.nif_referencia = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.secuencial = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                        //s.control = Convert.ToInt32(workSheet.Cells[f, c].Value); c++;
                        s.control = workSheet.Cells[f, c].Value.ToString(); c++;
                        s.estimacion_importe = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        s.estimacion_base = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        s.estimacion_impuestos = Convert.ToDouble(workSheet.Cells[f, c].Value); c++;
                        s.diaf = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString()); c++;
                        s.diav = Convert.ToDateTime(workSheet.Cells[f, c].Value.ToString()); c++;

                        lista.Add(s);

                    }

                }

                firstOnly = true;
                id = 0;
                for (int i = 0; i < lista.Count; i++)
                {
                    id++;
                    if (firstOnly)
                    {
                        sb.Append("replace into 13_seguimientos (seguimiento_id, ln, empresa_titular, entidad, nif, cliente, ccounips,");
                        sb.Append(" cupsree, referencia, ref_cif, sec, control, estimacion_importe, estimacion_base, estimacion_impuestos,");
                        sb.Append(" diaf, diav) values ");
                        firstOnly = false;
                    }
                    sb.Append("(").Append(factoring).Append(",");
                    sb.Append("'").Append(lista[i].linea_negocio).Append("',");
                    sb.Append("'").Append(lista[i].empresa_titular).Append("',");
                    sb.Append("'").Append(lista[i].entidad).Append("',");
                    sb.Append("'").Append(lista[i].nif).Append("',");
                    sb.Append("'").Append(lista[i].nombre_cliente).Append("',");

                    if (lista[i].cups13 != null)
                        sb.Append("'").Append(lista[i].cups13).Append("',");
                    else
                        sb.Append("null,");

                    if (lista[i].cups20 != null)
                        sb.Append("'").Append(lista[i].cups20).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(lista[i].referencia).Append("',");

                    if (lista[i].nif_referencia != null)
                        sb.Append("'").Append(lista[i].nif_referencia).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append(lista[i].secuencial).Append(",");
                    //sb.Append(lista[i].control).Append(",");
                    sb.Append("'").Append(lista[i].control).Append("',");
                    sb.Append(lista[i].estimacion_importe.ToString().Replace(",", ".")).Append(",");
                    sb.Append(lista[i].estimacion_base.ToString().Replace(",", ".")).Append(",");
                    sb.Append(lista[i].estimacion_impuestos.ToString().Replace(",", ".")).Append(",");
                    sb.Append("'").Append(lista[i].diaf.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(lista[i].diav.ToString("yyyy-MM-dd")).Append("'),");


                    if (id == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        id = 0;
                    }

                }

                if (id > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    id = 0;
                }

                strSql = "REPLACE INTO 13_seg_individuales" +
                    " SELECT s.seguimiento_id, NULL AS segmento, s.entidad, s.LN, s.empresa_titular, s.nif," +
                    " s.cliente, s.ccounips, s.cupsree, s.referencia, concat(s.referencia,'_',s.nif) as ref_nif, s.sec, s.control, s.estimacion_importe," +
                    " s.estimacion_base, s.estimacion_impuestos, NULL ref," +
                    " s.diaf, s.diav, null" +
                    " FROM 13_seguimientos s WHERE s.referencia LIKE '" + factoring + "_NR%'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "REPLACE INTO 13_seg_agrupadas" +
                    " SELECT s.seguimiento_id, NULL AS segmento, s.entidad, s.LN, s.empresa_titular, s.nif," +
                    " s.cliente, s.ccounips, s.cupsree, s.referencia, concat(s.referencia,'_',s.nif) as ref_nif, s.sec, s.control, s.estimacion_importe," +
                    " s.estimacion_base, s.estimacion_impuestos, NULL ref," +
                    " s.diaf, s.diav, null" +
                    " FROM 13_seguimientos s WHERE s.referencia LIKE '" + factoring + "_AG%'";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();



                fs = null;
                excelPackage = null;

            }
            catch (Exception e)
            {
                MessageBox.Show("Error en la línea " + linea + " --> " + e.Message,
                  "Error en la importación de adjudicaciones",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private bool EstructuraCorrecta(string cabecera)
        {
            if (cabecera.ToUpper().Trim() == "ENTIDADLNEMPRESA TITULARNIFCLIENTECCOUNIPSCUPSREEREFERENCIASECCONTROL")
                return true;
            else
                return false;
        }

        public int TotalRegistrosSeg_Individuales()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int total_registros = 0;

            strSql = "select count(*) total_registros from 13_seg_individuales";
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

        public double TotalImporteEstimacion_Individuales()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            double total_importe = 0;

            strSql = "select sum(estimacion_importe) as estimacion_importe from 13_seg_individuales";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                total_importe = Convert.ToDouble(r["estimacion_importe"]);
            }

            db.CloseConnection();
            return total_importe;

        }


        public double TotalImporteEstimacion_Agrupadas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            double total_importe = 0;

            strSql = "select sum(estimacion_importe) as estimacion_importe from 13_seg_agrupadas";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                total_importe = Convert.ToDouble(r["estimacion_importe"]);
            }

            db.CloseConnection();
            return total_importe;

        }

        public int TotalRegistrosSeg_Agrupadas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int total_registros = 0;

            strSql = "select count(*) total_registros from 13_seg_agrupadas";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while(r.Read())
            {
                total_registros = Convert.ToInt32(r["total_registros"]);
            }

            db.CloseConnection();
            return total_registros;
        }

        // Función para anexar facturas a la tabla 13_seg_facturas_individuales o 13_seg_facturas_agrupadas de los clientes migrados a SAP - tipo="IND" --> INDIVIDUALES o tipo="AGR" --> AGRUPADAS o tipo="AGR_N0_SD" --> AGRUPADAS SD (TABLA sap_tfactura_n0)
        // f_factura_* --> Fecha factura
        // ffact* --> Fecha consumo 
        // Además añadimos las obligaciones (flimpago y fptacobr) a nueva tabla 13_obligaciones_sap
        public void AnexaFacturasSAP(DateTime f_factura_des, DateTime f_factura_has,
            DateTime ffactdes, DateTime ffacthas, string tipo)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sb_oblig = new StringBuilder();
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";
            string tabla_destino = "";
            string strTMP = "";

            MySQLDB dbmy;
            MySqlCommand commandmy;

            bool firstOnly = true;
            int j = 0;
            int k = 0;
            int totalRegistros = 0;

            try
            {

                switch (tipo)
                {
                    case "IND":
                        //Obtenemos las facturas individuales de SAP
                        strSql = QueryFactIndSAP(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        tabla_destino = "13_seg_facturas_individuales";
                        break;
                    case "IND_GAS":
                        //Obtenemos las facturas individuales de GAS - SAP
                        strSql = QueryFactIndSAP_GAS(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        tabla_destino = "13_seg_facturas_individuales";
                        break;
                    case "AGR":
                        //Obtenemos las facturas agrupadas de SAP
                        strSql = QueryFactAgrSAP(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        tabla_destino = "13_seg_facturas_agrupadas";
                        break;
                    case "AGR_no_PS":
                        //Obtenemos las facturas SD agrupadas de SAP, sin cruce con PS (sin CUPS)
                        strSql = QueryFactAgrSAP_no_PS(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        tabla_destino = "13_seg_facturas_agrupadas";
                        break;
                    case "AGR_N0_SD":
                        //Obtenemos las facturas agrupadas SD de SAP de la tabla ed_owner.sap_tfactura_n0
                        strSql = QueryFactAgrSAP_N0_SD(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        tabla_destino = "13_seg_facturas_agrupadas";
                        break;
                        
                    default:
                        //Por defecto obtenemos las facturas individuales de SAP
                        strSql = QueryFactIndSAP(f_factura_des, f_factura_has, ffactdes, ffacthas);
                        tabla_destino = "13_seg_facturas_individuales";
                        break;
                }

                //ficheroLog.Add(strSql);
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
                        sb.Append("REPLACE INTO " + tabla_destino);
                        sb.Append(" (ccounips, cemptitu, creferen, secfactu, cfactura, ffactura, ffactdes, ffacthas, ifactura, iva,");
                        sb.Append(" iimpues2, iimpues3, ise, hidrocarburos, tfactura, testfact, cnifdnic, indempre, cupsree, cfactrec,");
                        sb.Append(" tiponegocio, comentario_refact) values ");

                        sb_oblig.Append("REPLACE INTO 13_obligaciones_sap");
                        sb_oblig.Append(" (id_fact, flimpago, fptacobr) values  ");

                        firstOnly = false;
                    }

                    #region Campos
                    
                    //CCOUNIPS (CHAR)
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        sb.Append("('").Append(r["CCOUNIPS"].ToString()).Append("',");
                    else
                        sb.Append("(null,");
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
                    //ISE (DECIMAL)
                    if (r["ISE"] != System.DBNull.Value)
                        sb.Append(r["ISE"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //hidrocarburos (DECIMAL)
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
                    {
                        if(r["CUPSREE"].ToString().Length>20)
                            sb.Append("'").Append(r["CUPSREE"].ToString().Substring(0, 20)).Append("',");
                        else
                            sb.Append("'").Append(r["CUPSREE"].ToString()).Append("',");
                    }
                    else
                        sb.Append("null,");
                    //CFACTREC (CHAR)
                    if (r["CFACTREC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CFACTREC"].ToString()).Append("',");
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
                //ficheroLog.addError("AnexaFacturasSAP: " + ex.Message);
                Console.WriteLine("AnexaFacturasSAP: " + ex.Message);
            }
        }
        
        // Función para anexar facturas a la tabla 13_seg_facturas_individuales de los clientes de SIGAME (GAS), sin los datos FLIMPAGO
        public void AnexaFacturasSIGAME(DateTime f_factura_des, DateTime f_factura_has,
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
                     

                //ficheroLog.Add(strSql);
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
                        sb.Append("REPLACE INTO 13_seg_facturas_individuales");
                        sb.Append(" (ccounips, cemptitu, creferen, secfactu, cfactura, ffactura, ffactdes, ffacthas, ifactura, iva,");
                        sb.Append(" iimpues2, iimpues3, ise, hidrocarburos, tfactura, testfact, cnifdnic, indempre, cupsree, cfactrec,");
                        sb.Append(" tiponegocio, comentario_refact) values ");

                        firstOnly = false;
                    }

                    #region Campos

                    //CCOUNIPS (CHAR)
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        sb.Append("('").Append(r["CCOUNIPS"].ToString()).Append("',");
                    else
                        sb.Append("(null,");
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
                    //ISE (DECIMAL)
                    if (r["ISE"] != System.DBNull.Value)
                        sb.Append(r["ISE"].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                    //hidrocarburos (DECIMAL)
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
                    {
                        if (r["CUPSREE"].ToString().Length > 20)
                            sb.Append("'").Append(r["CUPSREE"].ToString().Substring(0, 20)).Append("',");
                        else
                            sb.Append("'").Append(r["CUPSREE"].ToString()).Append("',");
                    }
                    else
                        sb.Append("null,");
                    //CFACTREC (CHAR)
                    if (r["CFACTREC"] != System.DBNull.Value)
                        sb.Append("'").Append(r["CFACTREC"].ToString()).Append("',");
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
                //ficheroLog.addError("AnexaFacturasSAP: " + ex.Message);
                Console.WriteLine("AnexaFacturasSIGAME: " + ex.Message);
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
                //[2025-06-03 GUS]: Vamos a recortar a cups20 al igual que hicimos en Prevision.cs
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

        private string QueryFactIndSAP_GAS(DateTime f_factura_des, DateTime f_factura_has,
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
                + " c.tx_apell_cli as DAPERSOC, c.cd_nif_cif_cli as CNIFDNIC, 'OPERACIONES' as INDEMPRE, dc.cd_cups_gas_ext as CUPSREE,"
                + " null as CFACTREC, case dc.cd_linea_negocio when '001' then '1' when '002' then '2' else '1' end as CLINNEG, f.cd_tp_cli as CSEGMERC,"
                + " 0 as NUMLABOR, 'G' as TIPONEGOCIO, null as COMENTARIOS,  o.fh_ini_vencimiento as FLIMPAGO, o.fh_puesta_cobro as FPTACOBR "
                + " from ed_owner.t_ed_h_sap_facts f"
                + " inner join ed_owner.t_ed_h_sap_dc dc on f.cd_di=dc.cd_di"
              //  + " inner join ed_owner.t_ed_h_ps ps on ps.cd_cups_ext = dc.cd_cups_ext"
                + "	left join ed_owner.t_ed_f_clis c on f.cl_cli = c.cl_cli"
                + "	left join ed_owner.t_ed_h_obligaciones o on f.id_fact = o.cd_fiscal_fact"
                + " where f.im_factdo_con_iva > 0 and f.fh_ini_fact >='" + ffactdes.ToString("yyyy-MM-dd") + "' and f.fh_fin_fact <= '" + ffacthas.ToString("yyyy-MM-dd") + "'"
                + " and f.fh_fact >='" + f_factura_des.ToString("yyyy-MM-dd") + "' and f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " and f.cl_empr = '006'"
                + " and dc.cd_linea_negocio = '002'"
              //  + " and ( (dc.cd_empr ='PT1Q' and ps.cd_tp_tension in ('MT','AT','MAT')) or (dc.cd_empr='ES21') )"
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
                //+ " case f.cd_tp_fact when 'AG' then '6' when 'SD' then '5' when 'CO' then '1'	when 'MC' then '2' else '1'end as TFACTURA,"
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
                + " inner join ed_owner.t_ed_h_ps ps on ps.cd_cups_ext = dc.cd_cups_ext"
                + "	left join ed_owner.t_ed_f_clis c on f.cl_cli = c.cl_cli"
                + "	left join ed_owner.t_ed_h_obligaciones o on f.id_fact = o.cd_fiscal_fact"
                + " where f.im_factdo_con_iva > 0 and f.fh_ini_fact is null and f.fh_fin_fact is null"
                + " and f.fh_fact >='" + f_factura_des.ToString("yyyy-MM-dd") + "' and f.fh_fact <= '" + f_factura_has.ToString("yyyy-MM-dd") + "'"
                + " and f.cl_empr = '006' and ( (dc.cd_empr ='PT1Q' and ps.cd_tp_tension in ('MT','AT','MAT')) or (dc.cd_empr='ES21') )"
                + " and f.cd_est_fact not in ('A','X','S')"
                + " and f.id_fact is not null"
               // + " and ps.lg_migrado_sap ='S'"
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
