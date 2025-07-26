using EndesaBusiness.servidores;
using EndesaEntity.cnmc.gas.V25_2019_12_17;
using Microsoft.Graph;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class Reubicaciones
    {
        EndesaBusiness.sigame.Addendas addendas;
        EndesaBusiness.sigame.SIGAME inventario_gas;
        cnmc.CNMC cnmc;
        List<EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26> lista_reubicaciones;
        public List<EndesaEntity.cnmc.gas.V25_2019_12_17.Informe_Reubicaciones> informe { get; set; }
        public List<EndesaEntity.cnmc.gas.V25_2019_12_17.Cargas> cargas { get; set; }

        public Reubicaciones(DateTime fd, DateTime fh, bool mostrar_con_tarifas_iguales)
        {
            addendas = new sigame.Addendas(fd, fh);
            inventario_gas = new sigame.SIGAME(fd, fh);
            cnmc = new cnmc.CNMC();
            CargaReubicaciones();
            CargaCargas();
            Crea_Informe(fd, fh, mostrar_con_tarifas_iguales);
            
            
        }

        public Reubicaciones()
        {
            
        }

        public void Crea_Informe(DateTime fd, DateTime fh, bool mostrar_con_tarifas_iguales)
        {
            informe = new List<Informe_Reubicaciones>();

            foreach(EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26 p in lista_reubicaciones)
            {

                inventario_gas.GetContrato(p.a1226.cups);

                if((inventario_gas.Tarifa(p.a1226.cups) != cnmc.Tarifa_CNMC_a_Tarifa_SIGAME(p.a1226.tolltype)) || mostrar_con_tarifas_iguales)
                {
                    //EndesaEntity.cnmc.gas.V25_2019_12_17.Informe_Reubicaciones c =
                    //    new Informe_Reubicaciones();

                    //c.cups = p.a1226.cups;
                    //c.tipo = "CONTRATO";
                    //c.tarifa_sigame = inventario_gas.tarifa;
                    //c.tarifa_reubicacion = cnmc.Tarifa_CNMC_a_Tarifa_SIGAME(p.a1226.tolltype);
                    //c.fecha_desde = inventario_gas.fecha_inicio;
                    //c.fecha_hasta = inventario_gas.fecha_fin;

                    //informe.Add(c);

                    List<EndesaEntity.contratacion.gas.Addenda> 
                        lista_addendas = addendas.GetAddendasReunibacionesSinFechaFin(p.a1226.cups, p.a1226.transfereffectivedate);

                    if(lista_addendas != null)
                        foreach (EndesaEntity.contratacion.gas.Addenda pp in lista_addendas)
                        {
                            if ((pp.tarifa_peaje != cnmc.Tarifa_CNMC_a_Tarifa_SIGAME(p.a1226.tolltype)) || mostrar_con_tarifas_iguales)
                            {
                                Informe_Reubicaciones c = new Informe_Reubicaciones();

                                c.cups = p.a1226.cups;
                                c.tipo = "ADDENDA";
                                c.tarifa_sigame = pp.tarifa_peaje;
                                c.tarifa_reubicacion = cnmc.Tarifa_CNMC_a_Tarifa_SIGAME(p.a1226.tolltype);
                                c.fecha_desde = pp.fecha_desde;
                                c.fecha_hasta = pp.fecha_hasta;
                                c.fecha_reubicacion = p.a1226.transfereffectivedate;

                                informe.Add(c);
                            }
                        }

                    
                    lista_addendas = addendas.GetAddendasReunibacionesConFechaFin(p.a1226.cups, p.a1226.transfereffectivedate);

                    if (lista_addendas != null)
                        foreach (EndesaEntity.contratacion.gas.Addenda pp in lista_addendas)
                        {
                            if ((pp.tarifa_peaje != cnmc.Tarifa_CNMC_a_Tarifa_SIGAME(p.a1226.tolltype)) || mostrar_con_tarifas_iguales)
                            {
                                Informe_Reubicaciones c = new Informe_Reubicaciones();

                                c.cups = p.a1226.cups;
                                c.tipo = "ADDENDA";
                                c.tarifa_sigame = pp.tarifa_peaje;
                                c.tarifa_reubicacion = cnmc.Tarifa_CNMC_a_Tarifa_SIGAME(p.a1226.tolltype);
                                c.fecha_desde = pp.fecha_desde;
                                c.fecha_hasta = pp.fecha_hasta;
                                c.fecha_reubicacion = p.a1226.transfereffectivedate;

                                informe.Add(c);
                            }
                        }

                }
            }

        }

        public void CargaReubicaciones()
        {
            lista_reubicaciones = new List<Proceso_A12_26>();

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            strSql = "select id_a1226, line, reqdate, reqhour, atrcode, cups, nationality, documenttype, documentnum, firstname,"
                + " familyname1, familyname2, telephone, fax, newcustomer, email, streettype, street, streetnumber, portal,"
                + " staircase, floor, door, province, city, zipcode, tolltype, qdgranted, qhgranted, singlenomination,"
                + " transfereffectivedate, finalclientyearlyconsumption, netsituation, outgoingpressuregranted, lastinspectionsdate,"
                + " lastinspectionsresult, readingtype, rentingamount, rentingperiodicity, canonircamount, canonircperiodicity,"
                + " canonircforlife, canonircdate, canonircmonth, othersamount, othersperiodicity, readingperiodicitycode,"
                + " transporter, transnet, gasusetype, caecode, communicationreason, titulartype, regularaddress,"
                + " created_by, created_date, last_update_by, last_update_date"
                + " from cont.cnmc_a1226 group by cups, reqdate ";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26 c =
                    new Proceso_A12_26();

                c.a1226 = new A1226();

                if (r["cups"] != System.DBNull.Value)
                    c.a1226.cups = r["cups"].ToString();
                if (r["tolltype"] != System.DBNull.Value)
                        c.a1226.tolltype = r["tolltype"].ToString();
                if (r["transfereffectivedate"] != System.DBNull.Value)
                    c.a1226.transfereffectivedate = Convert.ToDateTime(r["transfereffectivedate"]);

                lista_reubicaciones.Add(c);

            }
            db.CloseConnection();             

        }

        public void CargaCargas()
        {
            cargas = new List<Cargas>();

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            strSql = "select fecha_carga, num_archivos, fecha_minima_solicitud, fecha_maxima_solicitud,"
                + " created_by, created_date, last_update_by, last_update_date"
                + " from cont.atrgas_cargas_reubicaciones";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                Cargas c = new Cargas();                

                if (r["fecha_carga"] != System.DBNull.Value)
                    c.fecha_carga = Convert.ToDateTime(r["fecha_carga"].ToString());
                if (r["num_archivos"] != System.DBNull.Value)
                    c.num_archivos = r["num_archivos"].ToString();
                if (r["fecha_minima_solicitud"] != System.DBNull.Value)
                    c.fecha_minima_solicitud = r["fecha_minima_solicitud"].ToString();
                if (r["fecha_maxima_solicitud"] != System.DBNull.Value)
                    c.fecha_maxima_solicitud = r["fecha_maxima_solicitud"].ToString();

                cargas.Add(c);

            }
            db.CloseConnection();

        }

        public void GuardaXML(string file, EndesaEntity.cnmc.gas.V25_2019_12_17.Proceso_A12_26 xml)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int id = 0;

            try
            {
                if (firstOnly)
                {
                    id = GetIDFile(file);

                    GuardaHeading(id, xml.heading);
                    GuardaCounter(id, xml.a1226.counterlist);

                    sb.Append("replace into cnmc_a1226");
                    sb.Append("(id_a1226, line, reqdate, reqhour, atrcode, cups, nationality, documenttype,");
                    sb.Append(" documentnum, firstname, familyname1, familyname2, telephone, fax, newcustomer,");
                    sb.Append(" email, streettype, street, streetnumber, portal, staircase, floor, door, province,");
                    sb.Append(" city, zipcode, tolltype, qdgranted, qhgranted, singlenomination, transfereffectivedate,");
                    sb.Append(" finalclientyearlyconsumption, netsituation, outgoingpressuregranted, lastinspectionsdate,");
                    sb.Append(" lastinspectionsresult, readingtype, rentingamount, rentingperiodicity, canonircamount,");
                    sb.Append(" canonircperiodicity, canonircforlife, canonircdate, canonircmonth, othersamount, othersperiodicity,");
                    sb.Append(" readingperiodicitycode, transporter, transnet, gasusetype, caecode, communicationreason, titulartype,");
                    sb.Append(" regularaddress, created_by, created_date) values ");
                    firstOnly = false;
                }

                #region fields

                sb.Append("(").Append(id).Append(",");
                sb.Append(1).Append(",");
                sb.Append("'").Append(xml.a1226.reqdate.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(xml.a1226.reqhour.ToString("HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.a1226.atrcode).Append("',");
                sb.Append("'").Append(xml.a1226.cups).Append("',");

                if (xml.a1226.nationality != null)
                    sb.Append("'").Append(xml.a1226.nationality).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.documenttype != null)
                    sb.Append("'").Append(xml.a1226.documenttype).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.documentnum != null)
                    sb.Append("'").Append(xml.a1226.documentnum).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.firstname != null)
                    sb.Append("'").Append(xml.a1226.firstname).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.familyname1 != null)
                    sb.Append("'").Append(xml.a1226.familyname1).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.familyname2 != null)
                    sb.Append("'").Append(xml.a1226.familyname2).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.telephone != null)
                    sb.Append("'").Append(xml.a1226.telephone).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.fax != null)
                    sb.Append("'").Append(xml.a1226.fax).Append("',");
                else
                    sb.Append("null,");                

                if (xml.a1226.newcustomer != null)
                    sb.Append("'").Append(xml.a1226.newcustomer).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.email != null)
                    sb.Append("'").Append(xml.a1226.email).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.streettype != null)
                    sb.Append("'").Append(xml.a1226.streettype).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.street != null)
                    sb.Append("'").Append(xml.a1226.street).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.streetnumber != null)
                    sb.Append("'").Append(xml.a1226.streetnumber).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.portal != null)
                    sb.Append("'").Append(xml.a1226.portal).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.staircase != null)
                    sb.Append("'").Append(xml.a1226.staircase).Append("',");
                else
                    sb.Append("null,");
                
                if (xml.a1226.floor != null)
                    sb.Append("'").Append(xml.a1226.floor).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.door != null)
                    sb.Append("'").Append(xml.a1226.door).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.province != null)
                    sb.Append("'").Append(xml.a1226.province).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.city != null)
                    sb.Append("'").Append(xml.a1226.city).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.zipcode != null)
                    sb.Append("'").Append(xml.a1226.zipcode).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.tolltype != null)
                    sb.Append("'").Append(xml.a1226.tolltype).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.qdgranted != 0)
                    sb.Append(xml.a1226.qdgranted.ToString().Replace(",",".")).Append(",");
                else
                    sb.Append("null,");
                

                if (xml.a1226.qhgranted != 0)
                    sb.Append(xml.a1226.qhgranted).Append(",");
                else
                    sb.Append("null,");

                if (xml.a1226.singlenomination != null)
                    sb.Append("'").Append(xml.a1226.singlenomination).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.transfereffectivedate > DateTime.MinValue)
                    sb.Append("'").Append(xml.a1226.transfereffectivedate.ToString("yyyy-MM-dd")).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.finalclientyearlyconsumption != 0)
                    sb.Append(xml.a1226.finalclientyearlyconsumption).Append(",");
                else
                    sb.Append("null,");

                if (xml.a1226.netsituation != null)
                    sb.Append("'").Append(xml.a1226.netsituation).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.outgoingpressuregranted != 0)
                    sb.Append(xml.a1226.outgoingpressuregranted.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append("null,");

                if (xml.a1226.lastinspectionsdate > DateTime.MinValue)
                    sb.Append("'").Append(xml.a1226.lastinspectionsdate.ToString("yyyy-MM-dd")).Append("',");
                else
                    sb.Append("null,");
                

                if (xml.a1226.lastinspectionsresult != null)
                    sb.Append("'").Append(xml.a1226.lastinspectionsresult).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.readingtype != null)
                    sb.Append("'").Append(xml.a1226.readingtype).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.rentingamount != 0)
                    sb.Append(xml.a1226.rentingamount.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append("null,");

                if (xml.a1226.rentingperiodicity != null)
                    sb.Append("'").Append(xml.a1226.rentingperiodicity).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.canonircamount != 0)
                    sb.Append(xml.a1226.canonircamount.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append("null,");

                if (xml.a1226.canonircperiodicity != null)
                    sb.Append("'").Append(xml.a1226.canonircperiodicity).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.canonircforlife != null)
                    sb.Append("'").Append(xml.a1226.canonircforlife).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.canonircdate > DateTime.MinValue)
                    sb.Append("'").Append(xml.a1226.canonircdate.ToString("yyyy-MM-dd")).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.canonircmonth != null)
                    sb.Append("'").Append(xml.a1226.canonircmonth).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.othersamount != 0)
                    sb.Append(xml.a1226.othersamount.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append("null,");

                if (xml.a1226.othersperiodicity != null)
                    sb.Append("'").Append(xml.a1226.othersperiodicity).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.readingperiodicitycode != null)
                    sb.Append("'").Append(xml.a1226.readingperiodicitycode).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.transporter != null)
                    sb.Append("'").Append(xml.a1226.transporter).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.transnet != null)
                    sb.Append("'").Append(xml.a1226.transnet).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.gasusetype != null)
                    sb.Append("'").Append(xml.a1226.gasusetype).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.caecode != null)
                    sb.Append("'").Append(xml.a1226.caecode).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.communicationreason != null)
                    sb.Append("'").Append(xml.a1226.communicationreason).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.titulartype != null)
                    sb.Append("'").Append(xml.a1226.titulartype).Append("',");
                else
                    sb.Append("null,");

                if (xml.a1226.regularaddress != null)
                    sb.Append("'").Append(xml.a1226.regularaddress).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception ex)
            {

            }
        }

        private void GuardaHeading(int id, EndesaEntity.cnmc.gas.V25_2019_12_17.Heading heading)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            sb.Append("replace into cnmc_a1226_heading");
            sb.Append("(id, dispatchingcode, dispatchingcompany, destinycompany, communicationsdate,");
            sb.Append(" communicationshour, processcode, messagetype,");
            sb.Append(" created_by, created_date) values ");


            sb.Append("(").Append(id).Append(",");
            sb.Append("'").Append(heading.dispatchingcode).Append("',");
            sb.Append("'").Append(heading.dispatchingcompany).Append("',");
            sb.Append("'").Append(heading.destinycompany).Append("',");
            sb.Append("'").Append(heading.communicationsdate.ToString("yyyy-MM-dd")).Append("',");
            sb.Append("'").Append(heading.communicationshour.ToString("HH:mm:ss")).Append("',");
            sb.Append("'").Append(heading.processcode).Append("',");
            sb.Append("'").Append(heading.messagetype).Append("',");

            sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        private void GuardaCounter(int id, EndesaEntity.cnmc.gas.V25_2019_12_17.CounterList counterList)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int linea = 0;

            foreach(EndesaEntity.cnmc.gas.V25_2019_12_17.Counter p in counterList.counterlist)
            {
                if (firstOnly)
                {
                    sb.Append("replace into cnmc_a1226_counter");
                    sb.Append("(id_a1226, line, countermodel, countertype, counternumber,");
                    sb.Append(" counterproperty, reallecture, counterpressure,");
                    sb.Append(" created_by, created_date) values ");
                    firstOnly = false;
                }

                linea++;

                sb.Append("(").Append(id).Append(",");
                sb.Append(linea).Append(",");

                if(p.countermodel != null)
                    sb.Append("'").Append(p.countermodel).Append("',");
                else
                    sb.Append("null,");

                if (p.countertype != null)
                    sb.Append("'").Append(p.countertype).Append("',");
                else
                    sb.Append("null,");

                if (p.counternumber != null)
                    sb.Append("'").Append(p.counternumber).Append("',");
                else
                    sb.Append("null,");

                if (p.counterproperty != null)
                    sb.Append("'").Append(p.counterproperty).Append("',");
                else
                    sb.Append("null,");

                if (p.reallecture != 0)
                    sb.Append(p.reallecture).Append(",");
                else
                    sb.Append("null,");

                if (p.reallecture != 0)
                    sb.Append(p.counterpressure.ToString().Replace(",",".")).Append(",");
                 else
                    sb.Append("null,");

                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
            }

            

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        private int GetIDFile(string file)
        {
            // Nos da el ID del fichero XML
            // Si no existe lo inserta
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int id = 0;
            bool encontrado = false;

            strSql = "select id from cnmc_a1226_files where"
                + " filename = '" + file + "'";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                id = Convert.ToInt32(r["id"]);
                encontrado = true;
            }
            db.CloseConnection();
            if (encontrado)
                return id;
            else
            {
                id = GetLastIDFile() + 1;

                strSql = "insert into cnmc_a1226_files (id, filename,"
                    + " created_by, created_date) values"
                    + " (" + id + ","
                    + "'" + file + "',"
                    + "'" + System.Environment.UserName.ToUpper() + "',"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                return id;
            }
        }

        private int GetLastIDFile()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int id = 0;            

            strSql = "select max(id) id from cnmc_a1226_files";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["id"] != System.DBNull.Value)
                    id = Convert.ToInt32(r["id"]);
            }
            db.CloseConnection();

            if (id == 0)
                return 0;
            else
                return id;

                
        }

        public bool ExisteCUPS(string cups)
        {
            return addendas.ExisteCUPS(cups) || inventario_gas.ExisteCUPS(cups);
        }

        public void GeneraInformeExcel(string fichero)
        {
            int f = 1;
            int c = 1;


            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Reubicaciones");

            var headerCells = workSheet.Cells[1, 1, 1, 6];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "CUPS";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Tipo";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Tarifa SIGAME";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Tarifa Reubicación";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Fecha Desde";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            workSheet.Cells[f, c].Value = "Fecha Hasta";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            c++;

            foreach(EndesaEntity.cnmc.gas.V25_2019_12_17.Informe_Reubicaciones p in informe)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.cups;
                c++;

                workSheet.Cells[f, c].Value = p.tipo;
                c++;

                workSheet.Cells[f, c].Value = p.tarifa_sigame;
                c++;

                workSheet.Cells[f, c].Value = p.tarifa_reubicacion;
                c++;

                if (p.fecha_desde > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.fecha_desde;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                if (p.fecha_hasta > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.fecha_hasta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

            }

            var allCells = workSheet.Cells[1, 1, f, c];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:F1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();

        }

        public void InsertaRegistroSinArchivos()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "replace into atrgas_cargas_reubicaciones"
                    + " (fecha_carga, num_archivos, fecha_minima_solicitud, fecha_maxima_solicitud,"
                    + " created_by, created_date) values"
                    + " ('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + " 'No hay archivos','No hay archivos','No hay archivos',"
                    + "'" + System.Environment.UserName.ToUpper() + "',"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Insertar registro sin archivo",
                  MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void InsertaCarga(string num_archivos, string feha_minima_solicitud, string fecha_maxima_solicitud)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "replace into atrgas_cargas_reubicaciones"
                    + " (fecha_carga, num_archivos, fecha_minima_solicitud, fecha_maxima_solicitud,"
                    + " created_by, created_date) values"
                    + " ('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + "'" + num_archivos + "',"
                    + "'" + feha_minima_solicitud + "',"
                    + "'" + fecha_maxima_solicitud + "',"
                    + "'" + System.Environment.UserName.ToUpper() + "',"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Inserta carga",
                  MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void UpdateFecha(DateTime fecha_carga, string fecha_minima, string fecha_maxima)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update atrgas_cargas_reubicaciones set"
                    + " fecha_minima_solicitud = '" + fecha_minima + "',"
                    + " fecha_maxima_solicitud = '" + fecha_maxima + "',"
                    + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                    + " where fecha_carga = '" + fecha_carga.ToString("yyyy-MM-dd HH:mm:ss") + "'";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Actualiza fechas",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
