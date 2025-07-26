using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;


namespace EndesaBusiness.cartera
{
    public class Cartera_SalesForce : EndesaEntity.contratacion.Cartera
    {
        List<EndesaEntity.contratacion.Cartera> l_p;
        Dictionary<string, EndesaEntity.contratacion.Cartera> dic;
        Dictionary<string, EndesaEntity.contratacion.Cartera> dic_pt;
        private utilidades.Param param;
        logs.Log ficheroLog;
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        public Cartera_SalesForce()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_SalesForceCartera");
            param = new utilidades.Param("salesforce_cartera_param", servidores.MySQLDB.Esquemas.CON);
            l_p = new List<EndesaEntity.contratacion.Cartera>();
            ss_pp = new utilidades.Seguimiento_Procesos();
        }

        public Cartera_SalesForce(List<string> lista_nifs)
        {
            param = new utilidades.Param("salesforce_cartera_param", servidores.MySQLDB.Esquemas.CON);
            dic = Carga(lista_nifs);
        }

        private Dictionary<string, EndesaEntity.contratacion.Cartera> Carga(List<string> lista_nifs)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;

            Dictionary<string, EndesaEntity.contratacion.Cartera> d =
                new Dictionary<string, EndesaEntity.contratacion.Cartera>();

            try
            {

                strSql = "SELECT c.cups20, c.nombre_cuenta,"
                    + " c.nif, c.gestor, c.email_gestor, c.posicion, c.segmento,"
                    + " c.subdireccion, c.subdirector, c.territorio,"
                    + " c.responsable_territorial, c.zona, c.responsable_zona,"
                    + " rt.email_gestor as email_responsable_rt"
                    + " FROM salesforce_cartera c"
                    + " left outer join salesforce_cartera_rt rt on"
                    + " rt.responsable_territorial = c.responsable_territorial"
                    + " where c.nif in (";

                for (int i = 0; i < lista_nifs.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista_nifs[i] + "'";
                        firstOnly = false;
                    }
                    else
                    {
                        strSql += " ,'" + lista_nifs[i] + "'";
                    }

                }

                strSql += ")"
                    + " and estado_contrato = 'EN VIGOR' group by nif";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Cartera c = new EndesaEntity.contratacion.Cartera();                    
                    c.cups22 = r["cups20"].ToString();

                    if (r["nombre_cuenta"] != System.DBNull.Value)
                        c.nombre_cuenta = r["nombre_cuenta"].ToString();

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["gestor"] != System.DBNull.Value)
                        c.gestor = r["gestor"].ToString();

                    if (r["email_gestor"] != System.DBNull.Value)
                        c.email_gestor = r["email_gestor"].ToString();

                    if (r["posicion"] != System.DBNull.Value)
                        c.posicion = r["posicion"].ToString();

                    if (r["subdireccion"] != System.DBNull.Value)
                        c.subdireccion = r["subdireccion"].ToString();

                    if (r["subdirector"] != System.DBNull.Value)
                        c.subdirector = r["subdirector"].ToString();

                    if (r["responsable_territorial"] != System.DBNull.Value)
                        c.responsable_territorial = r["responsable_territorial"].ToString();

                    if (r["zona"] != System.DBNull.Value)
                        c.zona = r["zona"].ToString();

                    if (r["email_responsable_rt"] != System.DBNull.Value)
                        c.email_responsable_rt = r["email_responsable_rt"].ToString();

                    if (r["responsable_zona"] != System.DBNull.Value)
                        c.responsable_zona = r["responsable_zona"].ToString();

                    if (r["segmento"] != System.DBNull.Value)
                        c.segmento = r["segmento"].ToString();

                    EndesaEntity.contratacion.Cartera o;
                    if (!d.TryGetValue(c.nif, out o))
                        d.Add(c.nif, c);

                }
                db.CloseConnection();
                return d;


            }
            catch (Exception e)
            {               

                MessageBox.Show(e.Message,
                "Cartera - Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }

        }

        public void CopiaClientesCartera_RedShift()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            servidores.MySQLDB mdb;
            MySqlCommand mcommand;
            StringBuilder sb = new StringBuilder();
            string strSql = "";
            int total = 0;
            int totalReg = 0;
            int registrosLeidos = 0;
            bool firstOnly = true;

            try
            {

                ss_pp.Update_Fecha_Inicio("Contratación", "Copia ed_owner.t_ed_f_clis", "Copia ed_owner.t_ed_f_clis");

                strSql = "select count(*) TOTAL_REGISTROS"
                    + " from ed_owner.t_ed_f_clis c"
                    + " where c.cd_tp_cli = 'Y' or c.cd_tp_cli = 'N'";
                Console.WriteLine(strSql);
                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                    total = Convert.ToInt32(r["TOTAL_REGISTROS"]);
                db.CloseConnection();


                strSql = "select c.cd_nif_cif_cli, c.tx_apell_cli, c.nombre_gestor, c.de_posicion,"
                    + " c.cd_territorio1, c.cd_zona, c.de_subdir_territor_b2b, c.de_resp_territorio,"
                    + " c.de_resp_zona, c.gestor_kam, c.de_ambto, c.fec_act"
                    + " from ed_owner.t_ed_f_clis c"
                    + " where c.cd_tp_cli = 'Y' or c.cd_tp_cli = 'N'";
                Console.WriteLine(strSql);
                Console.WriteLine();
                ficheroLog.Add(strSql);
                db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    registrosLeidos++;
                    totalReg++;
                    if (firstOnly)
                    {
                        sb.Append("replace into salesforce_t_ed_f_clis (cd_nif_cif_cli, tx_apell_cli, ");
                        sb.Append("nombre_gestor, de_posicion, cd_territorio1, cd_zona, de_subdir_territor_b2b, ");
                        sb.Append("de_resp_territorio, de_resp_zona, gestor_kam, de_ambto, fec_act, ");
                        sb.Append("f_ult_mod) values ");

                        firstOnly = false;
                    }

                    sb.Append("('").Append(r["cd_nif_cif_cli"].ToString()).Append("',");

                    if(r["tx_apell_cli"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["tx_apell_cli"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["nombre_gestor"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["nombre_gestor"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_posicion"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["de_posicion"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_territorio1"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["cd_territorio1"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["cd_zona"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["cd_zona"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_subdir_territor_b2b"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["de_subdir_territor_b2b"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_resp_territorio"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["de_resp_territorio"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_resp_zona"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["de_resp_zona"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["gestor_kam"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["gestor_kam"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["de_ambto"] != System.DBNull.Value)
                        sb.Append("'").Append(EndesaBusiness.utilidades.FuncionesTexto.RT(r["de_ambto"].ToString())).Append("',");
                    else
                        sb.Append("null,");

                    if (r["fec_act"] != System.DBNull.Value)
                        sb.Append("'").Append(Convert.ToDateTime(r["fec_act"]).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                    

                    if (totalReg == 200)
                    {

                        Console.CursorLeft = 0;
                        Console.Write(registrosLeidos.ToString("N0") + "/" + total.ToString("N0"));
                        firstOnly = true;
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        totalReg = 0;
                    }


                }
                db.CloseConnection();

                if (totalReg > 0)
                {
                    firstOnly = true;
                    Console.CursorLeft = 0;
                    Console.Write(registrosLeidos.ToString("N0") + "/" + total.ToString("N0"));
                    mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                    mcommand.ExecuteNonQuery();
                    mdb.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    totalReg = 0;
                }

                ss_pp.Update_Fecha_Fin("Contratación", "Copia ed_owner.t_ed_f_clis", "Copia ed_owner.t_ed_f_clis");

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("CopiaClientesCartera_RedShift: " + e.Message);
            }

        }

        public bool ImportacionSalesForce(string fichero, string pais)
        {
            System.IO.StreamReader fileStream;
            string line;
            string[] c;
            int numLinea = 0;
            bool hayError = false;
            int p = 0;
            int total_lineas = 0;
            double percent = 0;
            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            bool firstOnly = true;
            string cabecera = "";
            string strSql = "";

            servidores.MySQLDB db;
            MySqlCommand command;

            string estructura_fichero = "";


            FileInfo file = new FileInfo(fichero);

            try
            {

                pb.Show();


                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                while ((line = fileStream.ReadLine()) != null)
                {
                    if (firstOnly)
                    {
                        cabecera = line.Trim();
                        firstOnly = false;
                    }
                    total_lineas++;
                }
                fileStream.Close();

                if (pais == "ES")                
                    estructura_fichero = param.GetValue("estructura_fichero_ES");
                else
                    estructura_fichero = param.GetValue("estructura_fichero_PT");

                if (cabecera == estructura_fichero)
                {
                    if(pais == "ES")
                        strSql = "delete from salesforce_cartera_tmp";
                    else
                        strSql = "delete from salesforce_cartera_pt_tmp";

                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = total_lineas;
                    pb.Text = "Importanto " + file.Name;

                    fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));                    
                    while ((line = fileStream.ReadLine()) != null)
                    {
                        c = line.Split(';');
                        numLinea++;

                        percent = (numLinea / Convert.ToDouble(total_lineas)) * 100;
                        pb.txtDescripcion.Text = "Importanto: "
                            + string.Format("{0}", numLinea.ToString("#,##0"))
                            + " de "
                            + string.Format("{0}", total_lineas.ToString("#,##0"));
                        pb.progressBar.Value = numLinea;
                        pb.progressBar.Value = numLinea - 1;
                        pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                        pb.Refresh();



                        if (numLinea > 1 && c.Length > 3)
                        {
                            p = 0;

                            EndesaEntity.contratacion.Cartera s = new EndesaEntity.contratacion.Cartera();

                            if (pais == "ES")
                            {
                                if (c[p] != "")
                                    s.contrato = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;


                                // CUPS 22
                                if (c[p] != "")
                                    if (utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Length >= 20)
                                        s.cups22 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Substring(0, 20);
                                    else
                                        s.cups22 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                // CUPS 20
                                if (c[p] != "")
                                    if (utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Length >= 20)
                                        s.cups20 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Substring(0, 20);
                                    else
                                        s.cups20 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.nombre_cuenta = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.nif = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                                p++;

                                if (c[p] != "")
                                    s.gestor = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.email_gestor = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.posicion = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.subdireccion = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.subdirector = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.territorio = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.responsable_territorial = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.zona = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.responsable_zona = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.segmento = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.linea_negocio = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.estado_ps = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.estado_contrato = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                // Si estan informados gas y electricidad se duplican las lineas
                                if (s.cups22 != "" && s.cups20 != "")
                                {
                                    s.linea_negocio = "Electricidad";
                                    l_p.Add(s);
                                    s.linea_negocio = "Gas";
                                    s.cups22 = s.cups20;
                                    l_p.Add(s);
                                }
                                else
                                {
                                    if (s.cups22 == "" && s.cups20 != "")
                                        s.cups22 = s.cups20;

                                    l_p.Add(s);
                                }

                            }
                            else
                            {

                                if (c[p] != "")
                                    s.contrato = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;


                                // CUPS 22
                                if (c[p] != "")
                                    if (utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Length >= 20)
                                        s.cups22 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Substring(0, 20);
                                    else
                                        s.cups22 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                // CUPS 20
                                if (c[p] != "")
                                    if (utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Length >= 20)
                                        s.cups20 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]).Substring(0, 20);
                                    else
                                        s.cups20 = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.nombre_cuenta = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.nif = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                                p++;

                                if (c[p] != "")
                                    s.gestor = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.email_gestor = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.posicion = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.subdireccion = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.subdirector = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.responsable_territorial = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                                p++;

                                if (c[p] != "")
                                    s.zona = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                                p++;

                                if (c[p] != "")
                                    s.responsable_zona = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                                p++;

                                if (c[p] != "")
                                    s.tension = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                                p++;
                                if (c[p] != "")
                                    s.estado_contrato = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);


                                // Si estan informados gas y electricidad se duplican las lineas
                                if (s.cups20 != "")
                                {                                   
                                    
                                    s.linea_negocio = "Gas";
                                    s.cups22 = s.cups20;
                                    
                                }
                                else
                                    s.linea_negocio = "Electricidad";

                                l_p.Add(s);
                                

                            }
                                                            

                            if (l_p.Count() > 350)
                                PasaMemoria_a_MySQL_Temp(pais);
                        }
                    }


                    fileStream.Close();
                    fileStream = null;

                    PasaMemoria_a_MySQL_Temp(pais);

                    pb.Close();

                    if(total_lineas == numLinea)
                        MessageBox.Show("Importación completada."
                            + System.Environment.NewLine
                            + "Se han importado " + numLinea.ToString("#,##0")
                            + " de un total de " + total_lineas.ToString("#,##0")
                            , "Importacion",
                          MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                    {
                        MessageBox.Show("Importación NO completada!!!!!"
                            + System.Environment.NewLine
                            + "Se han importado " + numLinea.ToString("#,##0")
                            + " de un total de " + total_lineas.ToString("#,##0")
                            , "Importacion",
                          MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("La estructura el fichero no es la correcta.", "Importacion",
                      MessageBoxButtons.OK, MessageBoxIcon.Error);
                    pb.Close();
                    hayError = true;
                }

                return hayError;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Importacion",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
        }

        private void PasaMemoria_a_MySQL_Temp(string pais)
        {
            VuelcaMySQL(l_p, pais);
            l_p.Clear();
        }

        public void Pasa_datos_tablas_definitivas(string pais)
        {
            servidores.MySQLDB db;
            MySqlCommand command;            
            string strSql = "";

            if (pais == "ES")
            {
                strSql = "delete from salesforce_cartera;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into salesforce_cartera"
                    + " select * from salesforce_cartera_tmp";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into salesforce_cartera"
                   + " select * from salesforce_cartera_tmp where"
                   + " estado_contrato = 'EN VIGOR'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            else
            {
                strSql = "delete from salesforce_cartera_pt;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into salesforce_cartera_pt"
                    + " select * from salesforce_cartera_pt_tmp";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "replace into salesforce_cartera_pt"
                   + " select * from salesforce_cartera_pt_tmp where"
                   + " estado_contrato = 'EN VIGOR'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
        }


        private void VuelcaMySQL(List<EndesaEntity.contratacion.Cartera> lc, string pais)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int numReg = 0;
            

            

            int x = 0;
            try
            {

                for (int i = 0; i < lc.Count(); i++)
                {
                    numReg++;
                    x++;

                    if (firstOnly)
                    {
                        if(pais == "ES")
                        {
                            sb.Append("replace into salesforce_cartera_tmp (");
                            sb.Append("contrato, cups20, nombre_cuenta, nif, gestor, email_gestor, posicion,");
                            sb.Append("subdireccion, subdirector, territorio, responsable_territorial, zona,");
                            sb.Append("responsable_zona, segmento, linea_negocio, estado_ps, estado_contrato, f_ult_mod) values ");
                        }
                        else
                        {
                            sb.Append("replace into salesforce_cartera_pt_tmp (");
                            sb.Append("contrato, cups20, nombre_cuenta, nif, gestor, email_gestor, posicion,");
                            sb.Append("subdireccion, subdirector, responsable_territorial, zona,");
                            sb.Append("responsable_zona, tension, linea_negocio, estado_contrato, f_ult_mod) values ");
                        }
                                               
                        firstOnly = false;
                    }

                    if (pais == "ES")
                    {
                        sb.Append("('").Append(lc[i].contrato).Append("',");                        

                        if (lc[i].cups22 != "")
                            sb.Append("'").Append(lc[i].cups22).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append("'").Append(lc[i].nombre_cuenta).Append("',");
                        sb.Append("'").Append(lc[i].nif).Append("',");
                        sb.Append("'").Append(lc[i].gestor).Append("',");
                        sb.Append("'").Append(lc[i].email_gestor).Append("',");
                        sb.Append("'").Append(lc[i].posicion).Append("',");
                        sb.Append("'").Append(lc[i].subdireccion).Append("',");
                        sb.Append("'").Append(lc[i].subdirector).Append("',");
                        sb.Append("'").Append(lc[i].territorio).Append("',");
                        sb.Append("'").Append(lc[i].responsable_territorial).Append("',");
                        sb.Append("'").Append(lc[i].zona).Append("',");
                        sb.Append("'").Append(lc[i].responsable_zona).Append("',");
                        sb.Append("'").Append(lc[i].segmento).Append("',");
                        sb.Append("'").Append(lc[i].linea_negocio).Append("',");
                        sb.Append("'").Append(lc[i].estado_ps).Append("',");
                        sb.Append("'").Append(lc[i].estado_contrato).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                    }
                    else
                    {
                        sb.Append("('").Append(lc[i].contrato).Append("',");                        

                        if (lc[i].cups22 != "")
                            sb.Append("'").Append(lc[i].cups22).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append("'").Append(lc[i].nombre_cuenta).Append("',");
                        sb.Append("'").Append(lc[i].nif).Append("',");
                        sb.Append("'").Append(lc[i].gestor).Append("',");
                        sb.Append("'").Append(lc[i].email_gestor).Append("',");
                        sb.Append("'").Append(lc[i].posicion).Append("',");
                        sb.Append("'").Append(lc[i].subdireccion).Append("',");
                        sb.Append("'").Append(lc[i].subdirector).Append("',");                        
                        sb.Append("'").Append(lc[i].responsable_territorial).Append("',");
                        sb.Append("'").Append(lc[i].zona).Append("',");
                        sb.Append("'").Append(lc[i].responsable_zona).Append("',");                                     
                        sb.Append("'").Append(lc[i].tension).Append("',");
                        sb.Append("'").Append(lc[i].linea_negocio).Append("',");
                        sb.Append("'").Append(lc[i].estado_contrato).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                    }

                    

                    if (numReg == 350)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }

                }

                if (numReg > 0)
                {
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                }
                

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                      "VuelcaMySQL",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
                

                // ficheroLog.AddError("ImportacionTPLs.VuelcaMySQL " + e.Message);
            }
        }

        public void GetCartera(string nif)
        {
            EndesaEntity.contratacion.Cartera o;
            if (dic.TryGetValue(nif, out o))
            {
                
                this.cups22 = o.cups22;
                this.nombre_cuenta = o.nombre_cuenta;
                this.nif = o.nif;
                this.gestor = o.gestor;
                this.email_gestor = o.email_gestor;
                this.posicion = o.posicion;
                this.subdireccion = o.subdireccion;
                this.subdirector = o.subdirector;
                this.territorio = o.territorio;
                this.responsable_territorial = o.responsable_territorial;
                this.zona = o.zona;
                this.responsable_zona = o.responsable_zona;
                this.segmento = o.segmento;
                this.email_responsable_rt = o.email_responsable_rt;


            }
            else
            {
                
                this.cups22 = "";
                this.nombre_cuenta = "";
                this.nif = "";
                this.gestor = "";
                this.email_gestor = "";
                this.posicion = "";
                this.subdireccion = "";
                this.subdirector = "";
                this.territorio = "";
                this.responsable_territorial = "";
                this.zona = "";
                this.responsable_zona = "";
                this.segmento = "";
                this.email_responsable_rt = "";
            }

        }


        public void GetCarpera_CUPS20(string cups20)
        {
            EndesaEntity.contratacion.Cartera o;
            foreach(KeyValuePair<string, EndesaEntity.contratacion.Cartera> p in dic)
            {
                if(p.Value.cups20 == cups20)
                {
                    this.cups22 = p.Value.cups22;
                    this.cups20 = p.Value.cups20;   
                    this.nombre_cuenta = p.Value.nombre_cuenta;
                    this.nif = p.Value.nif;
                    this.gestor = p.Value.gestor;
                    this.email_gestor = p.Value.email_gestor;
                    this.posicion = p.Value.posicion;
                    this.subdireccion = p.Value.subdireccion;
                    this.subdirector = p.Value.subdirector;
                    this.territorio = p.Value.territorio;
                    this.responsable_territorial = p.Value.responsable_territorial;
                    this.zona = p.Value.zona;
                    this.responsable_zona = p.Value.responsable_zona;
                    this.segmento = p.Value.segmento;
                    this.email_responsable_rt = p.Value.email_responsable_rt;
                }

            }


            
        }

        public bool ExisteCartera(string nif)
        {
            bool existe = false;
            EndesaEntity.contratacion.Cartera o;
            if (dic.TryGetValue(nif, out o))
            {
                GetCartera(nif);
                existe = true;
            }

            return existe;
        }




        public string Direccion(string nif)
        {
            EndesaEntity.contratacion.Cartera o;
            if (dic.TryGetValue(nif, out o))
                return o.zona;
            else
                return "";


        }

        public List<string> ListaMail_Responsables_RT()
        {
            // Para el envío de altas de EXXI se envía 
            // en el Para la lista de Responsables
            // territoriales.

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            List<string> listaMail = new List<string>();

            try
            {
                strSql = "select email_gestor from salesforce_cartera_rt";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["email_gestor"] != System.DBNull.Value)
                        listaMail.Add(r["email_gestor"].ToString());
                }
                db.CloseConnection();
                return listaMail;
            }
            catch(Exception ex)
            {

                MessageBox.Show(ex.Message,
                "ListaMail_Responsables_RT",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return null;
            }    
            

        }

        

    }
}
