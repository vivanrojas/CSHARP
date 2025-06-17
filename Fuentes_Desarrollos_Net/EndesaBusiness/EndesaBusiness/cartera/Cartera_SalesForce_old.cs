using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.cartera
{
    public class Cartera_SalesForce_old : EndesaEntity.cobros.CarteraSIOC_Tabla
    {
        List<EndesaEntity.cobros.Cartera_SalesForce> l_p;
        Dictionary<string, EndesaEntity.global.Jerarquia_CarteraSalesForce> dic_jerarquia;
        Dictionary<string, EndesaEntity.cobros.CarteraSIOC_Tabla> dic;
        public Cartera_SalesForce_old()
        {
            l_p = new List<EndesaEntity.cobros.Cartera_SalesForce>();
            dic = new Dictionary<string, EndesaEntity.cobros.CarteraSIOC_Tabla>();
            dic_jerarquia = CargaJerarquia();
        }

        public Cartera_SalesForce_old(List<string> lista_nifs)
        {
            l_p = new List<EndesaEntity.cobros.Cartera_SalesForce>();
            //dic_jerarquia = CargaJerarquia();
            dic = new Dictionary<string, EndesaEntity.cobros.CarteraSIOC_Tabla>();
            Carga(lista_nifs);
            
        }


        private Dictionary<string, EndesaEntity.global.Jerarquia_CarteraSalesForce> CargaJerarquia()
        {
            string strSql;
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.global.Jerarquia_CarteraSalesForce> d =
                new Dictionary<string, EndesaEntity.global.Jerarquia_CarteraSalesForce>();

            try
            {
                strSql = "select segmento, subdireccion, territorio, zona, nueva_posicion,"
                    + " gestor, responsable_zona, responsable_territorial, subdirector"
                    + " from salesforce_jerarquia";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.global.Jerarquia_CarteraSalesForce c =
                        new EndesaEntity.global.Jerarquia_CarteraSalesForce();
                    
                    if (r["gestor"] != System.DBNull.Value)
                        c.gestor = r["gestor"].ToString();

                    if (r["responsable_zona"] != System.DBNull.Value)
                        c.responsable_zona = r["responsable_zona"].ToString();

                    if (r["responsable_territorial"] != System.DBNull.Value)
                        c.responsable_territorial = r["responsable_territorial"].ToString();

                    if (r["subdirector"] != System.DBNull.Value)
                        c.subdirector = r["subdirector"].ToString();

                    if (r["territorio"] != System.DBNull.Value)
                        c.territorio = r["territorio"].ToString();

                    EndesaEntity.global.Jerarquia_CarteraSalesForce o;
                    if (!d.TryGetValue(c.gestor, out o))
                        d.Add(c.gestor, c);


                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        private void Carga(List<string> lista_nifs)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;

            try
            {

                strSql = "SELECT c.identificador AS cif, c.segmento," 
                    + " c.propietario_cuenta_nombre_completo AS nombreGestor,"
                    + " c.segmento AS direccion,"
                    + " j.nueva_posicion AS descResponsableTerritorial,"
                    + " j.responsable_territorial AS responsableTerritorial,"
                    + " j.responsable_zona AS responsableZona"
                    + " FROM salesforce_cartera c"
                    + " INNER JOIN salesforce_jerarquia j ON"
                    + " j.gestor = c.propietario_cuenta_nombre_completo where c.identificador in (";

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

                strSql += ");";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cobros.CarteraSIOC_Tabla c = new EndesaEntity.cobros.CarteraSIOC_Tabla();
                    c.cif = r["cif"].ToString();

                    if (r["nombreGestor"] != System.DBNull.Value)
                        c.nombreGestor = r["nombreGestor"].ToString();

                    if (r["segmento"] != System.DBNull.Value)
                        c.segmento = r["segmento"].ToString();


                    if (r["direccion"] != System.DBNull.Value)
                        c.direccion = r["direccion"].ToString();

                    if (r["descResponsableTerritorial"] != System.DBNull.Value)
                        c.descResponsableTerritorial = r["descResponsableTerritorial"].ToString();

                    if (r["responsableTerritorial"] != System.DBNull.Value)
                        c.responsableTerritorial = r["responsableTerritorial"].ToString();

                    if (r["responsableZona"] != System.DBNull.Value)
                        c.responsableZona = r["responsableZona"].ToString();

                    EndesaEntity.cobros.CarteraSIOC_Tabla o;
                    if(!dic.TryGetValue(c.cif, out o))
                        dic.Add(c.cif, c);

                }
                db.CloseConnection();



            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Cartera - Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }


        public bool ImportacionSalesForce(string fichero)
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

            FileInfo file = new FileInfo(fichero);
            

            

            try
            {
                
                pb.Show();
                

                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                while ((line = fileStream.ReadLine()) != null)
                {
                    total_lineas++;
                }
                fileStream.Close();

                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = total_lineas;
                pb.Text = "Importanto " + file.Name;

                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                fileStream = new System.IO.StreamReader(file.FullName);
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
                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                    pb.Refresh();

                    if (numLinea > 1 && c.Length > 3)
                    {
                        p = 0;

                        EndesaEntity.cobros.Cartera_SalesForce s = new EndesaEntity.cobros.Cartera_SalesForce();

                        if (c[p] != "")
                            s.nombre_cuenta = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.tipo_identificador = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.identificador = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;                        

                        if (c[p] != "")
                            s.segmento = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.cuenta_principal = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                        p++;

                        if (c[p] != "")
                            s.agrupacion = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.propietario_cuenta_nombre_completo = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.posicion_cuenta = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.responsable_zona = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.responsable_territorio = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.subdirector = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.subdireccion = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.territorio = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);

                        p++;

                        if (c[p] != "")
                            s.zona = utilidades.FuncionesTexto.ArreglaAcentos(c[p]);
                        

                        l_p.Add(s);

                        if (l_p.Count() > 350)
                            PasaMemoria_a_MySQL_Temp();
                    }
                }

                
                fileStream.Close();
                fileStream = null;
                PasaMemoria_a_MySQL_Temp();
                pb.Close();

                MessageBox.Show("Importación completada", "Importacion",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);

                return hayError;
            }
            catch (Exception e)
            {
               MessageBox.Show(e.Message,"Importacion",
                   MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
        }

        private void PasaMemoria_a_MySQL_Temp()
        {
            VuelcaMySQL(l_p);
            l_p.Clear();

        }

        private void VuelcaMySQL(List<EndesaEntity.cobros.Cartera_SalesForce> lc)
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
                        sb.Append("replace into salesforce_cartera_nueva (nombre_cuenta, tipo_identificador, identificador,");
                        sb.Append("segmento, cuenta_principal, agrupacion, propietario_cuenta_nombre_completo,");
                        sb.Append("posicion_cuenta, responsable_zona, responsable_territorio, subdirector, subdireccion,");
                        sb.Append("territorio, zona) values ");

                        //sb.Append("replace into salesforce_cartera (nombre_cuenta, Identificador, tipo_identificador,");
                        //sb.Append("segmento, apellidos_razonsocial, grupo_empresarial_nombre_cuenta, clasificacion_abc,");
                        //sb.Append("segmento4, segmento5, segmento1, posicion_cuenta, propietario_cuenta_nombre_completo,");
                        //sb.Append("propietario_cuenta_id, posicion_gestor, codigo_postal_envio, estado, cliente_principal,");
                        //sb.Append("tipo_registro_cuenta) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(lc[i].nombre_cuenta).Append("',");
                    sb.Append("'").Append(lc[i].tipo_identificador).Append("',");
                    sb.Append("'").Append(lc[i].identificador).Append("',");                    
                    sb.Append("'").Append(lc[i].segmento).Append("',");
                    sb.Append("'").Append(lc[i].cuenta_principal).Append("',");
                    sb.Append("'").Append(lc[i].agrupacion).Append("',");
                    sb.Append("'").Append(lc[i].propietario_cuenta_nombre_completo).Append("',");
                    sb.Append("'").Append(lc[i].posicion_cuenta).Append("',");
                    sb.Append("'").Append(lc[i].responsable_zona).Append("',");
                    sb.Append("'").Append(lc[i].responsable_territorio).Append("',");                    
                    sb.Append("'").Append(lc[i].subdirector).Append("',");
                    sb.Append("'").Append(lc[i].subdireccion).Append("',");
                    sb.Append("'").Append(lc[i].territorio).Append("',");
                    sb.Append("'").Append(lc[i].zona).Append("'),");
                   
                    

                    if (numReg == 500)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.AUX);
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
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                }




            }
            catch (Exception e)
            {
                // ficheroLog.AddError("ImportacionTPLs.VuelcaMySQL " + e.Message);
            }
        }

        
        public EndesaEntity.global.Jerarquia_CarteraSalesForce GetDatosJerarquia(string gestor)
        {
            EndesaEntity.global.Jerarquia_CarteraSalesForce o;
            if (dic_jerarquia.TryGetValue(gestor, out o))
                return o;
            else
                return null;
        }


        public void GetCartera(string nif)
        {
            EndesaEntity.cobros.CarteraSIOC_Tabla o;
            if (dic.TryGetValue(nif, out o))
            {
                this.apellido1Gestor = o.apellido1Gestor;
                this.apellido2Gestor = o.apellido2Gestor;
                this.nombreGestor = o.nombreGestor;
                this.responsableTerritorial = o.responsableTerritorial;
                this.responsableZona = o.responsableZona;
                this.direccion = o.direccion;
                this.descResponsableTerritorial = o.descResponsableTerritorial;
                this.segmento = o.segmento;
            }
            else
            {
                this.apellido1Gestor = "";
                this.apellido2Gestor = "";
                this.nombreGestor = "";
                this.responsableTerritorial = "";
                this.responsableZona = "";
                this.direccion = "";
                this.descResponsableTerritorial = "";
                this.segmento = "";
            }

        }

        public bool ExisteCartera(string nif)
        {
            bool existe = false;
            EndesaEntity.cobros.CarteraSIOC_Tabla o;
            if (dic.TryGetValue(nif, out o))
            {
                GetCartera(nif);
                existe = true;
            }


            return existe;
        }

        public string Direccion(string nif)
        {
            EndesaEntity.cobros.CarteraSIOC_Tabla o;
            if (dic.TryGetValue(nif, out o))
                return o.direccion;
            else
                return "";


        }

    }

}
