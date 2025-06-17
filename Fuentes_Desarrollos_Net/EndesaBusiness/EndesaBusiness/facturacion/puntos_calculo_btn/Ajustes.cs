using EndesaBusiness.global;
using EndesaBusiness.servidores;
using EndesaEntity.global;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Org.BouncyCastle.Bcpg;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySQLDB = EndesaBusiness.servidores.MySQLDB;

namespace EndesaBusiness.facturacion.puntos_calculo_btn
{
    public class Ajustes
    {

        public bool hayError { get; set; }
        public string descripcion_error { get; set; }
        List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste> lista_ajustes { get; set; }
        public Dictionary<string, List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste>> dic { get; set; }

        public Ajustes()
        {
            lista_ajustes = new List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste>();
            
        }

        public void Carga()
        {
            dic = Carga_Tabla_Ajustes();
        }

        private Dictionary<string, List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste>> Carga_Tabla_Ajustes()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;


            Dictionary<string, List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste>> d =
                new Dictionary<string, List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste>>();
            try
            {
                strSql = "SELECT cpe, nif, finiajus, ffinajus, finisajus, fcreajus, fbajaus,"
                    +  " testajus, parametro1, parametro2, parametro3, parametro4,"
                    + " created_by, created_date, last_update_by, last_update_date"
                    + " FROM fact.lpc_btn_ajustes";
                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.puntos_calculo_btn.Ajuste c = 
                        new EndesaEntity.facturacion.puntos_calculo_btn.Ajuste();

                    c.cupsree = r["cpe"].ToString();
                    c.nif = r["nif"].ToString();

                    if (r["finiajus"] != System.DBNull.Value)
                        c.finiajus = Convert.ToDateTime(r["finiajus"]);

                    if (r["ffinajus"] != System.DBNull.Value)
                        c.ffinajus = Convert.ToDateTime(r["ffinajus"]);

                    if (r["finisajus"] != System.DBNull.Value)
                        c.finisajus = Convert.ToDateTime(r["finisajus"]);

                    if (r["fcreajus"] != System.DBNull.Value)
                        c.fcreajus = Convert.ToDateTime(r["fcreajus"]);

                    if (r["fbajaus"] != System.DBNull.Value)
                        c.fbajajus = Convert.ToDateTime(r["fbajaus"]);

                    if (r["parametro3"] != System.DBNull.Value)
                        c.parametros[2] = r["parametro3"].ToString();

                    List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste> o;
                    if (!d.TryGetValue(c.cupsree, out o))
                    {
                        o = new List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste>();
                        o.Add(c);
                        d.Add(c.cupsree, o);
                    }
                    else
                        o.Add(c);

                        
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public double GetPCT_Ajuste(string cpe)
        {
            double resultado = 0;
            List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste> o;
            if (dic.TryGetValue(cpe, out o))
            {
                EndesaEntity.facturacion.puntos_calculo_btn.Ajuste c =
                    o.Find(z => z.finiajus <= DateTime.Now && z.ffinajus >= DateTime.Now);

                if (c != null)
                    resultado =  Convert.ToDouble(c.parametros[2]);
                
            }

            return resultado;
            

        }

        public void CargaExcel(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";
            string fecha_texto = "";


            try
            {
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();


                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 1000000; i++)
                {
                    c = 6;

                    if (firstOnly)
                    {
                        for (int j = 6; j <= 19; j++)
                            cabecera = cabecera + workSheet.Cells[1, j].Value.ToString();


                    }
                    if (!EstructuraCorrecta(cabecera))
                    {
                        this.hayError = true;
                        this.descripcion_error = "La estructura del archivo excel no es la correcta.";
                        break;
                    }

                    firstOnly = false;

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;

                    if (workSheet.Cells[f, 1].Value.ToString()
                        + workSheet.Cells[f, 2].Value.ToString()
                        + workSheet.Cells[f, 3].Value.ToString() == "")
                    {
                        break;
                    }
                    else
                    {
                        EndesaEntity.facturacion.puntos_calculo_btn.Ajuste ajus =
                            new EndesaEntity.facturacion.puntos_calculo_btn.Ajuste();

                        ajus.cupsree = workSheet.Cells[f, c].Value.ToString().Trim(); c++;
                        ajus.nif = workSheet.Cells[f, c].Value.ToString().Trim(); c++;

                        fecha_texto = workSheet.Cells[f, c].Value.ToString(); c++;
                        if (Convert.ToInt32(fecha_texto) > 0)
                        {
                            ajus.finiajus = new DateTime(Convert.ToInt32(fecha_texto.Substring(0, 4)),
                            Convert.ToInt32(fecha_texto.Substring(4, 2)),
                            Convert.ToInt32(fecha_texto.Substring(6, 2)));
                        }


                        fecha_texto = workSheet.Cells[f, c].Value.ToString(); c++;
                        if (Convert.ToInt32(fecha_texto) > 0)
                        {
                            ajus.ffinajus = new DateTime(Convert.ToInt32(fecha_texto.Substring(0, 4)),
                            Convert.ToInt32(fecha_texto.Substring(4, 2)),
                            Convert.ToInt32(fecha_texto.Substring(6, 2)));
                        }


                        fecha_texto = workSheet.Cells[f, c].Value.ToString(); c++;
                        if (Convert.ToInt32(fecha_texto) > 0)
                        {
                            ajus.finisajus = new DateTime(Convert.ToInt32(fecha_texto.Substring(0, 4)),
                            Convert.ToInt32(fecha_texto.Substring(4, 2)),
                            Convert.ToInt32(fecha_texto.Substring(6, 2)));
                        }

                        fecha_texto = workSheet.Cells[f, c].Value.ToString(); c++;
                        if (Convert.ToInt32(fecha_texto) > 0)
                        {
                            ajus.fcreajus = new DateTime(Convert.ToInt32(fecha_texto.Substring(0, 4)),
                            Convert.ToInt32(fecha_texto.Substring(4, 2)),
                            Convert.ToInt32(fecha_texto.Substring(6, 2)));
                        }

                        fecha_texto = workSheet.Cells[f, c].Value.ToString(); c++;
                        if (Convert.ToInt32(fecha_texto) > 0)
                        {
                            ajus.fbajajus = new DateTime(Convert.ToInt32(fecha_texto.Substring(0, 4)),
                            Convert.ToInt32(fecha_texto.Substring(4, 2)),
                            Convert.ToInt32(fecha_texto.Substring(6, 2)));
                        }

                        ajus.testajus = Convert.ToInt32(workSheet.Cells[f, c].Value.ToString()); c++;

                        for (int j = 0; j < 4; j++)
                        {
                            ajus.parametros[j] = workSheet.Cells[f, c].Value.ToString().Trim();
                            c++;
                        }

                        lista_ajustes.Add(ajus);

                    }



                }
                fs = null;
                excelPackage = null;


                GuardaDatos_Ajustes(lista_ajustes);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                        "Error en el formato del fichero",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
            }
        }

        private void GuardaDatos_Ajustes(List<EndesaEntity.facturacion.puntos_calculo_btn.Ajuste> lista)
        {

            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            MySQLDB db;
            MySqlCommand command;
            int x = 0;          
            

            try
            {
                if (lista != null)
                    foreach (EndesaEntity.facturacion.puntos_calculo_btn.Ajuste p in lista)
                    {
                        x++;

                        if (firstOnly)
                        {
                            sb.Append("replace into lpc_btn_ajustes (cpe, nif, finiajus,");
                            sb.Append(" ffinajus,finisajus, fcreajus, fbajaus, testajus,");
                            sb.Append(" parametro1, parametro2, parametro3, parametro4,");
                            sb.Append(" created_by, created_date) values ");
                            firstOnly = false;
                        }

                        sb.Append("('").Append(p.cupsree).Append("',");
                        sb.Append("'").Append(p.nif).Append("',");

                        if (p.finiajus > DateTime.MinValue)
                            sb.Append("'").Append(p.finiajus.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.ffinajus > DateTime.MinValue)
                            sb.Append("'").Append(p.ffinajus.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.finisajus > DateTime.MinValue)
                            sb.Append("'").Append(p.finisajus.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.fcreajus > DateTime.MinValue)
                            sb.Append("'").Append(p.fcreajus.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        if (p.fbajajus > DateTime.MinValue)
                            sb.Append("'").Append(p.fbajajus.ToString("yyyy-MM-dd")).Append("',");
                        else
                            sb.Append("null,");

                        sb.Append(p.testajus).Append(",");

                        for (int j = 0; j < 4; j++)
                            sb.Append("'").Append(p.parametros[j]).Append("',");

                        sb.Append("'").Append(System.Environment.UserName).Append("',");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");


                        if (x == 500)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
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
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                        "Error en el guardado de datos",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }
        }

        private bool EstructuraCorrecta(string cabecera)
        {


            if (cabecera.ToUpper().Trim() == "CUPSREENIFFINIAJUSFFINAJUSFINISAJUSFCREAJUSFBAJAJUSTESTAJUSPARAMETRO1PARAMETRO2PARAMETRO3PARAMETRO4TAPLHEREUSER")
                return true;
            else
                return false;
        }
    }
}
