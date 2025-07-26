using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.medida
{
    public partial class FrmExportacion_PMML : Form
    {
        EndesaBusiness.utilidades.Param p;

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmExportacion_PMML()
        {
            usage.Start("Medida", "FrmExportacion_PMML", "N/A");
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_exportar_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string archivo = ExportaTablaAccess();

            if (archivo != "" && (p.GetValue("subir_a_FTP") == "S"))
                Subir_FTP(archivo);
            Cursor.Current = Cursors.Default;


        }

        private void FrmExportacion_PMML_Load(object sender, EventArgs e)
        {
            p = new EndesaBusiness.utilidades.Param("scea_pmml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.MED);
            lbl_ultima_exportacion.Text = "Última Exportación: " + p.GetValue("ultima_exportacion");
            lbl_ruta_access.Text = "Ruta Access: " + p.GetValue("access_ruta");
        }

        private string ExportaTablaAccess()
        {
            string archivo = "";
            string strSql = "";
            EndesaBusiness.servidores.AccessDB ac;
            OleDbDataReader r;
            StringBuilder sb = new StringBuilder();
            string linea = "";
            bool firstOnly = true;

            try
            {
                
                string[] listaArchivos = Directory.GetFiles(p.GetValue("salida"), p.GetValue("prefijo_archivo") + "*.csv");
                for (int i = 0; i < listaArchivos.Length; i++)
                {
                    FileInfo file = new FileInfo(listaArchivos[i]);
                    file.Delete();
                }
                    


                archivo = p.GetValue("salida") +
                    p.GetValue("prefijo_archivo") +
                    "_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                StreamWriter swa = new StreamWriter(archivo, false);

                //strSql = "SELECT "
                //    + " [PM ML].[PM],"
                //    + " [PM ML].[CUPS],"
                //    + " [PM ML].[TIPOPM],"
                //    + " [PM ML].[ESTADOPM],"
                //    + " [PM ML].[VERSION],"
                //    + " [PM ML].[MODO_LECT],"
                //    + " [PM ML].[ESTATEL],"
                //    + " [PM ML].[TLF ML],"
                //    + " [PM ML].[TLF AUX],"
                //    + " [PM ML].[DE ML],"
                //    + " [PM ML].[PM ML],"
                //    + " [PM ML].[CL ML],"
                //    + " [PM ML].[FECHA_ALTA],"
                //    + " [PM ML].[FECHA_BAJA],"
                //    + " [PM ML].[FECHA_PRU_TEL],"
                //    + " [PM ML].[COD_REE],"
                //    + " [PM ML].[FUNCION_PM],"
                //    + " [PM ML].[Contrato],"
                //    + " [PM ML].[PF],"
                //    + " [PM ML].[CUPS22],"
                //    + " [PM ML].[cIns],"
                //    + " [PM ML].[PMPrincipal],"
                //    + " [PM ML].[fIns],"
                //    + " [PM ML].[fDesIns],"
                //    + " [PM ML].[fCalCon],"
                //    + " [PM ML].[fVigor],"
                //    + " [PM ML].[fHtco],"
                //    + " [PM ML].[tablaOrigen],"
                //    + " [PM ML].[nose],"
                //    + " [PM ML].[nPMs],"
                //    + " [PM ML].[nPMs P],"
                //    + " [PM ML].[IP ML],"
                //    + " [PM ML].[PUERTO ML]"
                //    + " FROM [PM ML]";

                strSql = "SELECT[PM ML].CUPS AS CUPS13, [PM ML].PF, [PM ML].PM AS CUPS15PM," +
                    " IIf([CUPS22]= 'ES','',[CUPS22]) AS CUPS22PM, SCEA.[Descripcion Estado SCE], " +
                    "SCEA.tdistri, [PM ML].Contrato, [PM ML].nPMs, [PM ML].[nPMs P], [PM ML].TIPOPM, " +
                    "[PM ML].VERSION, [PM ML].[TLF ML], [PM ML].[DE ML], [PM ML].[PM ML], Val([CL ML]) AS CL_ML, " +
                    "[PM ML].[TLF AUX], [PM ML].[IP ML], [PM ML].[PUERTO ML], [PM ML].MODO_LECT" +
                    " FROM[PM ML] LEFT JOIN SCEA ON[PM ML].CUPS = SCEA.IDU";

                strSql = strSql.Replace("'", "\"");

                ac = new EndesaBusiness.servidores.AccessDB(p.GetValue("access_ruta"));
                OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {
                    if (firstOnly)
                    {
                        linea = "CUPS13;PF;CUPS15PM;CUPS22PM;Descripcion Estado SCE;tdistri;Contrato;nPMs;nPMs P;" +
                            "TIPOPM;VERSION;TLF ML;DE ML;PM ML;CL_ML;TLF AUX;IP ML;PUERTO ML;MODO_LECT";
                        firstOnly = false;
                    }else
                    {
                        if (r["CUPS13"] != null)
                            linea = r["CUPS13"].ToString();
                        linea = linea + ";";

                        if (r["PF"] != null)
                            linea = linea + r["PF"].ToString();
                        linea = linea + ";";

                        if (r["CUPS15PM"] != null)
                            linea = linea + r["CUPS15PM"].ToString();
                        linea = linea + ";";

                        if (r["CUPS22PM"] != null)
                            linea = linea + r["CUPS22PM"].ToString();
                        linea = linea + ";";

                        if (r["Descripcion Estado SCE"] != null)
                            linea = linea + r["Descripcion Estado SCE"].ToString();
                        linea = linea + ";";

                        if (r["tdistri"] != null)
                            linea = linea + r["tdistri"].ToString();
                        linea = linea + ";";

                        if (r["Contrato"] != null)
                            linea = linea + r["Contrato"].ToString();
                        linea = linea + ";";

                        if (r["nPMs"] != null)
                            linea = linea + r["nPMs"].ToString();
                        linea = linea + ";";

                        if (r["nPMs P"] != null)
                            linea = linea + r["nPMs P"].ToString();
                        linea = linea + ";";

                        if (r["TIPOPM"] != null)
                            linea = linea + r["TIPOPM"].ToString();
                        linea = linea + ";";

                        if (r["VERSION"] != null)
                            linea = linea + r["VERSION"].ToString();
                        linea = linea + ";";

                        if (r["TLF ML"] != null)
                            linea = linea + r["TLF ML"].ToString();
                        linea = linea + ";";

                        if (r["DE ML"] != null)
                            linea = linea + r["DE ML"].ToString();
                        linea = linea + ";";

                        if (r["PM ML"] != null)
                            linea = linea + r["PM ML"].ToString();
                        linea = linea + ";";

                        if (r["CL_ML"] != null)
                            linea = linea + r["CL_ML"].ToString();
                        linea = linea + ";";

                        if (r["TLF AUX"] != null)
                            linea = linea + r["TLF AUX"].ToString();
                        linea = linea + ";";

                        if (r["IP ML"] != null)
                            linea = linea + r["IP ML"].ToString();
                        linea = linea + ";";

                        if (r["PUERTO ML"] != null)
                            linea = linea + r["PUERTO ML"].ToString();
                        linea = linea + ";";

                        if (r["MODO_LECT"] != null)
                            linea = linea + r["MODO_LECT"].ToString();
                    }
                    


                    swa.WriteLine(linea);
                }
                ac.CloseConnection();
                swa.Close();

                return archivo;
            }catch(Exception e)
            {

                MessageBox.Show(e.Message,
              "Exportación tabla PM ML a FTP",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);

                return "";

            }


        }

        private void Subir_FTP(string archivo)
        {

            EndesaBusiness.utilidades.UltimateFTP ftp;
            FileInfo file;
            EndesaBusiness.medida.PM_ML pmml;

            try
            {

                pmml = new EndesaBusiness.medida.PM_ML();

                file = new FileInfo(archivo);

                ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                        EndesaBusiness.utilidades.FuncionesTexto.Decrypt(p.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                        p.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                ftp.Upload(p.GetValue("ftp_ruta_salida") + file.Name, archivo);
                pmml.UpdateFechaCopia();

                p = new EndesaBusiness.utilidades.Param("scea_pmml_param", EndesaBusiness.servidores.MySQLDB.Esquemas.MED);
                lbl_ultima_exportacion.Text = "Última Exportación: " + p.GetValue("ultima_exportacion");


                MessageBox.Show("Se ha exportado el archivo "
                   + archivo + " correctamente.",
              "Exportación tabla PM ML a FTP",
              MessageBoxButtons.OK,
              MessageBoxIcon.Information);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
             "Exportación tabla PM ML a FTP",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }

        private void parámetrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización PM ML";
            p.tabla = "scea_pmml_param";
            p.esquemaString = "MED";
            p.Show();
        }

        private void btn_macro_Click(object sender, EventArgs e)
        {

        }

        private void FrmExportacion_PMML_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Medida", "FrmExportacion_PMML", "N/A");
        }
    }
}
