using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class Spot : EndesaEntity.facturacion.PreciosSpot
    {
        public bool hayDatos { get; set; }
        public int totalPeriodosCuartoHorarios { get; set; }
        public List<EndesaEntity.facturacion.PreciosSpot> lista { get; set; }

        public double precio_medio { get; set; }
        public Spot()
        {

        }

        public Spot(DateTime fd, DateTime fh)
        {
            utilidades.Fechas utilFechas = new utilidades.Fechas();
            totalPeriodosCuartoHorarios = utilFechas.NumHoras(fd, fh) * 4;
            lista = Carga(fd, fh);
            hayDatos = lista.Count() == totalPeriodosCuartoHorarios;

        }
        private List<EndesaEntity.facturacion.PreciosSpot> Carga(DateTime fd, DateTime fh)
        {
            List<EndesaEntity.facturacion.PreciosSpot> l = new List<EndesaEntity.facturacion.PreciosSpot>();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select fecha, precio"
                    + " FROM fact.ag_pt_spot where"
                    + " (fecha >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha < '" + fh.AddDays(1).ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.PreciosSpot c = new EndesaEntity.facturacion.PreciosSpot();
                    c.fecha_hora = Convert.ToDateTime(r["fecha"]);
                    c.precio = Convert.ToDouble(r["precio"]);
                    precio_medio = precio_medio + c.precio;
                    l.Add(c);
                }
                precio_medio = (precio_medio / l.Count);

                db.CloseConnection();
                return l;

            }
            catch (Exception e)
            {

                //MessageBox.Show("Carga completada satisfactoriamente.",
                //  "Clicks.Carga",
                //  MessageBoxButtons.OK,
                //  MessageBoxIcon.Error);
                Console.WriteLine(e.Message);
                return null;
            }


        }
        public bool ImportarPreciosSpot(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            int linea = 0;
            bool firstOnlyCabecera = true;
            bool firstOnlyFecha = true;
            bool firstOnly = true;
            string cabecera = "";
            DateTime fecha = new DateTime();
            string fecha_texto = "";
            List<EndesaEntity.facturacion.PreciosSpot> lista = new List<EndesaEntity.facturacion.PreciosSpot>();
            bool hayError = false;
            DateTime minfd = new DateTime();
            DateTime maxfh = new DateTime();
            int totalRegistros = 0;

            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 10000; i++)
                {
                    linea = 1;
                    c = 1;

                    if (firstOnlyCabecera)
                    {
                        for (int w = 1; w < 5; w++)
                            cabecera += workSheet.Cells[1, w].Value.ToString();

                        if (!EstructuraCorrectaSpot(cabecera))
                        {


                            MessageBox.Show("La estructura del archivo no es correcta y no se importará ningún valor.",
                                "Estructura Incorrecta!!!",
                                MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                            break;
                        }


                        firstOnlyCabecera = false;
                    }

                    f++;

                    cabecera = "";
                    if (workSheet.Cells[f, 5].Value != null)
                    {
                        totalRegistros++;
                        if (firstOnlyFecha)
                        {
                            fecha_texto = workSheet.Cells[f, 3].Value.ToString();
                            fecha = new DateTime(Convert.ToInt32(fecha_texto.Substring(0, 4)),
                                                 Convert.ToInt32(fecha_texto.Substring(5, 2)),
                                                 Convert.ToInt32(fecha_texto.Substring(8, 2)));
                            // fecha = fecha.AddHours(Convert.ToInt32(fecha_texto.Substring(11, 2)));
                            minfd = fecha;
                            firstOnlyFecha = false;
                        }
                        else
                        {
                            fecha = fecha.AddMinutes(15);
                            maxfh = fecha > maxfh ? fecha : maxfh;
                        }


                        EndesaEntity.facturacion.PreciosSpot s = new EndesaEntity.facturacion.PreciosSpot();
                        s.fecha_hora = fecha;
                        s.precio = Convert.ToDouble(workSheet.Cells[f, 5].Value.ToString());
                        lista.Add(s);

                    }

                }

                fs = null;
                excelPackage = null;

                id = 0;
                for (int i = 0; i < lista.Count; i++)
                {
                    id++;
                    if (firstOnly)
                    {
                        sb.Append("replace into ag_pt_spot (fecha, precio, usuario) values ");
                        firstOnly = false;
                    }
                    sb.Append("('").Append(lista[i].fecha_hora.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    sb.Append(lista[i].precio.ToString().Replace(",", ".")).Append(",");
                    sb.Append("'").Append(System.Environment.UserName).Append("'),");



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

                Cursor.Current = Cursors.Default;
                MessageBox.Show("Carga finalizada correctamente."
                         + System.Environment.NewLine
                         + "Se han procesado " + String.Format("{0:#,##0}", totalRegistros) + " registros"
                         + System.Environment.NewLine
                         + "desde las fechas " + minfd.ToString("dd/MM/yyyy") + " hasta " + maxfh.ToString("dd/MM/yyyy"),
                             "Carga precios spot finalizada correctamente",
                             MessageBoxButtons.OK,
                         MessageBoxIcon.Information);

                return false;

            }
            catch (Exception e)
            {
                MessageBox.Show("Error en la línea " + linea + " --> " + e.Message,
                  "Error en la importación de Precios Spot",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
                return true;
            }
        }
        private bool EstructuraCorrectaSpot(string linea)
        {
            return linea.Equals("geonamevaluedatetimedia");
        }
    }
}
