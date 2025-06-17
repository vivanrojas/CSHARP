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
    public class Clicks: EndesaEntity.facturacion.ClicksPT
    {
        public Dictionary<string, List<EndesaEntity.facturacion.ClicksPT>> dic { get; set; }

        public Clicks(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);

        }

        private Dictionary<string, List<EndesaEntity.facturacion.ClicksPT>> Carga(DateTime fd, DateTime fh)
        {
            Dictionary<string, List<EndesaEntity.facturacion.ClicksPT>> d = new Dictionary<string, List<EndesaEntity.facturacion.ClicksPT>>();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "SELECT i.nif, i.cliente, c.cpe, c.click, c.fecha, c.producto,"
                    + " c.mercado, c.fecha_desde, c.fecha_hasta, c.bl, c.fee, c.volumen, c.operacion"
                    + " FROM fact.ag_pt_clicks c inner join "
                    + " ag_pt_inventario i on"
                    + " i.cpe = c.cpe"
                    + " where fecha_desde <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + fh.ToString("yyyy-MM-dd") + "'"
                    + " order by c.cpe, c.click";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.ClicksPT c = new EndesaEntity.facturacion.ClicksPT();
                    c.nif = r["nif"].ToString();
                    c.cliente = r["cliente"].ToString();
                    c.cpe = r["cpe"].ToString();
                    c.click = Convert.ToInt32(r["click"]);
                    c.mercado = r["mercado"].ToString();
                    c.operacion = r["operacion"].ToString();
                    c.fecha_operacion = Convert.ToDateTime(r["fecha"]);
                    c.producto = r["producto"].ToString();
                    c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                    c.bl = Convert.ToDouble(r["bl"]);
                    c.fee = Convert.ToDouble(r["fee"]);
                    c.volumen = Convert.ToDouble(r["volumen"]);

                    List<EndesaEntity.facturacion.ClicksPT> o;
                    if (!d.TryGetValue(c.cpe, out o))
                    {
                        o = new List<EndesaEntity.facturacion.ClicksPT>();
                        o.Add(c);
                        d.Add(c.cpe, o);
                    }
                    else
                        o.Add(c);

                }

                db.CloseConnection();
                return d;

            }
            catch (Exception e)
            {

                MessageBox.Show("Carga completada satisfactoriamente.",
                  "Clicks.Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
                return null;
            }


        }

        public List<EndesaEntity.facturacion.ClicksPT> GetClicks(string cups20)
        {
            List<EndesaEntity.facturacion.ClicksPT> o;
            if (dic.TryGetValue(cups20, out o))
                return o;
            else
                return null;
        }

        public void ImportarClicks(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            int linea = 0;
            bool firstOnlyCabecera = true;
            bool firstOnly = true;
            string cabecera = "";
            DateTime fecha = new DateTime();
            string fecha_texto = "";
            List<string> lista_cups = new List<string>();

            List<EndesaEntity.facturacion.ClicksPT> lista = new List<EndesaEntity.facturacion.ClicksPT>();

            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;

            try
            {

                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();



                for (int i = 2; i < 10000; i++)
                {

                    if (workSheet.Cells[i, 1].Value != null)
                    {
                        f++;
                        lista_cups.Add(workSheet.Cells[i, 1].Value.ToString());
                    }
                }

                workSheet = excelPackage.Workbook.Worksheets[2];

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 10000; i++)
                {
                    linea = 1;
                    c = 1;

                    if (firstOnlyCabecera)
                    {
                        for (int w = 1; w < 13; w++)
                            cabecera += workSheet.Cells[1, w].Value.ToString();

                        if (!EstructuraCorrectaClicks(cabecera))
                        {


                            MessageBox.Show("La estructura del archivo no es correcta y no se importará ningún valor.",
                                "Estructura Incorrecta!!!",
                                MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                            break;
                        }
                        else
                        {
                            f = 3;
                        }


                        firstOnlyCabecera = false;
                    }



                    cabecera = "";
                    if (workSheet.Cells[f, 1].Value != null)
                    {

                        EndesaEntity.facturacion.ClicksPT s = new EndesaEntity.facturacion.ClicksPT();
                        s.click = Convert.ToInt32(workSheet.Cells[f, 1].Value.ToString());
                        fecha_texto = workSheet.Cells[f, 4].Value.ToString();
                        fecha = new DateTime(Convert.ToInt32(fecha_texto.Substring(6, 4)),
                                                Convert.ToInt32(fecha_texto.Substring(3, 2)),
                                                Convert.ToInt32(fecha_texto.Substring(0, 2)));
                        fecha = fecha.AddHours(Convert.ToInt32(fecha_texto.Substring(11, 2)));

                        s.fecha_operacion = fecha;
                        s.mercado = workSheet.Cells[f, 5].Value.ToString();
                        s.operacion = workSheet.Cells[f, 6].Value.ToString();
                        s.producto = workSheet.Cells[f, 7].Value.ToString();
                        List<DateTime> l = FechasDesdeProductoClick(s.producto);
                        s.fecha_desde = l[0];
                        s.fecha_hasta = l[1];
                        s.bl = Convert.ToDouble(workSheet.Cells[f, 10].Value.ToString());
                        s.fee = Convert.ToDouble(workSheet.Cells[f, 11].Value.ToString());
                        s.volumen = Convert.ToDouble(workSheet.Cells[f, 12].Value.ToString());
                        f++;

                        lista.Add(s);

                    }

                }

                fs = null;
                excelPackage = null;

                id = 0;
                for (int h = 0; h < lista_cups.Count; h++)
                {

                    firstOnly = true;
                    for (int i = 0; i < lista.Count; i++)
                    {
                        id++;
                        if (firstOnly)
                        {
                            sb.Append("replace into ag_pt_clicks (cpe, click, fecha, ");
                            sb.Append("mercado, operacion, producto, fecha_desde, fecha_hasta, ");
                            sb.Append("bl, fee, volumen, usuario) values ");
                            firstOnly = false;
                        }
                        sb.Append("('").Append(lista_cups[h]).Append("',");
                        sb.Append(lista[i].click).Append(",");
                        sb.Append("'").Append(lista[i].fecha_operacion.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                        sb.Append("'").Append(lista[i].mercado).Append("',");
                        sb.Append("'").Append(lista[i].operacion).Append("',");
                        sb.Append("'").Append(lista[i].producto).Append("',");
                        sb.Append("'").Append(lista[i].fecha_desde.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(lista[i].fecha_hasta.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append(lista[i].bl.ToString().Replace(",", ".")).Append(",");
                        sb.Append(lista[i].fee.ToString().Replace(",", ".")).Append(",");
                        sb.Append(lista[i].volumen.ToString().Replace(",", ".")).Append(",");
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

            }
            catch (Exception e)
            {
                MessageBox.Show("Error en la línea " + linea + " --> " + e.Message,
                  "Error en la importación de Clicks",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private bool EstructuraCorrectaClicks(string linea)
        {
            return linea.Equals("#JustificanteLoteFechaMercadoOperaciónProductoFecha desdeFecha hastaBLFeeVolumen [%]");
        }

        private List<DateTime> FechasDesdeProductoClick(string producto)
        {
            string periodo = producto.Substring(0, 2);
            int anio = Convert.ToInt32(producto.Substring(3, 2));
            DateTime fechaDesde = new DateTime();
            DateTime fechaHasta = new DateTime();
            List<DateTime> l = new List<DateTime>();
            switch (periodo)
            {
                case "YR":
                    fechaDesde = new DateTime((2000) + anio, 01, 01);
                    fechaHasta = new DateTime((2000) + anio, 12, 31);
                    break;
                case "Q1":
                    fechaDesde = new DateTime((2000) + anio, 01, 01);
                    fechaHasta = new DateTime((2000) + anio, 03, 31);
                    break;
                case "Q2":
                    fechaDesde = new DateTime((2000) + anio, 04, 01);
                    fechaHasta = new DateTime((2000) + anio, 06, 30);
                    break;
                case "Q3":
                    fechaDesde = new DateTime((2000) + anio, 07, 01);
                    fechaHasta = new DateTime((2000) + anio, 09, 30);
                    break;
                case "Q4":
                    fechaDesde = new DateTime((2000) + anio, 10, 01);
                    fechaHasta = new DateTime((2000) + anio, 12, 31);
                    break;
            }

            l.Add(fechaDesde);
            l.Add(fechaHasta);
            return l;
        }
    }
}
