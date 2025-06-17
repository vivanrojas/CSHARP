using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class ContratosGasDetalle_Informe
    {
        public Dictionary<string, List<ContratoGasDetalle>> dic { get; set; }
        Dictionary<string, ContratoGasDetalle> estado_factura;

        public ContratosGasDetalle_Informe(DateTime fd, DateTime fh, string cups20)
        {
            estado_factura = new Dictionary<string, ContratoGasDetalle>();
            dic = new Dictionary<string, List<ContratoGasDetalle>>();
            CargaEstadosFactura(fd, fh, cups20);
            LoadData(fd, fh, cups20);
        }

        private void LoadData(DateTime fd, DateTime fh, string cups20)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {

                strSql = "select"
                    + " c.cnifdnic,"
                    + " c.dapersoc,"
                    + " cd.cups20, cd.fecha_inicio, cd.fecha_fin, cd.qd,"
                    + " cd.tarifa, cd.id_solicitud, cd.linea_solicitud,"
                    + " cd.tipo, cd.last_update_date, cd.qi, cd.hora_inicio"
                    + " from cont.atrgas_contratos c"
                    + " inner join cont.atrgas_contratos_detalle cd on"
                    + " cd.nif = c.cnifdnic and"
                    + " cd.cups20 = c.cups20"
                    + " where";

                if (cups20 != null)
                    strSql += " c.cups20 = '" + cups20 + "' and";

                strSql += " (cd.fecha_inicio <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                    + " (cd.fecha_fin >= '" + fd.ToString("yyyy-MM-dd") + "' or cd.fecha_fin is null))"
                    // + " group by c.cups20, cd.fecha_inicio, cd.fecha_fin, cd.tipo, cd.comentario"
                     + " order by c.cnifdnic, c.cups20, cd.fecha_inicio asc";
                                
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    ContratoGasDetalle c = new ContratoGasDetalle();

                    if (r["cnifdnic"] != System.DBNull.Value)
                        c.vatnum = r["cnifdnic"].ToString();

                    if (r["dapersoc"] != System.DBNull.Value)
                        c.customer_name = r["dapersoc"].ToString();

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["hora_inicio"] != System.DBNull.Value)
                        c.hora_inicio = Convert.ToDateTime(r["hora_inicio"]);

                    if (r["qd"] != System.DBNull.Value)
                        c.qd = Convert.ToDouble(r["qd"]);

                    if (r["qi"] != System.DBNull.Value)
                        c.qi = Convert.ToDouble(r["qi"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["tipo"] != System.DBNull.Value)
                        c.tipo = r["tipo"].ToString();

                    c.cups20 = r["cups20"].ToString();

                    c.last_update_date = Convert.ToDateTime(r["last_update_date"]);

                    if (r["id_solicitud"] != System.DBNull.Value)
                        c.id_solicitud = Convert.ToInt64(r["id_solicitud"]);

                    if (r["linea_solicitud"] != System.DBNull.Value)
                        c.linea_solicitud = Convert.ToInt32(r["linea_solicitud"]);


                    ContratoGasDetalle con;
                    if (estado_factura.TryGetValue(c.cups20, out con))
                    {
                        c.estado = con.estado;
                        c.estadoContrato = con.estadoContrato;
                        c.ultimo_mes_facturado = con.ultimo_mes_facturado;
                    }

                    // Solo tratamos los anuales y trimestrales que vengan de adicionales.
                    // Para saber si vienen de adicionales deben tener informado el 
                    // id_solicitud y linea_solicitud.

                    if((c.tipo == "ANUAL" || c.tipo == "TRIMESTRAL") &&  
                        (c.id_solicitud > 0 && c.linea_solicitud > 0))
                    {
                        List<ContratoGasDetalle> o;
                        if (!dic.TryGetValue(c.cups20, out o))
                        {
                            o = new List<ContratoGasDetalle>();
                            o.Add(c);
                            dic.Add(c.cups20, o);
                        }
                        else
                            o.Add(c);
                    }else if ((c.tipo != "ANUAL" && c.tipo != "TRIMESTRAL"))
                    {
                        List<ContratoGasDetalle> o;
                        if (!dic.TryGetValue(c.cups20, out o))
                        {
                            o = new List<ContratoGasDetalle>();
                            o.Add(c);
                            dic.Add(c.cups20, o);
                        }
                        else
                            o.Add(c);
                    }




                }
                db.CloseConnection();

                // Cada vez que se saca el informe actualizamos las addendas.
                // 2023027




            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ContratosGasDetalle_Informe.LoadData", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CargaEstadosFactura(DateTime fd, DateTime fh, string cups20)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime _fd;
            DateTime _fh;
            DateTime mesAnterior;
            int anio;
            int mes;
            int dias_del_mes;

            try
            {

                mesAnterior = DateTime.Now.AddMonths(-1);
                anio = mesAnterior.Year;
                mes = mesAnterior.Month;
                dias_del_mes = DateTime.DaysInMonth(anio, mes);

                _fh = new DateTime(anio, mes, dias_del_mes);
                _fd = new DateTime(fh.Year, fh.Month, 1);

                if (fd.Month == _fd.Month && fh.Month == _fh.Month)
                {
                    strSql = "select i.ID_PS, i.CUPSREE,"
                        + " h.EstadoContrato, h.UltimoMesFacturado, h.Estado, h.Tipo"
                        + " from fact.cm_inventario_gas i"
                        + " left outer join fact.cm_detalle_hist h on"
                        + " h.IdPS = i.ID_PS where";

                    if (cups20 != null)
                        strSql += " i.CUPSREE = '" + cups20 + "' and";

                    strSql += " h.f_ult_mod = (select max(f_ult_mod) f_ult_mod from fact.cm_detalle_hist) and"
                        + " (i.FINICIO <= '" + fd.ToString("yyyy-MM-dd") + "'"
                        + " and (i.FFIN >= '" + fh.ToString("yyyy-MM-dd") + "' or i.FFIN is null)) group by i.CUPSREE;";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        ContratoGasDetalle c = new ContratoGasDetalle();

                        if (r["CUPSREE"] != System.DBNull.Value)
                        {
                            c.cups20 = r["CUPSREE"].ToString();
                            if (r["EstadoContrato"] != System.DBNull.Value)
                                c.estadoContrato = r["EstadoContrato"].ToString();
                            if (r["UltimoMesFacturado"] != System.DBNull.Value)
                                c.ultimo_mes_facturado = Convert.ToInt32(r["UltimoMesFacturado"]);
                            if (r["Estado"] != System.DBNull.Value)
                                c.estado = r["Estado"].ToString();
                            ContratoGasDetalle o;
                            if (!estado_factura.TryGetValue(c.cups20, out o))
                                estado_factura.Add(c.cups20, c);

                        }

                    }
                    db.CloseConnection();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ContratosGasDetalle_Informe.CargaEstadosFactura", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

       
    }
}
