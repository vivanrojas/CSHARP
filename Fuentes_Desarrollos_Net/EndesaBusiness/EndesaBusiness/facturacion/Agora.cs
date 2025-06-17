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
    public class Agora
    {
        public bool hayError { get; set; }
        public string descripcion_error { get; set; }

        public List<EndesaEntity.AgoraDetalle> ld { get; set; }
        public List<EndesaEntity.AgoraResumen> lr { get; set; }
        facturacion.Alarmas alarm;
        medida.TAM tam;
        contratacion.PS_AT_Funciones ps;

        public Agora(DateTime fdTrabajo, DateTime fhTrabajo)
        {
            ld = new List<EndesaEntity.AgoraDetalle>();
            lr = new List<EndesaEntity.AgoraResumen>();
            //alarm = new alarmas.Alarmas();
            //tam = new medida.tam.TAM();
            //tam.CargaTAM();
            //Carga();
            //CreaResumen();
            //ps = new contratacion.ps_at.PS_AT_Funciones();
            //ps.Carga_PS_AT();
            //CargaAgora(ld, fdTrabajo, fhTrabajo);
            //CargaSofisticados(ld, fdTrabajo, fhTrabajo);
        }

        private void Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            DateTime ultima_fecha = new DateTime();

            try
            {
                sb.Append("select max(F_ULT_MOD) as f_ult_mod from cm_detalle_hist");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    ultima_fecha = Convert.ToDateTime(r["f_ult_mod"]);
                }
                db.CloseConnection();

                sb.Clear();
                sb.Append("select * from fact.cm_detalle_hist h where h.F_ULT_MOD = '").Append(ultima_fecha.ToString("yyyy-MM-dd")).Append("'");
                sb.Append(" and h.Tipo in('AGORA SCE', 'AGORA MANUAL') order by Mes, Estado");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.AgoraDetalle d = new EndesaEntity.AgoraDetalle();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        d.nif = r["CNIFDNIC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        d.nombre_cliente = r["DAPERSOC"].ToString();

                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        d.cups13 = r["CCOUNIPS"].ToString();

                    if (r["ESTADO"] != System.DBNull.Value)
                        d.estado_ltp = r["ESTADO"].ToString();

                    if (r["TIPO"] != System.DBNull.Value)
                        d.tipo = r["TIPO"].ToString();

                    if (r["Mes"] != System.DBNull.Value)
                        d.ultimo_mes_facturado = r["Mes"].ToString();

                    d.tam = tam.GetTAM(d.cups13);

                    ld.Add(d);
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {

            }
        }

        private void CargaAgora(List<EndesaEntity.AgoraDetalle> ld, DateTime fdTrabajo, DateTime fhTrabajo)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Append("select ps.CCOUNIPS, r.CUPS20 ");
                sb.Append(" from fact.contratosPS ps");
                sb.Append(" left outer join relacion_cups r on");
                sb.Append(" r.CUPS_CORTO = ps.CCOUNIPS");
                sb.Append(" where  (('").Append(fdTrabajo.ToString("yyyy-MM-dd")).Append("' <= ps.FFINVESU or");
                sb.Append(" ps.FFINVESU is null) and");
                sb.Append(" ('").Append(fhTrabajo.ToString("yyyy-MM-dd")).Append("' <= ps.FBAJACON  or");
                sb.Append(" ps.FBAJACON is null)) and");
                sb.Append(" ps.FPSERCON < '").Append(fhTrabajo.ToString("yyyy-MM-dd")).Append("' and");
                sb.Append(" ps.CCOMPOBJ = 'A01'");
                sb.Append(" group by ps.CCOUNIPS,ps.FFINVESU;");

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.AgoraDetalle d = new EndesaEntity.AgoraDetalle();
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        d.cups13 = r["CCOUNIPS"].ToString();

                    if (r["CUPS20"] != System.DBNull.Value)
                    {
                        d.cups20 = r["CUPS20"].ToString();
                        d.alarma = alarm.GetAlarma(d.cups20);
                    }

                    d.tam = tam.GetTAM(d.cups13);

                    if (d.alarma != "")
                        d.tipo = "AGORA MANUAL";
                    else
                        d.tipo = "AGORA SCE";


                    EndesaEntity.contratacion.PS_AT o;
                    if (ps.l_ps_at.TryGetValue(d.cups20, out o))
                    {
                        d.nif = o.nif;
                        d.nombre_cliente = o.nombre_cliente;

                    }

                    ld.Add(d);
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {

            }
        }

        private void CargaSofisticados(List<EndesaEntity.AgoraDetalle> ld, DateTime fdTrabajo, DateTime fhTrabajo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Append("Select s.CCOUNIPS, s.PRECIOS, r.CUPS20 from cm_sofisticados s");
                sb.Append(" left outer join relacion_cups r on");
                sb.Append(" r.CUPS_CORTO = s.CCOUNIPS");
                sb.Append(" where (s.FD <= '").Append(fdTrabajo.ToString("yyyy-MM-dd")).Append("' and ");
                sb.Append("s.FH >= '").Append(fhTrabajo.ToString("yyyy-MM-dd")).Append("' or s.FH is null);");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.AgoraDetalle d;
                    d = ld.Find(z => z.cups13 == r["CCOUNIPS"].ToString());
                    if (d == null)
                    {
                        d = new EndesaEntity.AgoraDetalle();

                        if (r["CCOUNIPS"] != System.DBNull.Value)
                            d.cups13 = r["CCOUNIPS"].ToString();

                        if (r["PRECIOS"] != System.DBNull.Value)
                            d.precios = r["PRECIOS"].ToString();

                        if (r["CUPS20"] != System.DBNull.Value)
                        {
                            d.cups20 = r["CUPS20"].ToString();
                            d.alarma = alarm.GetAlarma(d.cups20);
                        }

                        d.tam = tam.GetTAM(d.cups13);

                        if (d.alarma != "")
                            d.tipo = "AGORA MANUAL";
                        else
                            d.tipo = "AGORA SCE";

                        EndesaEntity.contratacion.PS_AT o;
                        if (ps.l_ps_at.TryGetValue(d.cups20, out o))
                        {
                            d.nif = o.nif;
                            d.nombre_cliente = o.nombre_cliente;

                        }

                        ld.Add(d);
                    }
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {

            }
        }

        public void CreaResumen()
        {
            bool firstOnly = true;
            bool encontrado = false;

            int total_cups = 0;
            int total_cups_tam_no_blancos = 0;
            double media = 0;
            double suma = 0;

            lr.Clear();

            #region Datos <> Facturado
            for (int i = 0; i < ld.Count(); i++)
            {
                if (ld[i].ultimo_mes_facturado != null)
                {
                    if (firstOnly)
                    {
                        EndesaEntity.AgoraResumen r = new EndesaEntity.AgoraResumen();
                        r.ltp = TipoLTP(ld[i].estado_ltp);
                        r.promedio_tam = ld[i].tam;
                        r.sum_tam = ld[i].tam;
                        r.primer_mes_pdte = ld[i].ultimo_mes_facturado;
                        r.num_cups = 1;

                        if (!ld[i].tam_en_blanco)
                            total_cups_tam_no_blancos++;

                        lr.Add(r);
                        // Creamos el otro estado para que simpre salgan dos cada vez que tengamos un mes pdte.
                        r = new EndesaEntity.AgoraResumen();
                        r.ltp = this.CompletaOtroEstado(TipoLTP(ld[i].estado_ltp));
                        r.promedio_tam = 0;
                        r.sum_tam = 0;
                        r.primer_mes_pdte = ld[i].ultimo_mes_facturado;
                        r.num_cups = 0;

                        lr.Add(r);
                        firstOnly = false;
                    }
                    else
                    {
                        for (int j = 0; j < lr.Count(); j++)
                        {
                            if (
                                (lr[j].primer_mes_pdte == ld[i].ultimo_mes_facturado) &&
                                (lr[j].ltp == TipoLTP(ld[i].estado_ltp))
                                )
                            {
                                encontrado = true;
                                lr[j].num_cups = lr[j].num_cups + 1;
                                if (!ld[i].tam_en_blanco)
                                    lr[j].num_cups_con_tam = lr[j].num_cups_con_tam + 1;
                                lr[j].sum_tam = lr[j].sum_tam + ld[i].tam;
                                lr[j].promedio_tam = lr[j].sum_tam / lr[j].num_cups_con_tam;
                                break;
                            }
                            else
                                encontrado = false;
                        }
                        if (!encontrado)
                        {
                            EndesaEntity.AgoraResumen r = new EndesaEntity.AgoraResumen();
                            r.ltp = TipoLTP(ld[i].estado_ltp);
                            r.promedio_tam = ld[i].tam;
                            r.sum_tam = ld[i].tam;
                            r.primer_mes_pdte = ld[i].ultimo_mes_facturado;
                            r.num_cups = 1;
                            if (!ld[i].tam_en_blanco)
                                r.num_cups_con_tam = 1;
                            lr.Add(r);

                            // Creamos el otro estado para que simpre salgan dos cada vez que tengamos un mes pdte.
                            r = new EndesaEntity.AgoraResumen();
                            r.ltp = this.CompletaOtroEstado(TipoLTP(ld[i].estado_ltp));
                            r.promedio_tam = 0;
                            r.sum_tam = 0;
                            r.primer_mes_pdte = ld[i].ultimo_mes_facturado;
                            r.num_cups = 0;

                            lr.Add(r);


                        }

                    }
                }

            }
            #endregion

            lr = lr.OrderBy(z => z.primer_mes_pdte).ThenBy(z => z.ltp).ToList();

            #region Total Pdte
            for (int i = 0; i < ld.Count(); i++)
            {
                if (ld[i].ultimo_mes_facturado != null)
                {
                    total_cups++;
                    suma += ld[i].tam;
                }
            }

            if (total_cups != 0)
            {
                media = suma / total_cups;
                EndesaEntity.AgoraResumen r = new EndesaEntity.AgoraResumen();
                r.ltp = "TOTAL PDTE MEDIDA/FACTURACION";
                r.promedio_tam = media;
                r.sum_tam = suma;
                r.num_cups = total_cups;
                lr.Add(r);
            }


            #endregion

            #region Facturado
            total_cups_tam_no_blancos = 0;
            for (int i = 0; i < ld.Count(); i++)
            {
                if (ld[i].ultimo_mes_facturado == null)
                {
                    if (firstOnly)
                    {
                        EndesaEntity.AgoraResumen r = new EndesaEntity.AgoraResumen();
                        r.ltp = TipoLTP(ld[i].estado_ltp);
                        r.promedio_tam = ld[i].tam;
                        r.sum_tam = ld[i].tam;
                        r.primer_mes_pdte = ld[i].ultimo_mes_facturado;
                        r.num_cups = 1;
                        if (!ld[i].tam_en_blanco)
                            r.num_cups_con_tam = 1;
                        lr.Add(r);
                        firstOnly = false;
                    }
                    else
                    {
                        for (int j = 0; j < lr.Count(); j++)
                        {
                            if (
                                (lr[j].primer_mes_pdte == ld[i].ultimo_mes_facturado) &&
                                (lr[j].ltp == TipoLTP(ld[i].estado_ltp))
                                )
                            {
                                encontrado = true;
                                if (!ld[i].tam_en_blanco)
                                    lr[j].num_cups_con_tam = lr[j].num_cups_con_tam + 1;
                                lr[j].num_cups = lr[j].num_cups + 1;
                                lr[j].sum_tam = lr[j].sum_tam + ld[i].tam;
                                //lr[j].promedio_tam = lr[j].sum_tam / lr[j].num_cups;
                                lr[j].promedio_tam = lr[j].sum_tam / lr[j].num_cups_con_tam;
                                break;
                            }
                            else
                                encontrado = false;
                        }
                        if (!encontrado)
                        {
                            EndesaEntity.AgoraResumen r = new EndesaEntity.AgoraResumen();
                            r.ltp = TipoLTP(ld[i].estado_ltp);
                            r.promedio_tam = ld[i].tam;
                            r.sum_tam = ld[i].tam;
                            r.primer_mes_pdte = ld[i].ultimo_mes_facturado;
                            r.num_cups = 1;
                            lr.Add(r);
                        }

                    }
                }

            }
            #endregion

            #region totales
            total_cups = 0;
            total_cups_tam_no_blancos = 0;
            suma = 0;
            for (int i = 0; i < ld.Count(); i++)
            {
                if (!ld[i].tam_en_blanco)
                    total_cups_tam_no_blancos++;

                total_cups++;
                suma += ld[i].tam;
            }

            if (total_cups != 0)
            {
                media = suma / total_cups_tam_no_blancos;
                EndesaEntity.AgoraResumen r = new EndesaEntity.AgoraResumen();
                r.ltp = "TOTAL";
                r.promedio_tam = media;
                r.sum_tam = suma;
                r.num_cups = total_cups;
                lr.Add(r);
            }
            #endregion


        }

        private string TipoLTP(string ltp)
        {
            string tipoLTP = "";
            switch (ltp)
            {
                case "01. Pendiente de medida":
                    tipoLTP = "Falta medida";
                    break;
                case "02. CC Rechazada por CS":
                    tipoLTP = "Falta medida";
                    break;
                case "03. CC Completa en CS":
                    tipoLTP = "Falta medida";
                    break;
                case "04. CC Enviada a SCE ML":
                    tipoLTP = "Falta medida";
                    break;
                case "05. CC Rechazada por SCE ML":
                    tipoLTP = "Falta medida";
                    break;
                case "06. CC Incompleta  SCE ML":
                    tipoLTP = "Falta medida";
                    break;
                case "07. LTP SCE":
                    tipoLTP = "Falta medida";
                    break;
                case "08. El Punto no está Extraído":
                    tipoLTP = "Pdte. facturar";
                    break;
                case "09. El Punto está Extraído":
                    tipoLTP = "Pdte. facturar";
                    break;
                case "10. Prefactura pendiente":
                    tipoLTP = "Pdte. facturar";
                    break;
                default:
                    tipoLTP = "TOTAL FACTURADOS";
                    break;
            }

            return tipoLTP;
        }

        private string CompletaOtroEstado(string ltp)
        {
            string otroEstado = "";
            switch (ltp)
            {
                case "Falta medida":
                    otroEstado = "Pdte. facturar";
                    break;
                case "Pdte. facturar":
                    otroEstado = "Falta medida";
                    break;
                default:
                    otroEstado = "ERROR";
                    break;
            }

            return otroEstado;
        }

        public void CargaExcel(string fichero)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            string cabecera = "";

            List<EndesaEntity.AgoraDetalle> l = new List<EndesaEntity.AgoraDetalle>();

            try
            {
                // FileInfo file = new FileInfo(fichero);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                f = 1; // Porque la primera fila es la cabecera
                for (int i = 0; i < 250000; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        for (int w = 1; w <= 9; w++)
                            cabecera += utilidades.FuncionesTexto.ArreglaAcentos(workSheet.Cells[1, w].Value.ToString());


                        if (!EstructuraCorrecta(cabecera))
                        {
                            this.hayError = true;
                            this.descripcion_error = "La estructura del archivo excel no es la correcta." + System.Environment.NewLine
                                + " La cabecera debe contener las columnas: " + System.Environment.NewLine
                                + " NIF CLIENTE CCOUNIPS MES ESTADO TIPO TAM " + System.Environment.NewLine
                                + " Gestor Posición Gestor Posición Cuenta	" + System.Environment.NewLine;
                                

                            break;
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null &&
                         workSheet.Cells[f, 2].Value == null &&
                         workSheet.Cells[f, 3].Value == null)
                    {
                        break;
                    }
                    else
                    {
                        EndesaEntity.AgoraDetalle d = new EndesaEntity.AgoraDetalle();
                        if (workSheet.Cells[f, c].Value != null)
                            d.nif = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.nombre_cliente = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.cups13 = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.ultimo_mes_facturado = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.estado_ltp = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.tipo = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.tam = Convert.ToDouble(workSheet.Cells[f, c].Value.ToString());
                        else
                            d.tam_en_blanco = true;

                        c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.nombreGestor = workSheet.Cells[f, c].Value.ToString();

                        c++;

                        //if (workSheet.Cells[f, c].Value != null)
                        //    d.apellido1 = workSheet.Cells[f, c].Value.ToString();

                        //c++;

                        //if (workSheet.Cells[f, c].Value != null)
                        //    d.apellido2 = workSheet.Cells[f, c].Value.ToString();

                        //c++;

                        //if (workSheet.Cells[f, c].Value != null)
                        //    d.desc_Responsable_Territorial = workSheet.Cells[f, c].Value.ToString();

                        //c++;

                        if (workSheet.Cells[f, c].Value != null)
                            d.subDireccion = workSheet.Cells[f, c].Value.ToString();

                        ld.Add(d);
                        id++;
                    }

                }

                fs = null;
                excelPackage = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Agora - CargaExcel",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private bool EstructuraCorrecta(string cabecera)
        {
            
            //if (cabecera.Trim() == "NIFClienteCCOUNIPSMESEstadoTipoTAMNombre GestorApellido 1Apellido 2Desc Responsable TerritorialSubdirección")
            if (cabecera.Trim() == "NIFClienteCCOUNIPSMESEstadoTipoTAMGestorPosición Gestor")
                return true;
            else
                return false;
        }
    }
}
