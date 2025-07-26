using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class RetrasoFacturacion
    {
        public DateTime fd { get; set; }
        public DateTime fh { get; set; }
        public String tipoNegocio { get; set; }
        public String cupsREE { get; set; }
        public String cnifdnic { get; set; }
        public String tipoFactura { get; set; }
        public String conceptos { get; set; }
        public Boolean fechaFactura { get; set; }
        public string empresas { get; set; }

        List<EndesaEntity.facturacion.TotalRetrasoFacturacion> lt = 
            new List<EndesaEntity.facturacion.TotalRetrasoFacturacion>();
        List<EndesaEntity.facturacion.InformeRetrasoFacturacion> li =
            new List<EndesaEntity.facturacion.InformeRetrasoFacturacion>();


        public RetrasoFacturacion()
        {
            this.tipoNegocio = "";
            this.cnifdnic = "";
            this.cupsREE = "";
            this.tipoFactura = "";
            this.conceptos = "";
            this.empresas = "";
            this.fechaFactura = false;
        }

        public void CargaDatos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            Boolean firstOnly = true;
            Boolean encontradoNIF = true;
            Boolean encontradoMES = true;
            Boolean encontradoCUP = true;

            EndesaEntity.facturacion.MesDatosRetrasoFacturacion mes;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(this.GetSQLDatos(), db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.facturacion.TotalRetrasoFacturacion t =
                        new EndesaEntity.facturacion.TotalRetrasoFacturacion();
                    t.cnifdnic = reader["cnifdnic"].ToString();
                    t.dapersoc = reader["dapersoc"].ToString();
                    t.cups = reader["cupsree"].ToString();
                    t.fa = Convert.ToDateTime(reader["FFACTURA"]);
                    t.fd = Convert.ToDateTime(reader["FFACTDES"]);
                    t.fh = Convert.ToDateTime(reader["FFACTHAS"]);
                    t.dias = Convert.ToInt32(reader["DIAS"]);
                    lt.Add(t);

                }

                for (Int32 i = 0; i < lt.Count; i++)
                {
                    if (firstOnly)
                    {
                        EndesaEntity.facturacion.InformeRetrasoFacturacion ii =
                            new EndesaEntity.facturacion.InformeRetrasoFacturacion();
                        mes = new EndesaEntity.facturacion.MesDatosRetrasoFacturacion();

                        mes.mes = lt[i].fd.Month;
                        mes.dias = lt[i].dias;
                        mes.cupsree.Add(lt[i].cups);
                        mes.max = lt[i].dias;
                        mes.veces = 1;
                        mes.year = lt[i].fd.Year;

                        ii.cnifdnic = lt[i].cnifdnic;
                        ii.dapersoc = lt[i].dapersoc;

                        ii.listameses.Add(mes);

                        li.Add(ii);
                        firstOnly = false;
                    }
                    else
                    {
                        encontradoNIF = false;
                        for (Int32 j = 0; j < li.Count; j++)
                        {
                            if (li[j].cnifdnic == lt[i].cnifdnic)
                            {
                                // TRATAMOS MESES
                                encontradoMES = false;
                                for (int x = 0; x < li[j].listameses.Count; x++)
                                {
                                    if (li[j].listameses[x].mes == lt[i].fd.Month)
                                    {
                                        // TRATAMOS CUPS
                                        encontradoCUP = false;
                                        for (int h = 0; h < li[j].listameses[x].cupsree.Count; h++)
                                        {
                                            if ((li[j].listameses[x].cupsree[h] == lt[i].cups) &&
                                                (li[j].dapersoc == lt[j].dapersoc))
                                            {
                                                li[j].listameses[x].veces = li[j].listameses[x].veces + 1;
                                                li[j].listameses[x].dias = li[j].listameses[x].dias + lt[i].dias;
                                                li[j].listameses[x].max = li[j].listameses[x].max > lt[i].dias ? li[j].listameses[x].max : lt[i].dias;
                                                encontradoCUP = true;
                                            }
                                        }
                                        if (!encontradoCUP)
                                        {
                                            li[j].listameses[x].cupsree.Add(lt[i].cups);
                                            li[j].listameses[x].veces = li[j].listameses[x].veces + 1;
                                            li[j].listameses[x].dias = li[j].listameses[x].dias + lt[i].dias;
                                            li[j].listameses[x].max = li[j].listameses[x].max > lt[i].dias ? li[j].listameses[x].max : lt[i].dias;
                                        }
                                        encontradoMES = true;
                                    }
                                }

                                if (!encontradoMES)
                                {
                                    mes = new EndesaEntity.facturacion.MesDatosRetrasoFacturacion();
                                    mes.mes = lt[i].fd.Month;
                                    mes.dias = lt[i].dias;
                                    mes.cupsree.Add(lt[i].cups);
                                    mes.veces = 1;
                                    mes.max = lt[i].dias;
                                    mes.year = lt[i].fd.Year;
                                    li[j].listameses.Add(mes);
                                }
                                encontradoNIF = true;
                            }
                        }

                        if (!encontradoNIF)
                        {
                            EndesaEntity.facturacion.InformeRetrasoFacturacion ii =
                                new EndesaEntity.facturacion.InformeRetrasoFacturacion();
                            mes = new EndesaEntity.facturacion.MesDatosRetrasoFacturacion();

                            ii.cnifdnic = lt[i].cnifdnic;
                            ii.dapersoc = lt[i].dapersoc;
                            mes.mes = lt[i].fd.Month;
                            mes.cupsree.Add(lt[i].cups);
                            mes.max = lt[i].dias;
                            mes.veces = 1;
                            mes.year = lt[i].fd.Year;
                            ii.listameses.Add(mes);
                            li.Add(ii);
                        }
                    }
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la construcción de las datos en Retraso facturación",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }


        public void PintarExcel()
        {
            office.Excel excel = new office.Excel();
            excel.CambiaNombreHoja("RetrasoFacturación");
            Int32 c = 0;
            Int32 f = 0;
            f = 1;
            c = 2;

            for (int d = 0; d < li[0].listameses.Count; d++)
            {
                for (int i = 1; i <= 4; i++)
                {
                    excel.PonValor(f, c, this.MesLetra(li[0].listameses[d].mes) + " " + li[0].listameses[d].year.ToString());
                    c++;
                }
            }

            f = 2;
            c = 1;
            excel.PonValor(f, c, "NIF");
            c++;

            for (int d = 0; d < li[0].listameses.Count; d++)
            {
                excel.PonValor(f, c, "Nº CUPS");
                c++;
                excel.PonValor(f, c, "Nº DIAS");
                c++;
                excel.PonValor(f, c, "MÁXIMO");
                c++;
                excel.PonValor(f, c, "DIF");
                c++;
            }

            for (int i = 1; i <= c; i++)
            {
                excel.PonEstilo(1, i, office.Excel.Estilos.NEGRITA);
            }

            for (Int32 i = 0; i < li.Count; i++)
            {
                f++;
                c = 1;
                excel.PonValor(f, c, li[i].cnifdnic);

                for (int j = 0; j < li[i].listameses.Count; j++)
                {
                    c++;
                    excel.PonValor(f, c, li[i].listameses[j].veces);
                    c++;
                    excel.PonValor(f, c, Math.Round(Convert.ToDouble(li[i].listameses[j].dias / li[i].listameses[j].veces), 0));
                    c++;
                    excel.PonValor(f, c, li[i].listameses[j].max);
                    c++;
                    excel.PonValor(f, c, li[i].listameses[j].max - Math.Round(Convert.ToDouble(li[i].listameses[j].dias / li[i].listameses[j].veces), 0));

                }
            }

            excel.AjustarAncho();
            excel.Mostrar();
            excel.Cerrar();
            excel = null;

        }

        public String GetSQLDatos()
        {
            string strSql = "";
            String[] linNeg;
            String[] listEmpresas;
            String[] listConceptos;
            Boolean firstOnly = true;
            Boolean firstOnly2 = true;

            try
            {


                strSql = "select f.CNIFDNIC, f.DAPERSOC, f.CUPSREE, f.CFACTURA, f.FFACTDES, f.FFACTHAS,"
                    + "min(round(DATEDIFF(f.FFACTURA, f.FFACTHAS), 0)) as dias, f.FFACTURA, f.IFACTURA"
                    + " from fo f where";

                if (!this.fechaFactura)
                {
                    strSql = strSql + " (f.FFACTDES >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTHAS <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }
                else
                {
                    strSql = strSql + " (f.FFACTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " f.FFACTURA <= '" + fh.ToString("yyyy-MM-dd") + "')";
                }

                // LINEA DE NEGOCIO (LUZ / GAS)

                if (this.tipoNegocio != "")
                {
                    linNeg = this.tipoNegocio.Split(';');
                    for (int i = 0; i < linNeg.Count(); i++)
                    {
                        if (firstOnly)
                        {
                            strSql = strSql + " and TIPONEGOCIO in ('";
                            strSql = strSql + linNeg[i] + "'";
                            firstOnly = false;
                        }
                        else
                        {
                            strSql = strSql + ", '" + linNeg[i] + "'";
                        }
                    }

                    firstOnly = true;
                    strSql = strSql + ")";
                }

                // EMPRESAS

                if (this.empresas != "")
                {
                    firstOnly = true;
                    listEmpresas = this.empresas.Split(';');
                    for (int i = 0; i < this.empresas.Count(); i++)
                    {
                        if (firstOnly)
                        {
                            strSql = strSql + " and CEMPTITU in (";
                            strSql = strSql + this.empresas[i];
                            firstOnly = false;
                        }
                        else
                        {
                            strSql = strSql + ", " + this.empresas[i];
                        }
                    }

                    firstOnly = true;
                    strSql = strSql + ")";
                }

                // TIPOS FACTURAS

                if (this.tipoFactura != "")
                {
                    firstOnly = true;
                    for (int i = 0; i < this.tipoFactura.Count(); i++)
                    {
                        if (firstOnly)
                        {
                            strSql = strSql + " and tf.descripcion in ('";
                            strSql = strSql + this.tipoFactura[i] + "'";
                            firstOnly = false;
                        }
                        else
                        {
                            strSql = strSql + ", '" + this.tipoFactura[i] + "'";
                        }
                    }

                    firstOnly = true;
                    strSql = strSql + ")";
                }


                if (this.cnifdnic != "")
                {
                    strSql = strSql + " and cnifdnic = '" + this.cnifdnic + "'";
                }

                if (this.cupsREE != "")
                {
                    strSql = strSql + " and CUPSREE = '" + this.cupsREE + "'";
                }

                // ****************************************************************
                // *********** TRATAMIENTO DE LOS CONCEPTOS ***********************
                // ****************************************************************

                if (this.conceptos != "")
                {
                    listConceptos = this.conceptos.Split(';');
                    strSql = strSql + " and (";
                    for (int i = 1; i <= 9; i++)
                    {
                        firstOnly2 = true;
                        if (!firstOnly)
                        {
                            strSql = strSql + " or";
                        }
                        for (int x = 0; x < listConceptos.Count(); x++)
                        {

                            if (firstOnly2)
                            {
                                strSql = strSql + " TCONFAC" + i + " in ("
                                + listConceptos[x];
                                firstOnly2 = false;
                            }
                            else
                            {
                                strSql = strSql + ", " + listConceptos[x];
                            }

                        } // for(int x = 0;x< listConceptos.Count();x++)

                        strSql = strSql + ")";
                        firstOnly = false;
                    } // for(int i = 1; i<= 9; i++)

                    for (int i = 10; i <= 20; i++)
                    {
                        firstOnly2 = true;
                        if (!firstOnly)
                        {
                            strSql = strSql + " or";
                        }
                        for (int x = 0; x < listConceptos.Count(); x++)
                        {

                            if (firstOnly2)
                            {
                                strSql = strSql + " TCONFA" + i + " in ("
                                + listConceptos[x];
                                firstOnly2 = false;
                            }
                            else
                            {
                                strSql = strSql + ", " + listConceptos[x];
                            }

                        } // foreach(DataGridViewRow r in this.dgvConceptos.Rows)
                        strSql = strSql + ")";
                        firstOnly = false;
                    } // for (int i = 10; i <= 20; i++)
                    strSql = strSql + ")";
                } // if (this.HayConceptosSeleccionados())

                strSql = strSql + " and round(DATEDIFF(f.FFACTURA, f.FFACTHAS), 0) > 0"
                    + " group by f.CNIFDNIC, f.CREFEREN, f.FFACTDES, f.CUPSREE"
                    + " order by f.CNIFDNIC, f.FFACTDES";
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message,
                "Error en la construcción de la consulta",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

            return strSql;
        }

        private String MesLetra(int mes)
        {
            String _mes = "";

            switch (mes)
            {
                case 1:
                    _mes = "Enero";
                    break;
                case 2:
                    _mes = "Febrero";
                    break;
                case 3:
                    _mes = "Marzo";
                    break;
                case 4:
                    _mes = "Abril";
                    break;
                case 5:
                    _mes = "Mayo";
                    break;
                case 6:
                    _mes = "Junio";
                    break;
                case 7:
                    _mes = "Julio";
                    break;
                case 8:
                    _mes = "Agosto";
                    break;
                case 9:
                    _mes = "Septiempre";
                    break;
                case 10:
                    _mes = "Octubre";
                    break;
                case 11:
                    _mes = "Noviembre";
                    break;
                case 12:
                    _mes = "Diciembre";
                    break;
            }

            return _mes;
        }
    }
}
