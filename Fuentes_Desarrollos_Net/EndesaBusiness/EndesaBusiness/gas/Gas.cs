using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.gas
{
    class Gas
    {
       
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Gas");
        public Dictionary<int, EndesaEntity.facturacion.DetalleCuadroMando> ld { get; set; }
        private Dictionary<int, string> lista_comentario_cisterna;
        private Dictionary<int, bool> lista_top_gas;
        private Dictionary<int, double> lista_tam;
        private utilidades.Param param;
        public Gas(DateTime fd, DateTime fh)
        {
            ld = new Dictionary<int, EndesaEntity.facturacion.DetalleCuadroMando>();
            lista_comentario_cisterna = new Dictionary<int, string>();
            lista_top_gas = new Dictionary<int, bool>();
            lista_tam = new Dictionary<int, double>();

            param = new utilidades.Param("cm_param", MySQLDB.Esquemas.FAC);
            Console.WriteLine("Gas.CargaTopGas");
            CargaTopGas(fd, fh);
            Console.WriteLine("Gas.CargaComentariosCisternas");
            CargaComentariosCisternas();
            Console.WriteLine("Gas.CargaGas2000MedidasYFacturas");
            CargaGas2000MedidasYFacturas();
            Console.WriteLine("Gas.GetPuntosMedida");
            GetPuntosMedida(fd, fh);
            Console.WriteLine("Gas.Inventario");
            Inventario(fd, fh);
            Console.WriteLine("Gas.ComprobarFacturacion");
            ComprobarFacturacion(fd, fh);
            Console.WriteLine("Gas.CargaTAMGas");
            CargaTAMGas();
            Console.WriteLine("Gas.SetGrupoInforme");
            SetGrupoInforme();

        }



        private void CargaGas2000MedidasYFacturas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {
                strSql = "select a.`IdPto Medida` as ID, a.Mes, a.Comentario, a.Facturacion"
                    + " from cm_medidas_facturas a inner join"
                    + " (select b.`IdPto Medida`, max(Mes) Mes from cm_medidas_facturas b group by b.`IdPto Medida`) as b"
                    + " on a.`IdPto Medida` = b.`IdPto Medida` and"
                    + " a.mes = b.mes"
                    + " where a.Mes > date_format(last_day((now() - interval 2 year)),'%Y%m');";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.DetalleCuadroMando d = new EndesaEntity.facturacion.DetalleCuadroMando();
                    d.idps = Convert.ToInt32(r["ID"]);
                    if (r["Mes"] != System.DBNull.Value)
                        d.fecha_medida = Convert.ToInt32(r["Mes"]);
                    if (r["Comentario"] != System.DBNull.Value)
                        d.comentario = r["Comentario"].ToString();
                    if (r["Facturacion"] != System.DBNull.Value)
                        d.facturado = r["Facturacion"].ToString() == "OK";
                    EndesaEntity.facturacion.DetalleCuadroMando cc;
                    if (!ld.TryGetValue(d.idps, out cc))
                        ld.Add(d.idps, d);
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Gas.CargaGas2000MedidasyFacturas " + e.Message);
            }
        }

        private void Inventario(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {
                strSql = "SELECT ID_PS, ID_CTO_PS, DAPERSOC, CNIFDNIC, FINICIO, FFIN, ID_ESTADO_CTO, CUPSREE, CuentaSCE,"
                    + " Municipio, Provincia, CD_PAIS, Distribuidora, RedDistribucion, Grupo"
                    + " FROM cm_inventario_gas g WHERE "
                    + " g.FINICIO <= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                    + " (g.FFIN >= '" + fh.ToString("yyyy-MM-dd") + "' OR g.FFIN IS NULL) and "
                    + " g.ID_ESTADO_CTO in (3,6,7,8,9,10,15,16)";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.DetalleCuadroMando d = new EndesaEntity.facturacion.DetalleCuadroMando();
                    d.existeMedidasyFacturas = false;
                    d.existeMedidasyFacturasOK = false;
                    d.existe_T_SGM_M_LECTURASyCONSUMOS = false;

                    d.idps = Convert.ToInt32(r["ID_PS"]);

                    if (r["CD_PAIS"] != System.DBNull.Value)
                        d.cd_pais = r["CD_PAIS"].ToString();
                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        d.nif = r["CNIFDNIC"].ToString();
                    if (r["DAPERSOC"] != System.DBNull.Value)
                        d.nombre_cliente = r["DAPERSOC"].ToString();
                    if (r["CUPSREE"] != System.DBNull.Value)
                        d.cups20 = r["CUPSREE"].ToString();
                    if (r["grupo"] != System.DBNull.Value)
                        d.grupo = r["Grupo"].ToString();
                    if (r["FINICIO"] != System.DBNull.Value)
                        d.fd = Convert.ToDateTime(r["FINICIO"]);
                    if (r["FFIN"] != System.DBNull.Value)
                        d.fh = Convert.ToDateTime(r["FFIN"]);


                    EndesaEntity.facturacion.DetalleCuadroMando cc;
                    if (!ld.TryGetValue(d.idps, out cc))
                        ld.Add(d.idps, d);
                    else
                    {
                        cc.cd_pais = d.cd_pais;
                        cc.nif = d.nif;
                        cc.nombre_cliente = d.nombre_cliente;
                        cc.cups20 = d.cups20;
                        cc.grupo = d.grupo;
                        cc.fd = d.fd;
                        cc.fh = d.fh;
                        ld[d.idps] = cc;
                    }


                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Gas.Inventario: " + e.Message);
            }
        }

        private void GetPromedioFacturacion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            int idps = 0;

            try
            {
                strSql = "select g.ID_PS, g.tam from tam_gas g";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    idps = Convert.ToInt32(r["id_ps"]);

                    if (r["tam"] != System.DBNull.Value)
                    {
                        EndesaEntity.facturacion.DetalleCuadroMando e;

                        if (ld.TryGetValue(idps, out e))
                        {
                            e.tam = Math.Round(Convert.ToDouble(r["tam"]), 2);
                            ld[idps] = e;
                        }
                    }
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Gas.GetPromedioFacturacion: " + e.Message);
            }
        }

        private void ComprobarFacturacion(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string nfactura = "";
            int idps = 0;

            try
            {
                Console.WriteLine("Comprobando facturacion de Gas.");
                #region fo_s
                strSql = "Select ID_PS, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS, max(FH_FACTURA) as FH_FACTURA,"
                    + " max(ID_FACTURA) as ID_FACTURA"
                    + " from fo_s f where"
                    + " FH_INI_FACTURACION >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FIN_FACTURACION <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA >= '" + DateTime.Now.ToString("yyyy-MM-01") + "' and"
                    + " CD_NFACTURA_REALES_PS is not null and"
                    + " instr(CD_NFACTURA_REALES_PS, 'S') < 1"
                    + " group by ID_PS";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    idps = Convert.ToInt32(r["ID_PS"]);
                    EndesaEntity.facturacion.DetalleCuadroMando e;
                    if (ld.TryGetValue(idps, out e))
                    {
                        e.ultimo_mes_facturado = Convert.ToInt32(Convert.ToDateTime(r["FH_INI_FACTURACION"]).ToString("yyyyMM"));

                        if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        {
                            nfactura = r["CD_NFACTURA_REALES_PS"].ToString();
                            e.facturado = nfactura.Contains("Y") || nfactura.Contains("N");
                        }

                        ld[idps] = e;
                    }
                }
                db.CloseConnection();
                #endregion
                #region fo_s_sce
                strSql = "Select ID_PS, FH_INI_FACTURACION, FH_FIN_FACTURACION, CD_NFACTURA_REALES_PS, max(FH_FACTURA) as FH_FACTURA"
                    + " from fo_s_sce f where"
                    + " FH_INI_FACTURACION >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FIN_FACTURACION <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                    + " FH_FACTURA >= '" + DateTime.Now.ToString("yyyy-MM-01") + "' and"
                    + " CD_NFACTURA_REALES_PS is not null and"
                    + " instr(CD_NFACTURA_REALES_PS, 'S') < 1"
                    + " group by ID_PS";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    idps = Convert.ToInt32(r["ID_PS"]);
                    EndesaEntity.facturacion.DetalleCuadroMando e;
                    if (ld.TryGetValue(idps, out e))
                    {
                        e.ultimo_mes_facturado = Convert.ToInt32(Convert.ToDateTime(r["FH_INI_FACTURACION"]).ToString("yyyyMM"));
                        e.fecha_medida = Convert.ToInt32(Convert.ToDateTime(r["FH_INI_FACTURACION"]).ToString("yyyyMM"));

                        if (r["CD_NFACTURA_REALES_PS"] != System.DBNull.Value)
                        {
                            nfactura = r["CD_NFACTURA_REALES_PS"].ToString();
                            e.facturado = nfactura.Contains("Y") || nfactura.Contains("N");

                        }

                        ld[idps] = e;
                    }
                    else
                    {
                        e.ultimo_mes_facturado = Convert.ToInt32(Convert.ToDateTime(r["FH_INI_FACTURACION"]).ToString("yyyyMM"));
                        e.fecha_medida = Convert.ToInt32(Convert.ToDateTime(r["FH_INI_FACTURACION"]).ToString("yyyyMM"));
                        e = new EndesaEntity.facturacion.DetalleCuadroMando();
                        e.idps = Convert.ToInt32(r[" ID_PS"]);

                        ld.Add(e.idps, e);
                    }
                }
                db.CloseConnection();
                #endregion
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Gas.ComprobarFacturacion: " + e.Message);
            }
        }

        private void GetPuntosMedida(DateTime fd, DateTime fh)
        {
            string strSql;
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            bool firstOnly = true;
            int idps = 0;
            int id_pmedida = 0;

            try
            {
                strSql = "Select ID_PS,ID_PMEDIDA from T_SGM_M_PUNTOS_MEDIDA WHERE"
                    + " ID_PS IN (";
                foreach (KeyValuePair<int, EndesaEntity.facturacion.DetalleCuadroMando> d in ld)
                {
                    if (firstOnly)
                    {
                        strSql += d.Key;
                        firstOnly = false;
                    }
                    else
                    {
                        strSql += "," + d.Key;
                    }
                }


                strSql += ") AND "
                    + "FH_INICIO <= '" + fd.ToString("yyyy/MM/dd") + "' AND "
                    + "(FH_FIN >= '" + fh.ToString("yyyy/MM/dd") + "' OR "
                    + "FH_FIN is null)";
                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    idps = Convert.ToInt32(r["id_ps"]);
                    if (r["ID_PMEDIDA"] != System.DBNull.Value)
                        id_pmedida = Convert.ToInt32(r["ID_PMEDIDA"]);
                    else
                        id_pmedida = 0;

                    EndesaEntity.facturacion.DetalleCuadroMando e;
                    if (ld.TryGetValue(idps, out e))
                    {
                        ld[idps].id_pmedida = id_pmedida;
                    }
                    else
                    {
                        e = new EndesaEntity.facturacion.DetalleCuadroMando();
                        e.idps = idps;
                        e.id_pmedida = id_pmedida;
                        ld.Add(idps, e);
                    }
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Gas.GetPuntosMedida: " + e.Message);
            }
        }

        private void SetGrupoInforme()
        {
            int anio;
            int mes;
            int dias_del_mes;

            anio = DateTime.Now.Year;
            mes = (DateTime.Now.Month) - 1;
            dias_del_mes = DateTime.DaysInMonth(anio, mes);

            DateTime fhTrabajo = new DateTime(anio, mes, dias_del_mes);


            foreach (KeyValuePair<int, EndesaEntity.facturacion.DetalleCuadroMando> d in ld)
            {
                bool b;
                d.Value.esTop = lista_top_gas.TryGetValue(d.Key, out b);
                if (d.Value.ultimo_mes_facturado == Convert.ToInt32(fhTrabajo.ToString("yyyyMM")))
                    d.Value.estado_ltp = "FACTURADO";
                else if (d.Value.fecha_medida == Convert.ToInt32(fhTrabajo.ToString("yyyyMM")))
                    d.Value.estado_ltp = "Pend. Facturación";
                else
                    d.Value.estado_ltp = "Pdte. Medida";


                if (d.Value.esTop && d.Value.cd_pais == "ESP")
                    ld[d.Key].grupo_informe = "TOP GAS ESPAÑA";
                else if (d.Value.cd_pais == "POR")
                    ld[d.Key].grupo_informe = "GAS PORTUGAL";
                else if (d.Value.grupo == "Cisterna")
                {
                    ld[d.Key].grupo_informe = "CISTERNAS GAS ESPAÑA";
                    ld[d.Key].comentario = GetComentarioCisterna(d.Value.idps);
                }
                else
                    ld[d.Key].grupo_informe = "NO TOP GAS ESPAÑA";
            }


        }
        private string GetComentarioCisterna(int idps)
        {
            string comentario = "";
            if (lista_comentario_cisterna.TryGetValue(idps, out comentario))
                return comentario;
            else
                return "";
        }

        private void CargaComentariosCisternas()
        {
            int f = 0;

            try
            {
                string fichero = param.GetValue("ExcelComentariosCisternas", DateTime.Now, DateTime.Now);
                FileStream fs = new FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();
                for (int i = 2; i < 1000000; i++)
                {
                    f = i;


                    if (workSheet.Cells[f, 2].Value == null
                        || workSheet.Cells[f, 3].Value == null
                        || workSheet.Cells[f, 4].Value == null)
                    {
                        break;
                    }
                    else
                    {
                        lista_comentario_cisterna.Add(Convert.ToInt32(workSheet.Cells[f, 2].Value), workSheet.Cells[f, 5].Value.ToString());
                    }
                }


                excelPackage = null;

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Gas.CargaComentariosCisternas: " + e.Message);
            }
        }

        private void CargaTopGas(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int maxMes = 0;

            try
            {

                strSql = "select max(mes) MaxMes from cm_top_gas";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    maxMes = Convert.ToInt32(r["MaxMes"]);
                }
                db.CloseConnection();

                strSql = "select IdPtoSuministro from cm_top_gas where"
                    + " MES = " + (Convert.ToInt32(fd.ToString("yyyyMM")) > maxMes ? maxMes : Convert.ToInt32(fd.ToString("yyyyMM")));

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    lista_top_gas.Add(Convert.ToInt32(r["IdPtoSuministro"]), true);

                }
                db.CloseConnection();
            }

            catch (Exception e)
            {
                ficheroLog.AddError("Gas.CargaTopGas: " + e.Message);
            }
        }
        private void CargaTAMGas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;


            try
            {

                strSql = "select g.ID_PS, g.tam from tam_gas g";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    lista_tam.Add(Convert.ToInt32(r["ID_PS"]), Convert.ToDouble(r["tam"]));
                }
                db.CloseConnection();
            }

            catch (Exception e)
            {
                ficheroLog.AddError("Gas.CargaTAMGas: " + e.Message);
            }
        }
    }
}
