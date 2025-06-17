using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    public class NRIs
    {

        logs.Log ficheroLog;
        public int[,] resumenNRI;
        public Dictionary<int, EndesaEntity.facturacion.cuadroDeMando.NRI> dic;
        public NRIs()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Informe_CuadroDeMando");
            resumenNRI = new int[4,5];
            dic = new Dictionary<int, EndesaEntity.facturacion.cuadroDeMando.NRI>();
        }

        public void Copia_NRI()
        {

            string strSql = "";
            int numreg = 0;
            bool firstOnly = true;
            int j = 0;

            string plazo = "";
            int dias = 0;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;

            StringBuilder sb = new StringBuilder();

            try
            {
                ficheroLog.Add("Copia_NRI");
                Console.WriteLine("Copiando NRI´s. ");

                #region Query
                strSql = "Select * from cm_nri where"
                    + " F_ULT_MOD = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                #endregion

                #region Parte Oracle
                if (!reader.Read())
                {
                    db.CloseConnection();
                    strSql = "SELECT DISTINCT h.FECHA,"
                        + "(SELECT e.ESTADO FROM estados e WHERE e.CODIGO_ESTADO = n.CODIGO_ESTADO) AS ESTADO,"
                        + " n.CODIGO_NRI,"
                        + " trim(n.CIF) cif,"
                        + " n.CLIENTE,"
                        + " l.ln,"
                        + " tm.tipo,"
                        + " m.MOTIVO"
                        + " FROM nri_nri n"
                        + " INNER JOIN nri_historico h"
                        + " ON n.CODIGO_NRI = h.CODIGO_NRI"
                        + " AND n.CODIGO_ESTADO = h.CODIGO_ESTADO"
                        + " INNER JOIN LN l"
                        + " on n.CODIGO_LN = l.CODIGO_LN"
                        + " INNER JOIN TIPOS_MOTIVO tm"
                        + " on n.CODIGO_TIPO = tm.codigo_tipo"
                        + " INNER JOIN MOTIVOS m"
                        + " on n.CODIGO_CAUSA = m.CODIGO_MOTIVO"
                        + " INNER JOIN SUBDIRECCION_CICLO sc on sc.CODIGO_SUB_CICLO = n.CODIGO_SUB_CICLO"
                        + " WHERE n.FECHA_CIERRE Is Null and sc.SUB_CICLO = 'OPERACIONES'";

                    ora_db = new OracleServer(OracleServer.Servidores.OWE);
                    ora_command = new OracleCommand(strSql, ora_db.con);
                    Console.WriteLine("Consultando BBDD ...");
                    r = ora_command.ExecuteReader();
                    Console.WriteLine("Copiando datos ...");

                    while (r.Read())
                    {

                        dias = Convert.ToInt32(DateTime.Now.Date.Subtract(Convert.ToDateTime(r["fecha"])).TotalDays);

                        j++;
                        numreg++;

                        if (firstOnly)
                        {
                            sb.Append("REPLACE into cm_nri (COD_NRI,CNIFDNIC,DAPERSOC,ESTADO,FECHA_ULTIMO_ESTADO,PLAZO,MOTIVO,SUBTIPO,LN,F_ULT_MOD) values ");
                            firstOnly = false;
                        }

                        sb.Append("(").Append(r["CODIGO_NRI"].ToString()).Append(", ");
                        sb.Append("'").Append(r["CIF"].ToString().Trim()).Append(" ', ");
                        sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(r["Cliente"].ToString())).Append("', ");
                        sb.Append("'").Append(r["Estado"].ToString()).Append("', ");
                        sb.Append("'").Append(Convert.ToDateTime(r["fecha"]).ToString("yyyy-MM-dd")).Append("', ");

                        if (dias > 60)
                            plazo = "Más de 2 meses";
                        else if (dias > 30)
                            plazo = "Entre Mes 1 y 2";
                        else if (dias > 15)
                            plazo = "Entre 15 y 30 Días";
                        else
                            plazo = "0-15 Días";

                        sb.Append("'").Append(plazo).Append("', ");
                        sb.Append("'").Append(r["tipo"].ToString()).Append("', ");
                        sb.Append("'").Append(r["MOTIVO"]).Append("', ");
                        sb.Append("'").Append(r["Ln"].ToString()).Append("', ");
                        sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd")).Append("'),");


                        if (numreg == 250)
                        {
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                            numreg = 0;
                        }
                    }

                    if (numreg > 0)
                    {

                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        numreg = 0;
                    }

                    Console.Write(j + " lineas.");

                    #endregion

                }

                db.CloseConnection();

                //this.ActualizaFechaProceso_OK("NRI", DateTime.Now);


            }
            catch (Exception e)
            {
                ficheroLog.AddError("Copia_NRI --> " + e.Message);
            }
        }

        public void CargaMatrizNRI()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            int fila = 0;
            int columna = 0;

            try
            {

                strSql = "select estado, plazo, count(*) total from cm_nri where"
                    + " F_ULT_MOD = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by estado,plazo";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    switch (r["plazo"].ToString())
                    {
                        case "0 - 15 Días":
                            columna = 0;
                            break;
                        case "Entre 15 y 30 Días":
                            columna = 1;
                            break;
                        case "Entre Mes 1 y 2":
                            columna = 2;
                            break;
                        default:
                            columna = 3;
                            break;
                    }

                    switch (r["estado"].ToString())
                    {
                        case "PENDIENTE DE RECIBIR CONTESTACION":
                            fila = 0;
                            break;                            
                        case "EN TRÁMITE":
                            fila = 1;
                            break;                
                        default:
                            fila = 2;
                            break;
                    }

                    resumenNRI[fila, columna] = resumenNRI[fila, columna] + Convert.ToInt32(r["total"]);
                    resumenNRI[fila, 4] = resumenNRI[fila, 4] + Convert.ToInt32(r["total"]);

                }
                db.CloseConnection();

                //Aplicamos totales por columna
                for (int f = 0; f <= 2; f++)
                    for (int c = 0; c <= 4; c++)
                        resumenNRI[3, c] = resumenNRI[3, c] + resumenNRI[f, c];


                strSql = "SELECT COD_NRI, CNIFDNIC, DAPERSOC, ESTADO, FECHA_ULTIMO_ESTADO, PLAZO,"
                    + " MOTIVO, SUBTIPO, LN, F_ULT_MOD"
                    + " FROM fact.cm_nri WHERE"
                    + " F_ULT_MOD = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.cuadroDeMando.NRI c = new EndesaEntity.facturacion.cuadroDeMando.NRI();

                    c.codigo_nri = Convert.ToInt32(r["COD_NRI"]);
                    c.nif = r["CNIFDNIC"].ToString();
                    c.cliente = r["DAPERSOC"].ToString();
                    c.estado = r["ESTADO"].ToString();

                    if (r["FECHA_ULTIMO_ESTADO"] != System.DBNull.Value)
                        c.fecha_ultimo_estado = Convert.ToDateTime(r["FECHA_ULTIMO_ESTADO"]);

                    c.plazo = r["PLAZO"].ToString();
                    c.motivo_alta = r["MOTIVO"].ToString();
                    c.submotivo_alta = r["SUBTIPO"].ToString();
                    c.linea_negocio = r["LN"].ToString();

                    EndesaEntity.facturacion.cuadroDeMando.NRI o;
                    if (!dic.TryGetValue(c.codigo_nri, out o))
                        dic.Add(c.codigo_nri, c);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

    }
}
