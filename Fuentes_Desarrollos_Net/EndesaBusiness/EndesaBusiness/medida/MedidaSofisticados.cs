using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class MedidaSofisticados
    {

        utilidades.Param param;
        public MedidaSofisticados()
        {
            param = new utilidades.Param("cm_param", MySQLDB.Esquemas.FAC);
        }

        public void CopiaCurvas()
        {
            int i = 0;
            int x = 0;
            string strSql = "";
            servidores.AccessDB ac;
            OleDbDataReader r;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            bool firstOnlyHora = true;
            DateTime fecha = new DateTime();

            MySQLDB db;
            MySqlCommand command;
            int total_registros = 0;
            int estacion = 0;

            EndesaBusiness.utilidades.Fechas utilFechas = new utilidades.Fechas();

            try
            {

                //strSql = "delete from cont.eer_curvas_cuarto_horarias_tmp";

                //db = new MySQLDB(MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //string ruta_access = @"E:\OneDrive - Atos\Endesa\peticiones_en_curso\20200618_MED_MERA_ERCROS_2018\MEDIDAS Y CURVAS TOP_20161103.accdb";
                //string tabla = "[HISTORICO LISTADO CURVAS_HASTA 201606]";
                string ruta_access = @"E:\OneDrive - Atos\Endesa\peticiones_en_curso\20200618_MED_MERA_ERCROS_2018\MEDIDAS Y CURVAS TOP con Historicos hasta 20181231.accdb";
                //ruta_access = @"E:\OneDrive - Atos\Endesa\peticiones_en_curso\20200618_MED_MERA_ERCROS_2018\MEDIDAS Y CURVAS TOP_2019.accdb";
                string tabla = "[HISTORICO LISTADO CURVAS_HASTA 201612]";
                tabla = "[HISTORICO LISTADO CURVAS DE 201701_A_201806]";
                tabla = "[HISTORICO LISTADO CURVAS]";

                strSql = "select count(*) as total_registros from " + tabla
                    + " where FECHA = #10/28/2018#";
                // ac = new servidores.AccessDB(param.GetValue("ruta_access_curvas", DateTime.Now, DateTime.Now));
                ac = new servidores.AccessDB(ruta_access);
                OleDbCommand cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                    total_registros = Convert.ToInt32(r["total_registros"]);

                ac.CloseConnection();

                Console.WriteLine("Total registros encontrados en: "
                    + ruta_access
                    + " --> " + string.Format("{0:#,##0}", total_registros));


                strSql = "select  CUPS, FECHA, HORA, ACTIVA, REACTIVA"
                     + " from " + tabla
                     + " where FECHA = #10/28/2018#";

                Console.WriteLine(strSql);

                ac = new servidores.AccessDB(ruta_access);
                cmd = new OleDbCommand(strSql, ac.con);
                r = cmd.ExecuteReader();
                while (r.Read())
                {


                    if (firstOnly)
                    {
                        sb.Append("replace into ag_medida");
                        sb.Append(" (CUPS, FECHA, HORA, ESTACION, ");
                        sb.Append("ACTIVA, REACTIVA) values ");
                        firstOnly = false;
                    }

                    if (r["CUPS"] != System.DBNull.Value &&
                        r["FECHA"] != System.DBNull.Value &&
                        r["HORA"] != System.DBNull.Value &&                        
                        r["ACTIVA"] != System.DBNull.Value &&
                        r["REACTIVA"] != System.DBNull.Value)
                    {

                        i++;
                        x++;

                        sb.Append("('").Append(r["CUPS"].ToString()).Append("',");
                                                
                        sb.Append("'").Append(Convert.ToDateTime(r["fecha"]).ToString("yyyy-MM-dd")).Append("',");
                        sb.Append(Convert.ToInt32(r["hora"])).Append(",");

                        fecha = Convert.ToDateTime(r["fecha"]).AddHours(Convert.ToInt32(r["hora"]));

                        if(fecha.Date == utilFechas.UltimoDomingoOctubre(fecha))
                        {
                            if(Convert.ToInt32(r["hora"]) == 2)
                            {
                                if (firstOnlyHora)
                                {
                                    estacion = utilFechas.Estacion(fecha, false);
                                    firstOnlyHora = false;
                                }else
                                    estacion = utilFechas.Estacion(fecha, true);

                            }
                            else
                            {
                                firstOnlyHora = true;
                                estacion = utilFechas.Estacion(fecha, false);
                            }
                                

                        }
                        else                        
                            estacion = utilFechas.Estacion(fecha, false);



                        sb.Append(estacion).Append(",");

                        if (r["ACTIVA"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["ACTIVA"]).ToString().Replace(",", ".")).Append(",");
                        else
                            sb.Append("null,");


                        if (r["REACTIVA"] != System.DBNull.Value)
                            sb.Append(Convert.ToDouble(r["REACTIVA"]).ToString().Replace(",", "."));
                        else
                            sb.Append("null");
                                               


                        sb.Append("),");

                        if (i == 500)
                        {
                            Console.CursorLeft = 0;
                            Console.Write(string.Format("{0:#,##0}", x) + " de " + string.Format("{0:#,##0}", total_registros));

                            i = 0;
                            firstOnly = true;
                            db = new MySQLDB(MySQLDB.Esquemas.FAC);
                            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                            command.ExecuteNonQuery();
                            db.CloseConnection();
                            sb = null;
                            sb = new StringBuilder();
                        }
                    }
                }

                if (i > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write(x + " de " + total_registros);
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                }

                ac.CloseConnection();

                //MessageBox.Show("Copia completada con éxito", "Copia de Curva Access", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                // MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


      

    }
}
