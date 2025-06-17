using EndesaBusiness.servidores;
using Microsoft.IdentityModel.Tokens;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types.InlineQueryResults;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    public class Resumen : EndesaEntity.facturacion.cuadroDeMando.Resumen
    {
        Dictionary<string, List<EndesaEntity.facturacion.cuadroDeMando.Resumen>> dic;
        logs.Log ficheroLog;
        public Resumen()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_Informe_CuadroDeMando");
        }

        public Resumen(int aniomes)
        {
            dic = Carga(aniomes);
            if (dic.Count == 0)
            {
                CreaResumen(aniomes);
                dic = Carga(aniomes);
            }
                
        }

        private Dictionary<string, List<EndesaEntity.facturacion.cuadroDeMando.Resumen>> Carga(int aniomes)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            string clave = "";

            Dictionary<string, List<EndesaEntity.facturacion.cuadroDeMando.Resumen>> d
                = new Dictionary<string, List<EndesaEntity.facturacion.cuadroDeMando.Resumen>>();

            try
            {
                strSql = "select aniomes, pais, segmento, grupo, id_concepto, concepto";
                for (int i = 1; i <= 31; i++)
                    strSql += ", d" + i;

                strSql += " from ccmm_resumen"
                    + " where aniomes = " + aniomes;
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.facturacion.cuadroDeMando.Resumen c = new EndesaEntity.facturacion.cuadroDeMando.Resumen();
                    clave = r["pais"].ToString() + "_"
                        + r["segmento"].ToString() + "_"
                        + r["grupo"].ToString();

                    c.aniomes = Convert.ToInt32(r["aniomes"]);
                    c.pais = r["pais"].ToString();
                    c.segmento = r["segmento"].ToString();
                    c.grupo = r["grupo"].ToString();
                    c.id_concepto = Convert.ToInt32(r["aniomes"]);

                    for (int i = 1; i <= 31; i++)
                    {
                        if (r["d" + i] != System.DBNull.Value)
                            c.dias[i] = Convert.ToInt32(r["d" + i]);
                    }


                    List<EndesaEntity.facturacion.cuadroDeMando.Resumen> o;
                    if(!d.TryGetValue(clave, out o))
                    {
                        o = new List<EndesaEntity.facturacion.cuadroDeMando.Resumen>();
                        o.Add(c);
                        d.Add(clave, o);
                    }
                    else
                        o.Add(c);
                }
                db.CloseConnection();

                return d;

            }catch(Exception ex)
            {
                ficheroLog.AddError("Resumen.Carga " + ex.Message);
                return null;
            }
        }

        public List<EndesaEntity.facturacion.cuadroDeMando.Resumen> GetValor(string pais, string segmento, string grupo)
        {
            string clave = "";
            
            clave = pais + "_" + segmento + "_" + grupo;

            List<EndesaEntity.facturacion.cuadroDeMando.Resumen> o;
            if (dic.TryGetValue(clave, out o))
                return o;
            else
                return null;
        }

        public void Save_Electricidad(List<EndesaEntity.facturacion.cuadroDeMando.Informe> informe, 
            int aniomes, string pais, string segmento, string grupo, int dia)
        {
                   

            int pendiente_de_medida = 0;
            int cc_rechazada_por_cs = 0;
            int cc_enviada_a_sce_ml = 0;
            int cc_rechazada_por_sce_ml = 0;
            int cc_incompleta_sce_ml = 0;
            int ltp_sce = 0;
            int punto_no_esta_extraido = 0;
            int punto_esta_extraido = 0;
            int prefactura_pendiente = 0;
            int facturado = 0;
            int pendiente = 0;

            try
            {
                foreach(EndesaEntity.facturacion.cuadroDeMando.Informe p in informe)
                {
                    switch (p.estado_LTP)
                    {
                        case "01. Pendiente de medida":
                            pendiente_de_medida++;
                            pendiente++;
                            break;
                        case "02. CC Rechazada por CS":
                            cc_rechazada_por_cs++;
                            pendiente++;
                            break;
                        case "04. CC Enviada a SCE ML":
                            cc_enviada_a_sce_ml++;
                            pendiente++;
                            break;
                        case "05. CC Rechazada por SCE ML":
                            cc_rechazada_por_sce_ml++;
                            pendiente++;
                            break;
                        case "06. CC Incompleta  SCE ML":
                            cc_incompleta_sce_ml++;
                            pendiente++;
                            break;
                        case "07. LTP SCE":
                            ltp_sce++;
                            pendiente++;
                            break;
                        case "08. El Punto no está Extraído":
                            punto_no_esta_extraido++;
                            pendiente++;
                            break;
                        case "09. El Punto está Extraído":
                            punto_esta_extraido++;
                            pendiente++;
                            break;
                        case "10. Prefactura pendiente":
                            prefactura_pendiente++;
                            pendiente++;
                            break;
                        case "FACTURADO":
                            facturado++;                            
                            break;
                        default:
                            pendiente++;
                            break;
                    }
                }

                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 1, pendiente_de_medida);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 2, cc_rechazada_por_cs);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 3, cc_enviada_a_sce_ml);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 4, cc_rechazada_por_sce_ml);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 5, cc_incompleta_sce_ml);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 6, ltp_sce);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 7, punto_no_esta_extraido);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 8, punto_esta_extraido);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 9, prefactura_pendiente);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 10, facturado);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 11, pendiente);

                dic = null;
                dic = Carga(aniomes);

            }
            catch(Exception e)
            {
                ficheroLog.AddError("Resumen.Save " + e.Message);
            }
        }

        // Duplicamos funcion y modificamos para SAP, solo va a contener 7 conceptos en resumen en vez de 11 como los SCE:
        // 1 - Pendiente de medida
        // 2 - Orden de Cálculo Calculable
        // 3 - DC Generado sin DI
        // 4 - Doc. Impresión Apartado
        // 5 - DC no generado MDS
        // 10 - FACTURADO
        // 11 - PENDIENTE
        public void Save_Electricidad_SAP(List<EndesaEntity.facturacion.cuadroDeMando.Informe> informe,
           int aniomes, string pais, string segmento, string grupo, int dia)
        {


            int pendiente_de_medida = 0;        // Concepto 1 - 01. Pendiente medida
            int orden_calculo_calculable = 0;   // Concepto 2 - 02. Orden de Cálculo Calculable
            int dc_generado_sin_di = 0;         // Concepto 3 - 03. DC Generado sin DI
            int doc_impresion_apartado = 0;     // Concepto 4 - 04. Doc. Impresión Apartado
            int dc_no_generado_mds = 0;         // Concepto 5 - 05. DC no generado MDS
            int facturado = 0;                  // Concepto 10 - FACTURADO
            int pendiente = 0;                  // Concepto 11 - PENDIENTE

            try
            {
                foreach (EndesaEntity.facturacion.cuadroDeMando.Informe p in informe)
                {

                    if (p.estado_LTP.IsNullOrEmpty())
                    {
                        pendiente++;
                    }
                    else
                    {
                        switch ((p.estado_LTP).Split('.')[0])
                        {
                            case "01":
                                pendiente_de_medida++;
                                pendiente++;
                                break;
                            case "02":
                                orden_calculo_calculable++;
                                pendiente++;
                                break;
                            case "03":
                                dc_generado_sin_di++;
                                pendiente++;
                                break;
                            case "04":
                                doc_impresion_apartado++;
                                pendiente++;
                                break;
                            case "05":
                                dc_no_generado_mds++;
                                pendiente++;
                                break;
                            case "FACTURADO":
                                facturado++;
                                break;
                            default:
                                pendiente++;
                                break;
                        }
                    }
                }

                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 1, pendiente_de_medida);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 2, orden_calculo_calculable);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 3, dc_generado_sin_di);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 4, doc_impresion_apartado);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 5, dc_no_generado_mds);
                
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 10, facturado);
                Update_Save_Electricidad(aniomes, pais, segmento, grupo, dia, 11, pendiente);

                dic = null;
                dic = Carga(aniomes);

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Resumen.Save " + e.Message);
            }
        }

        private void Update_Save_Electricidad(int aniomes, string pais, 
            string segmento, string grupo, int dia, int id_concepto, int total)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "update ccmm_resumen set"
                        + " d" + dia + " = " + total
                        + " where aniomes = " + aniomes + " and"
                        + " pais = '" + pais + "' and"
                        + " segmento = '" + segmento + "' and"
                        + " grupo = '" + grupo + "' and"
                        + " id_concepto = " + id_concepto;

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void Save_GAS(List<EndesaEntity.facturacion.cuadroDeMando.InformeGas> informe,
            int aniomes, string pais, string segmento, string grupo, int dia)
        {
            

            int pendiente_medida = 0;
            int pendiente_facturacion = 0;

            int facturado = 0;
            int pendiente = 0;

            try
            {
                foreach (EndesaEntity.facturacion.cuadroDeMando.InformeGas p in informe)
                {
                    switch (p.pendiente)
                    {
                        case "Pend. Facturación":
                            pendiente_facturacion++;
                            pendiente++;
                            break;
                        case "Pdte. Medida":
                            pendiente_medida++;
                            pendiente++;
                            break;                        
                        case "":
                            facturado++;                            
                            break;
                        default:
                            pendiente++;
                            break;
                    }
                }

                Update_Save_GAS(aniomes, pais, segmento, grupo, dia, 1, pendiente_medida);
                Update_Save_GAS(aniomes, pais, segmento, grupo, dia, 2, pendiente_facturacion);
                Update_Save_GAS(aniomes, pais, segmento, grupo, dia, 3, facturado);
                Update_Save_GAS(aniomes, pais, segmento, grupo, dia, 4, pendiente);

                dic = null;
                dic = Carga(aniomes);

            }
            catch (Exception e)
            {
                ficheroLog.AddError("Resumen.Save " + e.Message);
            }
        }

        private void Update_Save_GAS(int aniomes, string pais,
            string segmento, string grupo, int dia, int id_concepto, int total)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            strSql = "update ccmm_resumen set"
                        + " d" + dia + " = " + total
                        + " where aniomes = " + aniomes + " and"
                        + " pais = '" + pais + "' and"
                        + " segmento = '" + segmento + "' and"
                        + " grupo = '" + grupo + "' and"
                        + " id_concepto = " + id_concepto;

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void CreaResumen(int aniomes)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "replace into ccmm_resumen"
                    + " SELECT " + aniomes + " ,r.pais, r.segmento, r.grupo, r.id_concepto, r.concepto";
                for (int i = 1; i <= 31; i++)
                    strSql += " ,null";

                strSql += ", '" + Environment.UserName + "'," 
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', r.last_update_by, r.last_update_date"
                    + " FROM ccmm_resumen_plantilla r";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch(Exception e)
            {
                ficheroLog.AddError("Resumen.CreaResumen " + e.Message);
            }


                

        }
    }
}

