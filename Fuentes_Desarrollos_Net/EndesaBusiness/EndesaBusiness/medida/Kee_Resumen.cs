using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class Kee_Resumen
    {

        public Kee_Resumen()
        {
            List<EndesaEntity.medida.Kee_Resumen_Tabla> lr = new List<EndesaEntity.medida.Kee_Resumen_Tabla>();
            List<EndesaEntity.medida.Kee_Resumen_Tabla> lr2 = new List<EndesaEntity.medida.Kee_Resumen_Tabla>();
            List<EndesaEntity.medida.CurvaResumen> lista_cr = new List<EndesaEntity.medida.CurvaResumen>();
            
            
            Borrar_Tabla("kee_resumen");
            Borrar_Tabla("kee_reporte_extraccion_ch_horizontal");

            Kee_Extraccion_Formulas kef = new Kee_Extraccion_Formulas(Kee_Extraccion_Formulas.Tipo.CH);
            Kee_Reporte_Extraccion_CH kre = new Kee_Reporte_Extraccion_CH();
            CRRD cr = new CRRD(kef.list.Select(z => z.cups20).ToList(), kef.MinDateValue(), kef.MaxDateValue(), "F");


            foreach(EndesaEntity.medida.Kee_Extraccion_Formulas p in kef.list)
            {
                lista_cr = cr.GetCurvasResumen(p.cups20, p.fecha_desde, p.fecha_hasta);
                if(lista_cr.Count() > 0)
                {
                    foreach (EndesaEntity.medida.CurvaResumen j in lista_cr)
                    {   
                       
                        EndesaEntity.medida.Kee_Resumen_Tabla c = new EndesaEntity.medida.Kee_Resumen_Tabla();
                        c.cups20 = p.cups20;
                        c.cups22 = j.cups22;
                        c.fd_extraccion_formulas = p.fecha_desde;
                        c.fh_extraccion_formulas = p.fecha_hasta;
                        c.fuente_extraccion_formulas = p.tipo;
                        c.fd_resumen_sce_ml = j.fd;
                        c.fh_resumen_sce_ml = j.fh;
                        c.activa_resumen = j.activa;
                        c.reactiva_resumen = j.reactiva;
                        c.existe_resumen_sce_ml = true;

                        kre.GetCurva(j.cups22, j.fd, j.fh, p.tipo);
                        if (kre.existe_curva)
                        {
                            c.existe_kee_ch = true;
                            c.activa_kre_ch = kre.total_activa;
                            c.reactiva_kre_ch = kre.total_reactiva;

                            c.dif_activa = c.activa_resumen - c.activa_kre_ch;
                            c.dif_reactiva = c.reactiva_resumen - c.reactiva_kre_ch;
                        }

                        lr.Add(c);
                    }
                }                
                else
                {
                    EndesaEntity.medida.Kee_Resumen_Tabla c = new EndesaEntity.medida.Kee_Resumen_Tabla();
                    c.cups20 = p.cups20;
                    c.fd_extraccion_formulas = p.fecha_desde;
                    c.fh_extraccion_formulas = p.fecha_hasta;
                    c.fuente_extraccion_formulas = p.tipo;
                    lr.Add(c);
                }
            }           
            Guardar_Kee_Resumen(lr);
        }

        private void Borrar_Tabla(string tabla)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
                        
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            strSql = "delete from " + tabla;
            Console.WriteLine(strSql);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void Guardar_Kee_Resumen(List<EndesaEntity.medida.Kee_Resumen_Tabla> list)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int registros = 0;
            int totalregistros = 0;
            try
            {
                foreach(EndesaEntity.medida.Kee_Resumen_Tabla p in list)
                {
                    registros++;
                    totalregistros++;
                    if (firstOnly)
                    {
                        sb.Append("replace into kee_resumen");
                        sb.Append(" (cups20, cups22, fd, fh, tipo, ffactdes, ffacthas, activa_resumen, reactiva_resumen, activa_kee, reactiva_kee, dif_activa, dif_reactiva) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.cups20).Append("',");
                    sb.Append("'").Append(p.cups22).Append("',");
                    sb.Append("'").Append(p.fd_extraccion_formulas.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(p.fh_extraccion_formulas.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(p.fuente_extraccion_formulas).Append("',");
                    
                    if (p.existe_resumen_sce_ml)
                    {
                        sb.Append("'").Append(p.fd_resumen_sce_ml.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append("'").Append(p.fh_resumen_sce_ml.ToString("yyyy-MM-dd")).Append("',");
                        sb.Append(p.activa_resumen.ToString().Replace(",", ".")).Append(",");
                        sb.Append(p.reactiva_resumen.ToString().Replace(",", ".")).Append(",");
                    }
                    else
                        sb.Append("null,null,null,null,");

                    if (p.existe_kee_ch)
                    {
                        sb.Append(p.activa_kre_ch.ToString().Replace(",", ".")).Append(",");
                        sb.Append(p.reactiva_kre_ch.ToString().Replace(",", ".")).Append(",");
                    }
                    else
                        sb.Append("null,null,");

                    if (p.existe_resumen_sce_ml && p.existe_kee_ch)
                    {
                        sb.Append(p.dif_activa.ToString().Replace(",", ".")).Append(",");
                        sb.Append(p.dif_reactiva.ToString().Replace(",", ".")).Append("),");
                    }
                    else
                        sb.Append("null,null),");

                    if (registros == 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write("Guardando " + totalregistros + " de " + list.Count);
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        registros = 0;
                    }
                }

                if (registros > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write("Guardando " + totalregistros + " de " + list.Count);
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    registros = 0;
                }

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }
    }

    
}
