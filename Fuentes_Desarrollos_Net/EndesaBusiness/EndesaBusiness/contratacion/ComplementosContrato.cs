using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class ComplementosContrato
    {

        Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> dic;
        Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> dic_historico;
        
        public ComplementosContrato(DateTime fd, DateTime fh, string complemento)
        {
            if(fd.Month == DateTime.Now.Month)
            {
                dic = CargaActual(complemento, fd, fh);
            }
            else
            {
                dic_historico = CargaHistorico(complemento);
                dic = Carga(complemento, fd, fh);
            }
            
        }

        private Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> CargaActual(string complemento, DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool add_record = false;
            int numreg = 0;


            Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> d
                = new Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>>();

            try
            {



                strSql = "SELECT ps.CCOUNIPS, if(r.CUPS22 is null, rr.CUPS20,"
                    + " substr(r.CUPS22,1,20)) as CUPS20, ps.FINICPLA, ps.CNUMSCPS,"
                    + " ps.TESTCONT, ps.FFINVESU , ps.FALTACON, ps.FBAJACON,"
                    + " if(ps.FPSERCON = '0000-00-00', null, ps.FPSERCON) as FPSERCON,"
                    + " ps.CCOMPOBJ"
                    + " FROM cont.contratosPS ps"
                    + " INNER JOIN cont.PS_AT r ON r.IDU = ps.ccounips"
                    + " left outer join aux1.RELACION_CUPS rr on"
                    + " rr.CUPS_CORTO = ps.ccounips"
                    + " WHERE ps.CCOMPOBJ = '" + complemento + "'"
                    + " and ps.TESTCONT IN (1,3,4,8)"
                    //+ " and ps.FPSERCON <> '0000-00-00'"
                    + " GROUP BY ps.CCOUNIPS";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    // 1-- > PDTE O / S
                    // 3-- > VIGOR
                    // 4-- > PDTE BAJA
                    // 8-- > REV CONTRATO

                    //numreg++;
                    //if (d.Count == 2186)
                    //    Console.WriteLine("a");

                    add_record = false;
                    EndesaEntity.contratacion.ContratosPS_Tabla c
                        = new EndesaEntity.contratacion.ContratosPS_Tabla();

                    c.tipo_estado_contrato = Convert.ToInt32(r["TESTCONT"]);
                    c.version_contrato = Convert.ToInt32(r["CNUMSCPS"]);
                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPS20"].ToString();

                    //if (c.cups20 == "ES0031405870000001WQ")
                    //    Console.WriteLine("");


                    if (r["FPSERCON"] != System.DBNull.Value)
                    {
                        if (r["FPSERCON"].ToString() != "0000-00-00")
                            c.fecha_inicio_version = Convert.ToDateTime(r["FPSERCON"]);
                        else
                            c.fecha_inicio_version = fd;
                    }
                    else
                        c.fecha_inicio_version = fd;
                    
                    List<EndesaEntity.contratacion.ContratosPS_Tabla> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.contratacion.ContratosPS_Tabla>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    else
                        o.Add(c);
                    

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }
        private Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> Carga(string complemento, DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";          
            bool add_record = false;
            int numreg = 0;


            Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> d
                = new Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>>();

            try
            {
               
                

                strSql = "SELECT ps.CCOUNIPS, if(r.CUPS22 is null, rr.CUPS20,"
                    + " substr(r.CUPS22,1,20)) as CUPS20, ps.FINICPLA, ps.CNUMSCPS,"
                    + " ps.TESTCONT, ps.FFINVESU , ps.FALTACON, ps.FBAJACON," 
                    + " if(ps.FPSERCON = '0000-00-00', null, ps.FPSERCON) as FPSERCON,"
                    + " ps.CCOMPOBJ"
                    + " FROM cont.contratosPS ps"
                    + " INNER JOIN cont.PS_AT r ON r.IDU = ps.ccounips"                    
                    + " left outer join aux1.RELACION_CUPS rr on"
                    + " rr.CUPS_CORTO = ps.ccounips"                    
                    + " WHERE ps.CCOMPOBJ = '" + complemento + "'"
                    +  " and ps.TESTCONT IN (1,3,4,8)"
                    //+ " and ps.FPSERCON <> '0000-00-00'"
                    + " GROUP BY ps.CCOUNIPS";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    // 1-- > PDTE O / S
                    // 3-- > VIGOR
                    // 4-- > PDTE BAJA
                    // 8-- > REV CONTRATO

                    //numreg++;
                    //if (d.Count == 2186)
                    //    Console.WriteLine("a");

                    add_record = false;
                    EndesaEntity.contratacion.ContratosPS_Tabla c
                        = new EndesaEntity.contratacion.ContratosPS_Tabla();

                    c.tipo_estado_contrato = Convert.ToInt32(r["TESTCONT"]);
                    c.version_contrato = Convert.ToInt32(r["CNUMSCPS"]);                    
                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPS20"].ToString();

                    


                    if (r["FPSERCON"] != System.DBNull.Value)
                    {
                        if(r["FPSERCON"].ToString() != "0000-00-00")
                            c.fecha_inicio_version = Convert.ToDateTime(r["FPSERCON"]);
                        else
                            c.fecha_inicio_version = fd;
                    }                        
                    else
                        c.fecha_inicio_version = fd;

                    switch (c.tipo_estado_contrato)
                    {
                        case 1:
                            if (c.version_contrato > 1 &&
                                ExisteContratoHistorico(c.cups20, (c.version_contrato - 1), fh))
                                add_record = true;
                            break;
                        case 3:
                            if (c.fecha_inicio_version <= fd)
                                add_record = true;
                            else if (ExisteContratoHistorico(c.cups20, (c.version_contrato - 1), fh))
                                add_record = true;
                            break;
                        case 4:
                            if (c.fecha_inicio_version <= fd)
                                add_record = true;
                            else if (ExisteContratoHistorico(c.cups20, (c.version_contrato - 1), fh))
                                add_record = true;
                            break;
                        case 8:
                            if (ExisteContratoHistorico(c.cups20, (c.version_contrato - 1), fh))
                                add_record = true;
                            break;

                    }

                    if (add_record)
                    {
                        List<EndesaEntity.contratacion.ContratosPS_Tabla> o;
                        if (!d.TryGetValue(c.cups20, out o))
                        {
                            o = new List<EndesaEntity.contratacion.ContratosPS_Tabla>();
                            o.Add(c);
                            d.Add(c.cups20, o);
                        }
                        else
                            o.Add(c);
                    }

                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        


        private Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> 
            CargaHistorico(string complemento)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";           

            Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> d
                = new Dictionary<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>>();

            try
            {

                strSql = "SELECT ps.CCOUNIPS, psat.CUPS22,"
                    + "ps.TESTCONT, ps.CNUMSCPS,"
                    + " if(ps.FPSERCON = '0000-00-00', null, ps.FPSERCON) as FPSERCON"
                    + " FROM cont.contratosPS_hist ps"
                    + " INNER JOIN cont.PS_AT psat ON psat.IDU = ps.ccounips"
                    + " WHERE ps.TESTCONT IN (1,3,4,8) AND ps.CCOMPOBJ = '" + complemento + "'"
                    //+ " and ps.FPSERCON <> '0000-00-00'"
                    + " GROUP BY ps.CCOUNIPS, ps.CNUMSCPS ORDER BY ps.CCOUNIPS, ps.CNUMSCPS DESC";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.ContratosPS_Tabla c
                        = new EndesaEntity.contratacion.ContratosPS_Tabla();

                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPS22"].ToString().Substring(0,20);
                    c.tipo_estado_contrato = Convert.ToInt32(r["TESTCONT"]);
                    c.version_contrato = Convert.ToInt32(r["CNUMSCPS"]);

                    //if (r["FPSERCON"] != System.DBNull.Value)
                    //    c.fecha_inicio_version = Convert.ToDateTime(r["FPSERCON"]);
                    //else
                    //    c.fecha_inicio_version = new DateTime(2010, 1, 1);

                    if (r["FPSERCON"] != System.DBNull.Value)
                    {
                        if (r["FPSERCON"].ToString() != "0000-00-00")
                            c.fecha_inicio_version = Convert.ToDateTime(r["FPSERCON"]);
                        else
                            c.fecha_inicio_version = c.fecha_inicio_version = new DateTime(2010, 1, 1);
                    }
                    else
                        c.fecha_inicio_version = c.fecha_inicio_version = new DateTime(2010, 1, 1);



                    List<EndesaEntity.contratacion.ContratosPS_Tabla> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.contratacion.ContratosPS_Tabla>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    o.Add(c);                  



                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private bool ExisteContratoHistorico(string cups20, int version, DateTime fecha)
        {
            List<EndesaEntity.contratacion.ContratosPS_Tabla> o;
            if (dic_historico.TryGetValue(cups20, out o))
            {
                for(int i = 0; i < o.Count; i++)
                {
                    if (o[i].version_contrato == version &&
                        o[i].fecha_inicio_version <= fecha)
                        return true;
                }
            }
            return false;
        }

        public bool TieneComplemento(string cups20)
        {        


            List<EndesaEntity.contratacion.ContratosPS_Tabla> o;

            if(cups20.Length != 13)
            {
                if (dic.TryGetValue(cups20, out o))
                    return true;
                else
                    return false;
            }
            else
            {
                foreach(KeyValuePair<string, List<EndesaEntity.contratacion.ContratosPS_Tabla>> p in dic)
                    for(int i = 0; i < p.Value.Count; i++)
                    {
                        if (p.Value[i].cups13 == cups20)
                            return true;
                    }
                return false;   
            }
           
        }
    }
}
