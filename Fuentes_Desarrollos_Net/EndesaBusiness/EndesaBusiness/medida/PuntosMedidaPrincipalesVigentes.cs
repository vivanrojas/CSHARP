using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class PuntosMedidaPrincipalesVigentes
    {


        // cups20
        public Dictionary<string, EndesaEntity.medida.PuntoSuministro> dic { get; set; }
        public List<string> lista_cups15 { get; set; }
        public List<string> lista_cups22 { get; set; }

        public PuntosMedidaPrincipalesVigentes(List<string> lista_cups)
        {
            dic = new Dictionary<string, EndesaEntity.medida.PuntoSuministro>();
            BuscaDatos(lista_cups);
            lista_cups15 = new List<string>();
            lista_cups22 = new List<string>();
            foreach (KeyValuePair<string, EndesaEntity.medida.PuntoSuministro> p in dic)
            {
                for (int i = 0; i < p.Value.cups15.Count(); i++)
                    lista_cups15.Add(p.Value.cups15[i]);

                for (int i = 0; i < p.Value.cups22.Count(); i++)
                    lista_cups22.Add(p.Value.cups22[i]);
            }

            // Copiamos los puntos de que no encontramos vigentes a sus listas correspondientes

            //for(int i = 0; lista_cups.Count; i++)
            //{
            //    if(lista_cups[i].Length > 13)
            //        if(lista_cups22.Find.)
            //}



        }

        public PuntosMedidaPrincipalesVigentes(List<string> lista_cups, DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, EndesaEntity.medida.PuntoSuministro>();                        
            BuscaDatos(lista_cups);
            BuscaDatosHistorico(lista_cups, fd, fh);
            lista_cups15 = new List<string>();
            lista_cups22 = new List<string>();
            foreach (KeyValuePair<string, EndesaEntity.medida.PuntoSuministro> p in dic)
            {
                for (int i = 0; i < p.Value.cups15.Count(); i++)
                    lista_cups15.Add(p.Value.cups15[i]);

                for (int i = 0; i < p.Value.cups22.Count(); i++)
                    lista_cups22.Add(p.Value.cups22[i]);
            }

            // Copiamos los puntos de que no encontramos vigentes a sus listas correspondientes

            //for(int i = 0; lista_cups.Count; i++)
            //{
            //    if(lista_cups[i].Length > 13)
            //        if(lista_cups22.Find.)
            //}



        }

        public PuntosMedidaPrincipalesVigentes(List<EndesaEntity.medida.PuntoSuministro> lc)
        {
            List<string> lista_cups = new List<string>();
            dic = new Dictionary<string, EndesaEntity.medida.PuntoSuministro>();

            for (int i = 0; i < lc.Count(); i++)
            {
                if (lc[i].cups13 != null)
                    lista_cups.Add(lc[i].cups13);
                if (lc[i].cups20 != null)
                    lista_cups.Add(lc[i].cups20);
            }

            BuscaDatos(lista_cups);
            lista_cups15 = new List<string>();
            lista_cups22 = new List<string>();
            foreach (KeyValuePair<string, EndesaEntity.medida.PuntoSuministro> p in dic)
            {
                for (int i = 0; i < p.Value.cups15.Count(); i++)
                    lista_cups15.Add(p.Value.cups15[i]);

                for (int i = 0; i < p.Value.cups22.Count(); i++)
                    lista_cups22.Add(p.Value.cups22[i]);
            }
        }

        private void BuscaDatos(List<string> lista_cups)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int j = 0;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


            for (int i = 0; i < lista_cups.Count(); i++)
            {
                j++;

                if (firstOnly)
                {
                    sb.Append("select cd_counips as cups13, cd_puntmed as cups15,");
                    sb.Append(" cd_cups20 as cups20, cd_cups22 as cups22 from med.dt_vw_ed_f_puntos_ml where");
                    if (lista_cups[i].Length == 13)
                    {
                        sb.Append(" cd_counips in (");
                        sb.Append("'").Append(lista_cups[i]).Append("'");
                    }
                    else
                    {
                        sb.Append(" substr(cd_cups22,1,20) in (");
                        sb.Append("'").Append(lista_cups[i]).Append("'");
                    }

                    firstOnly = false;
                }
                else
                {
                    sb.Append(",'").Append(lista_cups[i]).Append("'");
                }

                if (j == 500)
                {
                    sb.Append(") and cd_tfunpmed = 'P'");
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString(), db.con);
                    r = command.ExecuteReader();

                    while (r.Read())
                    {

                        EndesaEntity.medida.PuntoSuministro o;
                        if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                        {
                            EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                            if (r["cups13"] != System.DBNull.Value)
                                c.cups13 = r["cups13"].ToString();

                            if (r["cups20"] != System.DBNull.Value)
                                c.cups20 = r["cups20"].ToString();

                            if (r["cups15"] != System.DBNull.Value)
                                c.cups15.Add(r["cups15"].ToString());

                            if (r["cups22"] != System.DBNull.Value)
                            {
                                c.cups22.Add(r["cups22"].ToString());
                                if (c.cups20 == null)
                                    c.cups20 = r["cups22"].ToString().Substring(0, 20);
                            }

                            dic.Add(c.cups20, c);
                        }
                        else
                        {
                            if (r["cups15"] != System.DBNull.Value)
                                o.cups15.Add(r["cups15"].ToString());

                            if (r["cups22"] != System.DBNull.Value)
                                o.cups22.Add(r["cups22"].ToString());
                        }

                    }

                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;


                }
            } // for (int i = 0; i < lc.Count(); i++)    

            if (j > 0)
            {
                sb.Append(") and cd_tfunpmed = 'P'");
                Console.WriteLine("Consultando " + j + " registros...");
                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.medida.PuntoSuministro o;
                    if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                    {
                        EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                        if (r["cups13"] != System.DBNull.Value)
                            c.cups13 = r["cups13"].ToString();

                        if (r["cups20"] != System.DBNull.Value)
                            c.cups20 = r["cups20"].ToString();

                        if (r["cups15"] != System.DBNull.Value)
                            c.cups15.Add(r["cups15"].ToString());

                        if (r["cups22"] != System.DBNull.Value)
                        {
                            c.cups22.Add(r["cups22"].ToString());
                            if (c.cups20 == null)
                                c.cups20 = r["cups22"].ToString().Substring(0, 20);
                        }

                        dic.Add(c.cups20, c);
                    }
                    else
                    {
                        if (r["cups15"] != System.DBNull.Value)
                            o.cups15.Add(r["cups15"].ToString());

                        if (r["cups22"] != System.DBNull.Value)
                            o.cups22.Add(r["cups22"].ToString());
                    }

                }

                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                j = 0;

            }
        }

        public string GetCUPS22(string cups20)
        {
            EndesaEntity.medida.PuntoSuministro o;
            if (dic.TryGetValue(cups20, out o))
                return o.cups22[0];
            else
                return cups20;
        }

        private void BuscaDatosHistorico(List<string> lista_cups, DateTime fd, DateTime fh)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int j = 0;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            for (int i = 0; i < lista_cups.Count(); i++)
            {
                j++;

                if (firstOnly)
                {
                    sb.Append("select cd_cups as cups13, cd_puntmed as cups15,");
                    sb.Append(" substr(cups_pm,1,20) as cups20, cups_pm as cups22 from med.dt_t_ed_h_puntos_medida where");
                    sb.Append(" fh_altapme <= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and");
                    sb.Append(" (fh_bajapme >= '").Append(fh.ToString("yyyy-MM-dd")).Append("' or fh_bajapme IS NULL) and");

                    if (lista_cups[i].Length == 13)
                    {
                        sb.Append(" cd_cups in (");
                        sb.Append("'").Append(lista_cups[i]).Append("'");
                    }
                    else
                    {
                        sb.Append(" substr(cups_pm,1,20) in (");
                        sb.Append("'").Append(lista_cups[i]).Append("'");
                    }

                    firstOnly = false;
                }
                else
                {
                    sb.Append(",'").Append(lista_cups[i]).Append("'");
                }

                if (j == 500)
                {
                    sb.Append(") and cd_tfunpmed = 'P' group by cd_puntmed");
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString(), db.con);
                    r = command.ExecuteReader();

                    while (r.Read())
                    {

                        EndesaEntity.medida.PuntoSuministro o;
                        if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                        {
                            EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                            if (r["cups13"] != System.DBNull.Value)
                                c.cups13 = r["cups13"].ToString();

                            if (r["cups20"] != System.DBNull.Value)
                                c.cups20 = r["cups20"].ToString();

                            if (r["cups15"] != System.DBNull.Value)
                                c.cups15.Add(r["cups15"].ToString());

                            if (r["cups22"] != System.DBNull.Value)
                            {
                                c.cups22.Add(r["cups22"].ToString());
                                if (c.cups20 == null)
                                    c.cups20 = r["cups22"].ToString().Substring(0, 20);
                            }
                            if (c.cups20 != null)
                                dic.Add(c.cups20, c);
                        }
                        //else
                        //{
                        //    if (r["cups15"] != System.DBNull.Value)
                        //        o.cups15.Add(r["cups15"].ToString());

                        //    if (r["cups22"] != System.DBNull.Value)
                        //        o.cups22.Add(r["cups22"].ToString());
                        //}

                    }

                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;


                }
            } // for (int i = 0; i < lc.Count(); i++)    

            if (j > 0)
            {
                sb.Append(") and cd_tfunpmed = 'P' group by cd_puntmed");
                Console.WriteLine("Consultando " + j + " registros...");
                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.medida.PuntoSuministro o;
                    if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                    {
                        EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                        if (r["cups13"] != System.DBNull.Value)
                            c.cups13 = r["cups13"].ToString();

                        if (r["cups20"] != System.DBNull.Value)
                            c.cups20 = r["cups20"].ToString();

                        if (r["cups15"] != System.DBNull.Value)
                            c.cups15.Add(r["cups15"].ToString());

                        if (r["cups22"] != System.DBNull.Value)
                        {
                            c.cups22.Add(r["cups20"].ToString());
                            if (c.cups20 == null)
                                c.cups20 = r["cups22"].ToString().Substring(0, 20);
                        }
                        if (c.cups20 != null)
                            dic.Add(c.cups20, c);
                    }
                    //else
                    //{
                    //    if (r["cups15"] != System.DBNull.Value)
                    //        o.cups15.Add(r["cups15"].ToString());

                    //    if (r["cups22"] != System.DBNull.Value)
                    //        o.cups22.Add(r["cups22"].ToString());
                    //}

                }

                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                j = 0;

            }
        }


        // Dada una lista de cups13 nos devuelve la lista de cups15 principales
        public List<EndesaEntity.medida.PuntoSuministro> CompletaCups15(List<EndesaEntity.medida.PuntoSuministro> lc)
        {

            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int j = 0;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();

            for (int i = 0; i < lc.Count(); i++)
            {
                j++;

                if (firstOnly)
                {
                    sb.Append("select cd_counips as cups13, cd_puntmed as cups15 from med.dt_vw_ed_f_puntos_ml where cd_counips in (");
                    sb.Append("'").Append(lc[i].cups13).Append("'");
                    firstOnly = false;
                }
                else
                {
                    sb.Append(",'").Append(lc[i].cups13).Append("'");
                }

                if (j == 500)
                {
                    sb.Append(") and cd_tfunpmed = 'P'");
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString(), db.con);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        c = lc.Find(z => z.cups13 == reader["cups13"].ToString());
                        if (c != null)
                            lc[c.id].cups15.Add(reader["cups15"].ToString());
                    }

                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    j = 0;


                }
            } // for (int i = 0; i < lc.Count(); i++)    

            if (j > 0)
            {
                sb.Append(") and cd_tfunpmed = 'P'");
                Console.WriteLine("Consultando " + j + " registros...");
                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(sb.ToString(), db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    c = lc.Find(z => z.cups13 == reader["cups13"].ToString());
                    if (c != null)
                        lc[c.id].cups15.Add(reader["cups15"].ToString());

                }

                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                j = 0;

            }

            return lc;
        }
    }
}

