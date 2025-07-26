using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.cups
{
    public class PuntosSuministro
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "PuntosSuministro");

        public PuntosSuministro()
        {

        }

        public List<EndesaEntity.medida.PuntoSuministro> CompletaCups13(List<EndesaEntity.medida.PuntoSuministro> lc)
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
                    sb.Append("select CUPS_CORTO, CUPS20 from aux1.RELACION_CUPS where CUPS20 in (");
                    //sb.Append("select CUPS13 as CUPS_CORTO, CUPS20 from med.med_listado_scea_cups_historico where CUPS20 in (");
                    sb.Append("'").Append(lc[i].cups20).Append("'");
                    firstOnly = false;
                }
                else
                {
                    sb.Append(",'").Append(lc[i].cups20).Append("'");
                }

                if (j == 500)
                {
                    sb.Append(") and ACTIVO = 'S'");
                    Console.WriteLine("Guardando " + i + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(sb.ToString(), db.con);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {

                        for (int x = 0; x < lc.Count(); x++)
                        {
                            c = lc.Find(z => z.cups20 == reader["CUPS20"].ToString() && z.cups13 == null);
                            if (c != null)
                                lc[c.id].cups13 = reader["CUPS_CORTO"].ToString();
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
                sb.Append(") and ACTIVO = 'S'");
                Console.WriteLine("Consultando " + j + " registros...");
                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(sb.ToString(), db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    for (int i = 0; i < lc.Count(); i++)
                    {
                        c = lc.Find(z => z.cups20 == reader["CUPS20"].ToString() && z.cups13 == null);
                        if (c != null)
                            lc[c.id].cups13 = reader["CUPS_CORTO"].ToString();
                    }

                }

                db.CloseConnection();
                sb = null;
                sb = new StringBuilder();
                j = 0;

            }

            return lc;
        }
        public List<EndesaEntity.PuntoSuministro> CompletaCups22(List<EndesaEntity.PuntoSuministro> lc)
        {
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int j = 0;

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            EndesaEntity.PuntoSuministro c = new EndesaEntity.PuntoSuministro();

            try
            {
                for (int i = 0; i < lc.Count(); i++)
                {
                    j++;

                    if (firstOnly)
                    {
                        sb.Append("select CUPS_CORTO, CUPS22 from aux1.RELACION_CUPS where CUPS_CORTO in (");
                        sb.Append("'").Append(lc[i].cups13).Append("'");
                        firstOnly = false;
                    }
                    else
                    {
                        sb.Append(",'").Append(lc[i].cups13).Append("'");
                    }

                    if (j == 500)
                    {
                        sb.Append(") and ACTIVO = 'S'");
                        sb.Append(" and CUPS22 <> ''");
                        Console.WriteLine("Guardando " + i + " registros...");
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.AUX);
                        command = new MySqlCommand(sb.ToString(), db.con);
                        reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            c = lc.Find(z => z.cups13 == reader["CUPS_CORTO"].ToString());
                            if (reader["CUPS22"] != System.DBNull.Value)
                                lc[c.id].cups20 = reader["CUPS22"].ToString();
                        }

                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        j = 0;


                    }
                } // for (int i = 0; i < lc.Count(); i++)

                if (j > 0)
                {
                    sb.Append(") and ACTIVO = 'S'");
                    sb.Append(" and CUPS22 <> ''");
                    Console.WriteLine("Consultando " + j + " registros...");
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(sb.ToString(), db.con);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        c = lc.Find(z => z.cups13 == reader["CUPS_CORTO"].ToString());
                        if (reader["CUPS22"] != System.DBNull.Value && c != null)
                            lc[c.id].cups20 = reader["CUPS22"].ToString();
                    }

                    db.CloseConnection();
                    sb = null;

                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

            return lc;
        }
    }
}
