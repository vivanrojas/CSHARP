using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.cnae
{
    public class ImportarCNAE_RedShift
    {
        public ImportarCNAE_RedShift()
        {

        }


       

        public void Importar()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            servidores.MySQLDB mdb;
            MySqlCommand mcommand;
            StringBuilder sb = new StringBuilder();

            string strSql = "";
            int registrosLeidos = 0;
            int totalReg = 0;
            bool firstOnly = true;

            try
            {
                strSql = "SELECT SUBSTRING(c.cd_cups_ext,1,20) as cups20,"
                    + " c.cd_cnae"
                    + " FROM ed_owner.t_ed_h_crtos c"
                    + " where c.cd_cups_ext like 'ES%'"
                    + " and c.cd_cnae is not null"
                    + " group by c.cd_cups_ext, c.cd_cnae";
                db = new servidores.RedShiftServer();
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    registrosLeidos++;
                    totalReg++;

                    if (firstOnly)
                    {
                        sb.Append("replace into cnae_dt");
                        sb.Append(" (cups20, cnae) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(r["cups20"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cnae"].ToString()).Append("'),");

                    if (totalReg == 1000)
                    {

                        Console.CursorLeft = 0;
                        Console.Write(registrosLeidos);
                        firstOnly = true;
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        totalReg = 0;
                    }
                }
                db.CloseConnection();

                if (totalReg > 0)
                {
                    firstOnly = true;
                    Console.CursorLeft = 0;
                    Console.Write(registrosLeidos);
                    mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                    mcommand.ExecuteNonQuery();
                    mdb.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    totalReg = 0;
                }

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void Importar_t_ed_f_uvcrtos_diario()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;

            servidores.MySQLDB mdb;
            MySqlCommand mcommand;
            StringBuilder sb = new StringBuilder();

            string strSql = "";
            int registrosLeidos = 0;
            int totalReg = 0;
            bool firstOnly = true;

            try
            {
                strSql = "select d.cd_cups20_metra, d.cd_cnae"
                    + " FROM ed_owner.t_ed_f_uvcrtos_diario d"
                     + " where d.fh_baja_crto is null; ";
                db = new servidores.RedShiftServer();
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    registrosLeidos++;
                    totalReg++;

                    if (firstOnly)
                    {
                        sb.Append("replace into cnae_dt");
                        sb.Append(" (cups20, cnae) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(r["cd_cups20_metra"].ToString()).Append("',");
                    sb.Append("'").Append(r["cd_cnae"].ToString()).Append("'),");

                    if (totalReg == 500)
                    {

                        Console.CursorLeft = 0;
                        Console.Write(registrosLeidos);
                        firstOnly = true;
                        mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                        mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                        mcommand.ExecuteNonQuery();
                        mdb.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        totalReg = 0;
                    }
                }
                db.CloseConnection();

                if (totalReg > 0)
                {
                    firstOnly = true;
                    Console.CursorLeft = 0;
                    Console.Write(registrosLeidos);
                    mdb = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    mcommand = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), mdb.con);
                    mcommand.ExecuteNonQuery();
                    mdb.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    totalReg = 0;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
