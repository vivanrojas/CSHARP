using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.ocasional
{
    public class Ocasional_Adicionales_GAS_mant_NEDGIA
    {

        EndesaBusiness.utilidades.TelegramMensajes telegram;
        utilidades.Param pp;

        public Ocasional_Adicionales_GAS_mant_NEDGIA()
        {
            pp = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);
            telegram = new EndesaBusiness.utilidades.TelegramMensajes(pp.GetValue("telegram_token_contratacion"), 
                pp.GetValue("telegram_channel_id"));
        }

        public void Cambiar_FTP_SCTD()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_param SET VALUE = 'ftpgdi.endesa.es'" 
                     + "WHERE code = 'ftp_sctd_server';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha cambiado la URL del ftp del SCTD a ftpgdi.endesa.es");

            }
            catch (Exception e)
            {
                telegram.SendMessage("Error en cambio de ftp: " + e.Message);
            }
        }

        public void Cambiar_FTP_Contrata()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE ps_param SET VALUE = 'ftpgdi.endesa.es'"
                     + "WHERE code = 'ftp_server';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha cambiado la URL del ftp del SCTD a ftpgdi.endesa.es");

            }
            catch (Exception e)
            {
                telegram.SendMessage("Error en cambio de ftp: " + e.Message);
            }
        }

        public void Cambiar_a_Mail_NEDGIA()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_distribuidoras SET tramitacion = 'Mail' WHERE nombre = 'NEDGIA'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha aplicado la tramitacion de NEDGIA de XML_SCTD a Mail.");

            }catch(Exception e)
            {
                telegram.SendMessage("Error en cambio de tramitacion NEDGIA " + e.Message);
            }

           
        }

        public void Cambiar_a_Mail_REDEXIS()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_distribuidoras SET tramitacion = 'Mail' WHERE nombre = 'REDEXIS'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha aplicado la tramitacion de REDEXIS de XML_SCTD a Mail.");

            }
            catch (Exception e)
            {
                telegram.SendMessage("Error en cambio de tramitacion REDEXIS " + e.Message);
            }
        }

        public void Cambiar_a_Mail_NORTEGAS()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_distribuidoras SET tramitacion = 'Mail' WHERE nombre = 'Nortegas'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha aplicado la tramitacion de NORTEGAS de XML_SCTD a Mail.");

            }
            catch (Exception e)
            {
                telegram.SendMessage("Error en cambio de tramitacion NORTEGAS " + e.Message);
            }


        }

        public void Cambiar_a_SCTD_NEDGIA()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_distribuidoras SET tramitacion = 'XML_SCTD' WHERE nombre = 'NEDGIA';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha aplicado la tramitacion de NEDGIA Mail a XML_SCTD.");
            }
            catch(Exception e)
            {
                telegram.SendMessage("Error en cambio de tramitacion NEDGIA " + e.Message);
            }
           


        }

        public void Cambiar_a_SCTD_NORTEGAS()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_distribuidoras SET tramitacion = 'XML_SCTD' WHERE nombre = 'Nortegas';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha aplicado la tramitacion de NORTEGAS Mail a XML_SCTD.");
            }
            catch (Exception e)
            {
                telegram.SendMessage("Error en cambio de tramitacion NORTEGAS " + e.Message);
            }



        }

        public void Cambiar_a_SCTD_REDEXIS()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                strSql = "UPDATE atrgas_distribuidoras SET tramitacion = 'XML_SCTD' WHERE nombre = 'REDEXIS';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                telegram.SendMessage("Se ha aplicado la tramitacion de REDEXIS Mail a XML_SCTD.");
            }
            catch (Exception e)
            {
                telegram.SendMessage("Error en cambio de tramitacion REDEXIS " + e.Message);
            }



        }
    }
}
