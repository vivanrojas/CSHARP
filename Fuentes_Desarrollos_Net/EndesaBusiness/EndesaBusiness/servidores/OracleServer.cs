using System;
using Oracle.ManagedDataAccess.Client;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.servidores
{
    public class OracleServer
    {
        public enum Servidores
        {
            BDO,
            OWE,
            COMPOR,
            COMPOR_DES,
            LOCAL

        }
        public OracleConnection con;


        public OracleServer(Servidores servidor)
        {
            try
            {
                con = new OracleConnection(GetConnectionString(servidor));
                con.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        public void CloseConnection()
        {
            con.Close();
            con.Dispose();
            //Console.WriteLine("Se ha desconectado");
        }

        private string GetConnectionString(Servidores servidor)
        {
            string server = null;
            string user = null;
            string pass = null;
            string port = null;
            string ddbb = null;

            switch (servidor)
            {
                case Servidores.BDO:                    
                    server = "elaaemco1d01.enelint.global";
                    ddbb = "OCO1BE";
                    user = "BDO_APL";
                    pass = "BD123_APL";
                    port = "1521";
                    break;
                case Servidores.OWE:
                    server = "elaaemowed00.enelint.global";
                    ddbb = "OOWEAE";
                    user = "OWEN_APL";
                    pass = "ow3_n4pl";
                    port = "1521";
                    break;
                case Servidores.COMPOR:
                    server = "elaaemco1d01.enelint.global";
                    ddbb = "OCO1BE";
                    user = "COMPOR_APL";
                    pass = "c0mp0r07";
                    port = "1521";
                    break;
                case Servidores.COMPOR_DES:
                    server = "plaaemco1d00.enelint.global";
                    ddbb = "OCO1BE";
                    user = "COMPOR_OWNER";
                    pass = "seQ20-wzR";
                    port = "1521";
                    break;
                case Servidores.LOCAL:
                    server = "DESKTOP-G3C4EQV";
                    ddbb = "XE";
                    //ddbb = "XEPDB1";
                    user = "system";
                    pass = "vernao10";
                    port = "1521";
                    break;

            }


            return "User Id=" + user + ";Password=" + pass + ";Data Source=(DESCRIPTION="
                + "(ADDRESS=(PROTOCOL=TCP)(HOST=" + server + ")(PORT=" + port + "))"
                + "(CONNECT_DATA=(SID=" + ddbb + ")));";


        }
    }
}
