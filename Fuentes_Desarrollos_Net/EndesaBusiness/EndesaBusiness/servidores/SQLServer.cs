using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.servidores
{
    class SQLServer
    {
        public SqlConnection con;
        public SQLServer()
        {
            try
            {
                con = new SqlConnection(GetConnectionString());
                con.Open();

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
            }

        }


        private string GetConnectionString()
        {
            string server = null;
            string ddbb = null;
            string user = null;
            string pass = null;

            // server = "RLPDPW902";
            server = "rdssigamepro.endesa.es";
            ddbb = "SIGAME";
            user = "SGM_Consulta";
            pass = "b2ngl2d3sh";
            return "Data Source = " + server + "; Initial Catalog = " + ddbb + "; uid = " + user + "; pwd = " + pass;
        }

        public void CloseConnection()
        {
            con.Close();
            con.Dispose();

        }
    }
}
