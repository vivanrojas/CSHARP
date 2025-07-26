using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.servidores
{
    public class AccessDB
    {
        public OleDbConnection con;

        public AccessDB(string dataBase)
        {
            try
            {
                con = new OleDbConnection(GetConnectionString(dataBase));
                con.Open();

            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private string GetConnectionString(string dataBase)
        {
            
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dataBase;
            // string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataBase;
            return connString;
        }

        public void CloseConnection()
        {
            con.Close();
            con.Dispose();
        }
    }
}
