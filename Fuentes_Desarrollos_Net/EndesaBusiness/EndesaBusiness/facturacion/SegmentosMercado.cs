using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class SegmentosMercado
    {
        public EndesaEntity.EmpresaTitular et;
        private List<EndesaEntity.EmpresaTitular> let = new List<EndesaEntity.EmpresaTitular>();
        private int pNumRegistros;

        public SegmentosMercado()
        {
            pNumRegistros = 0;
            this.CargaEmpresasTitulares();
        }

        private void CargaEmpresasTitulares()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select * from fo_empresas;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    pNumRegistros++;
                    et = new EndesaEntity.EmpresaTitular();
                    et.cemptitu = Convert.ToInt32(reader["cemptitu"]);
                    et.segmento = reader["segmento"].ToString();
                    et.descripcion = reader["descripcion"].ToString();
                    this.let.Add(et);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public void GetPosicionID(int pos)
        {
            et = new EndesaEntity.EmpresaTitular();
            et.cemptitu = let[pos - 1].cemptitu;
            et.segmento = let[pos - 1].segmento;
            et.descripcion = let[pos - 1].descripcion;
        }

        public int NumRegistros
        {
            get { return pNumRegistros; }
            set { pNumRegistros = value; }
        }
    }
}
