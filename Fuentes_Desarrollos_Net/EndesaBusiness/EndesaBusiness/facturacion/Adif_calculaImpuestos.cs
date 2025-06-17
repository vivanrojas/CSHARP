using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class Adif_calculaImpuestos
    {

        private List<EndesaEntity.facturacion.Adif_impuestos> impuestos =
            new List<EndesaEntity.facturacion.Adif_impuestos>();
        public Adif_calculaImpuestos()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            try
            {
                strSql = "Select * from adif_impuestos";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.facturacion.Adif_impuestos impuesto =
                        new EndesaEntity.facturacion.Adif_impuestos();

                    impuesto.tipoImpuesto = reader["tipoImpuesto"].ToString();
                    impuesto.descripcion = reader["descripcion"].ToString();
                    impuesto.fechaDesde = Convert.ToDateTime(reader["fechadesde"]);
                    impuesto.fechaHasta = Convert.ToDateTime(reader["fechahasta"]);
                    impuesto.valor = Convert.ToDouble(reader["valor"]);
                    impuesto.unidad = reader["unidad"].ToString();
                    impuestos.Add(impuesto);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public double ImpuestoValor(String tipoImpuesto, DateTime fd, DateTime fh)
        {

            EndesaEntity.facturacion.Adif_impuestos impuesto =
                new EndesaEntity.facturacion.Adif_impuestos();
            impuesto = impuestos.Find(p => (p.tipoImpuesto == tipoImpuesto &&
                (p.fechaDesde <= fd && p.fechaHasta >= fh)));

            return impuesto.valor;
        }
    }
}
