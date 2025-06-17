using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class CurvaCuartoHorariaSCE
    {
        Dictionary<string, EndesaEntity.medida.DicCurva> dic;
        public CurvaCuartoHorariaSCE(List<string> lista_cups15, DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, EndesaEntity.medida.DicCurva>();
            CargaMedidaCuartoHoraria(lista_cups15, fd, fh, "F");
            CargaMedidaCuartoHoraria(lista_cups15, fd, fh, "R");
        }

        private void CargaMedidaCuartoHoraria(List<string> lista_cups15, DateTime fd, DateTime fh, string estado)
        {
            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            int j = 0;

            try
            {
                for (int i = 0; i < lista_cups15.Count; i++)
                {
                    j++;
                    if (firstOnly)
                    {
                        sb.Append("select CCOUNIPS as cups15, Fecha, Unidad as estado, totala, totalr");
                        for (int x = 1; x <= 25; x++)
                            sb.Append(" ,A" + x);

                        for (int x = 1; x <= 25; x++)
                            sb.Append(" ,R" + x);

                        sb.Append(" from med.medida_cuartohoraria_sce ");
                        sb.Append(" where");
                        sb.Append(" CCOUNIPS in (");
                        sb.Append("'").Append(lista_cups15[i]).Append("'");
                        firstOnly = false;
                    }
                    else
                    {
                        sb.Append(" ,'").Append(lista_cups15[i]).Append("'");
                    }
                    if (j == 50)
                    {
                        sb.Append(")");
                        sb.Append(" and (Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and Fecha <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                        sb.Append(" and Unidad = '").Append(estado).Append("';");

                        j = 0;
                        firstOnly = true;
                        RunQuery(sb.ToString());
                        sb = null;
                        sb = new StringBuilder();
                    }

                }

                if (j > 0)
                {
                    sb.Append(")");
                    sb.Append(" and (Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and Fecha <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                    sb.Append(" and Unidad = '").Append(estado).Append("';");
                    j = 0;
                    firstOnly = true;
                    RunQuery(sb.ToString());
                    sb = null;

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "CurvaCuartoHoraria.CargaMedidaCuartoHoraria",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }


        private void RunQuery(string q)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(q, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {


                EndesaEntity.medida.CurvaDeCarga c = new EndesaEntity.medida.CurvaDeCarga();

                c.fecha = Convert.ToDateTime(r["Fecha"]);

                if (r["totala"] != System.DBNull.Value)
                    c.total_activa = Convert.ToDouble(r["totala"]);
                if (r["totalr"] != System.DBNull.Value)
                    c.total_reactiva = Convert.ToDouble(r["totalr"]);

                for (int i = 1; i < 26; i++)
                {
                    if (r["a" + i] != System.DBNull.Value)
                        c.horaria_activa[i - 1] = Convert.ToDouble(r["a" + i]);
                    if (r["r" + i] != System.DBNull.Value)
                        c.horaria_reactiva[i - 1] = Convert.ToDouble(r["r" + i]);

                }

                EndesaEntity.medida.DicCurva o;
                if (!dic.TryGetValue(r["cups15"].ToString(), out o))
                {
                    EndesaEntity.medida.DicCurva d = new EndesaEntity.medida.DicCurva();
                    d.dic.Add(c.fecha, c);
                    dic.Add(r["cups15"].ToString(), d);
                }
                else
                {
                    EndesaEntity.medida.CurvaDeCarga x;
                    if (!o.dic.TryGetValue(c.fecha, out x))
                    {
                        o.dic.Add(c.fecha, c);
                    }
                }

            }
            db.CloseConnection();
        }

        public List<EndesaEntity.medida.CurvaDeCarga> GetCurva(string cups13, DateTime fd, DateTime fh)
        {
            bool firstOnly = true;
            List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
            Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> d = new Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga>();

            Dictionary<string, EndesaEntity.medida.DicCurva> matches =
                dic.Where(z => z.Key.Substring(0, 13) == cups13).ToDictionary(z => z.Key, z => z.Value);

            foreach (KeyValuePair<string, EndesaEntity.medida.DicCurva> p in matches)
            {
                if (firstOnly)
                {
                    lista = p.Value.dic.Where(z => z.Key >= fd && z.Key <= fh).Select(z => z.Value).ToList();
                    firstOnly = false;
                }
                else
                {
                    for (int i = 0; i < lista.Count(); i++)
                    {
                        EndesaEntity.medida.CurvaDeCarga o;
                        if (p.Value.dic.TryGetValue(lista[i].fecha, out o))
                        {
                            lista[i].total_activa += o.total_activa;
                            lista[i].total_reactiva += o.total_reactiva;
                            for (int j = 0; j < 25; j++)
                            {
                                lista[i].horaria_activa[j] += o.horaria_activa[j];
                                lista[i].horaria_reactiva[j] += o.horaria_reactiva[j];
                            }
                        }
                    }
                }


            }
            return lista;

        }

        public double TotalActiva(string cups13, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.medida.CurvaDeCarga> lista = GetCurva(cups13, fd, fh);
            if (lista != null)
                total = lista.Sum(z => z.total_activa);

            return total;
        }

        public double TotalReactiva(string cups13, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.medida.CurvaDeCarga> lista = GetCurva(cups13, fd, fh);
            if (lista != null)
                total = lista.Sum(z => z.total_reactiva);

            return total;
        }
    }
}
