using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot.Types;

namespace EndesaBusiness.medida.licitaciones
{
    public class Licitaciones
    {

        public Licitaciones()
        {

        }


        public void Extraccion()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();

            EndesaBusiness.utilidades.Fechas utilfechas =
                new Fechas();

            string linea = "";
            int num_periodos = 0;
            bool firstOnly = true;

            EndesaBusiness.adif.NuevaMedidaADIF medida_adif;

            FileInfo fileInfo = new FileInfo(@"c:\Temp\ADIF_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            StreamWriter swa = new StreamWriter(fileInfo.FullName, false);

            linea = "FECHA;HORA;AE;AS;R1;R2;R3;R4;CUPS22";
            swa.WriteLine(linea);

            strSql = "SELECT c.*, e.`FECHA DESDE` as fd, e.`FECHA HASTA` as fh"
                + " FROM med.adif_extraccion e"
                + " LEFT OUTER JOIN med.adif_cups c ON"
                + " c.CUPS20 = e.CUPS";
                //+ " WHERE c.CUPS20 = 'ES0027700041536001YR'";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                List<string> lista_cups = new List<string>();

                if (r["CUPS20"] != System.DBNull.Value)
                    lista_cups.Add(r["CUPS20"].ToString());

                if (r["fd"] != System.DBNull.Value)
                    fd = Convert.ToDateTime(r["fd"]);

                if (r["fh"] != System.DBNull.Value)
                    fh = Convert.ToDateTime(r["fh"]);

                medida_adif = new adif.NuevaMedidaADIF(lista_cups, fd, fh);
                
                
                foreach(KeyValuePair<string,List<EndesaEntity.medida.CurvaDeCarga>> p in medida_adif.dic_activa)
                    foreach(EndesaEntity.medida.CurvaDeCarga pp in p.Value)
                    {
                        num_periodos = utilfechas.NumPeriodosHorarios(pp.fecha);
                        for (int h = 0; h < num_periodos; h++)
                        {
                            switch (num_periodos)
                            {
                                case 24:
                                    linea = pp.fecha.Date.ToString("dd/MM/yyyy") + ";" +
                                        pp.fecha.AddHours(h).ToString("HH:mm") + ";" +
                                        pp.horaria_activa[h] + ";0;0;0;0;" +
                                        pp.horaria_reactiva[h] + ";" +
                                        p.Key;
                                    break;
                                case 23:
                                    if(h > 1)
                                    {

                                        linea = pp.fecha.Date.ToString("dd/MM/yyyy") + ";" +
                                            pp.fecha.AddHours(h+1).ToString("HH:mm") + ";" +
                                            pp.horaria_activa[h+1] + ";0;0;0;0;" +
                                            pp.horaria_reactiva[h+1] + ";" +
                                            p.Key;
                                    }
                                    else
                                    {
                                        linea = pp.fecha.Date.ToString("dd/MM/yyyy") + ";" +
                                            pp.fecha.AddHours(h).ToString("HH:mm") + ";" +
                                            pp.horaria_activa[h] + ";0;0;0;0;" +
                                            pp.horaria_reactiva[h] + ";" +
                                            p.Key;
                                    }

                                    break;
                            }

                            

                            swa.WriteLine(linea);
                        }
                    }
                        
                        

                
            }
            db.CloseConnection();
            swa.Close();

            MessageBox.Show("Informe terminado.",
                 "Informe",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);
        }


        public void Extraccion_P01011()
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            string cups20 = "";
            DateTime fechahora = new DateTime();
            bool firstOnly = true;
            int num_periodos = 0;
            string linea = "";

            Dictionary<string, Dictionary<DateTime, List<EndesaEntity.medida.P1>>> dic =
                new Dictionary<string, Dictionary<DateTime, List<EndesaEntity.medida.P1>>>();

            EndesaBusiness.utilidades.Fechas utilfechas =
                new Fechas();

            List<EndesaEntity.medida.P1> lista_p1 = new List<EndesaEntity.medida.P1>();

            FileInfo fileInfo = new FileInfo(@"c:\Temp\P01011_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt");



            #region Activa
            strSql = "SELECT c.*, e.`FECHA DESDE`, e.`FECHA HASTA`,m.*"
                + " FROM med.adif_extraccion e"
                + " LEFT OUTER JOIN med.adif_cups c ON"
                + " c.CUPS20 = e.CUPS"
                + " LEFT OUTER JOIN adif_medida_horaria_adif m ON"
                + " m.ID = c.ID_CUPS AND"
                + " m.Fecha >= e.`FECHA DESDE` AND m.Fecha <= e.`FECHA HASTA`"
                + " WHERE c.CUPS20 = 'ES0027700041536001YR' AND m.TipoEnergia = 'A'"
                + " ORDER BY c.ID_CUPS, m.Fecha,  m.TipoEnergia, m.FechaCarga desc";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                firstOnly = true;
                cups20 = r["CUPS20"].ToString();
                fechahora = Convert.ToDateTime(r["Fecha"]);
                
                if (r["CUPS20"] != System.DBNull.Value)
                {
                    num_periodos = utilfechas.NumPeriodosHorarios(fechahora);

                    switch(num_periodos)
                    {
                        case 24:
                            for (int h = 1; h <= 24; h++)
                            {
                                EndesaEntity.medida.P1 c = new EndesaEntity.medida.P1();
                                if (firstOnly)
                                {
                                    c.fecha_hora = fechahora;
                                    c.ae = Convert.ToDouble(r["Value" + h]);
                                    firstOnly = false;
                                }
                                else
                                {
                                    c.ae = Convert.ToDouble(r["Value" + h]);
                                    c.fecha_hora = c.fecha_hora.AddHours(h);
                                }

                                Dictionary<DateTime, List<EndesaEntity.medida.P1>> o;
                                if (!dic.TryGetValue(cups20, out o))
                                {
                                    lista_p1 = new List<EndesaEntity.medida.P1>();
                                    lista_p1.Add(c);
                                    o = new Dictionary<DateTime, List<EndesaEntity.medida.P1>>();
                                    o.Add(fechahora, lista_p1);
                                    dic.Add(cups20, o);
                                }
                                else
                                {
                                    List<EndesaEntity.medida.P1> oo;
                                    if (!o.TryGetValue(fechahora, out oo))
                                    {
                                        oo = new List<EndesaEntity.medida.P1>();
                                        oo.Add(c);
                                        o.Add(fechahora, oo);
                                    }
                                    else
                                        oo.Add(c);

                                }


                            }
                            break;
                        case 23:
                            for (int h = 1; h <= 23; h++)
                            {
                                EndesaEntity.medida.P1 c = new EndesaEntity.medida.P1();
                                if (firstOnly)
                                {
                                    c.fecha_hora = fechahora;
                                    c.ae = Convert.ToDouble(r["Value" + h]);
                                    firstOnly = false;
                                }
                                else
                                {
                                    if (h >= 2)
                                    {
                                        c.ae = Convert.ToDouble(r["Value" + (h + 1)]);
                                        c.fecha_hora = c.fecha_hora.AddHours(h + 1);
                                    }
                                    else
                                    {
                                        c.ae = Convert.ToDouble(r["Value" + h]);
                                        c.fecha_hora = c.fecha_hora.AddHours(h);
                                    }
                                       
                                    
                                }

                                Dictionary<DateTime, List<EndesaEntity.medida.P1>> o;
                                if (!dic.TryGetValue(cups20, out o))
                                {
                                    lista_p1 = new List<EndesaEntity.medida.P1>();
                                    lista_p1.Add(c);
                                    o = new Dictionary<DateTime, List<EndesaEntity.medida.P1>>();
                                    o.Add(fechahora, lista_p1);
                                    dic.Add(cups20, o);
                                }
                                else
                                {
                                    List<EndesaEntity.medida.P1> oo;
                                    if (!o.TryGetValue(fechahora, out oo))
                                    {
                                        oo = new List<EndesaEntity.medida.P1>();
                                        oo.Add(c);
                                        o.Add(fechahora, oo);
                                    }
                                    else
                                        oo.Add(c);

                                }


                            }
                            break;
                    }

                    
                }
                

            }
            db.CloseConnection();

            #endregion

            #region Reactiva
            strSql = "SELECT c.*, e.`FECHA DESDE`, e.`FECHA HASTA`,m.*"
                + " FROM med.adif_extraccion e"
                + " LEFT OUTER JOIN med.adif_cups c ON"
                + " c.CUPS20 = e.CUPS"
                + " LEFT OUTER JOIN adif_medida_horaria_adif m ON"
                + " m.ID = c.ID_CUPS AND"
                + " m.Fecha >= e.`FECHA DESDE` AND m.Fecha <= e.`FECHA HASTA`"
                + " WHERE c.CUPS20 = 'ES0027700041536001YR' AND m.TipoEnergia = 'R'"
                + " ORDER BY c.ID_CUPS, m.Fecha,  m.TipoEnergia, m.FechaCarga desc";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                firstOnly = true;
                cups20 = r["CUPS20"].ToString();
                fechahora = Convert.ToDateTime(r["Fecha"]);

                if (r["CUPS20"] != System.DBNull.Value)
                {
                    num_periodos = utilfechas.NumPeriodosHorarios(fechahora);

                    switch (num_periodos)
                    {
                        case 24:
                            for (int h = 1; h <= 24; h++)
                            {
                                EndesaEntity.medida.P1 c = new EndesaEntity.medida.P1();
                                if (firstOnly)
                                {
                                    c.fecha_hora = fechahora;
                                    c.reactiva[3] = Convert.ToDouble(r["Value" + h]);
                                    firstOnly = false;
                                }
                                else
                                {
                                    c.reactiva[3] = Convert.ToDouble(r["Value" + h]);
                                    c.fecha_hora = c.fecha_hora.AddHours(h);
                                }

                                Dictionary<DateTime, List<EndesaEntity.medida.P1>> o;
                                if (!dic.TryGetValue(cups20, out o))
                                {
                                    lista_p1 = new List<EndesaEntity.medida.P1>();
                                    lista_p1.Add(c);
                                    o = new Dictionary<DateTime, List<EndesaEntity.medida.P1>>();
                                    o.Add(fechahora, lista_p1);
                                    dic.Add(cups20, o);
                                }
                                else
                                {
                                    List<EndesaEntity.medida.P1> oo;
                                    if (!o.TryGetValue(fechahora, out oo))
                                    {
                                        oo = new List<EndesaEntity.medida.P1>();
                                        oo.Add(c);
                                        o.Add(fechahora, oo);
                                    }
                                    else
                                        oo.Add(c);

                                }


                            }
                            break;
                        case 23:
                            for (int h = 1; h <= 23; h++)
                            {
                                EndesaEntity.medida.P1 c = new EndesaEntity.medida.P1();
                                if (firstOnly)
                                {
                                    c.fecha_hora = fechahora;
                                    c.reactiva[3] = Convert.ToDouble(r["Value" + h]);
                                    firstOnly = false;
                                }
                                else
                                {
                                    if (h >= 2)
                                    {
                                        c.reactiva[3] = Convert.ToDouble(r["Value" + (h + 1)]);
                                        c.fecha_hora = c.fecha_hora.AddHours(h + 1);
                                    }
                                    else
                                    {
                                        c.reactiva[3] = Convert.ToDouble(r["Value" + h]);
                                        c.fecha_hora = c.fecha_hora.AddHours(h);
                                    }
                                }

                                Dictionary<DateTime, List<EndesaEntity.medida.P1>> o;
                                if (!dic.TryGetValue(cups20, out o))
                                {
                                    lista_p1 = new List<EndesaEntity.medida.P1>();
                                    lista_p1.Add(c);
                                    o = new Dictionary<DateTime, List<EndesaEntity.medida.P1>>();
                                    o.Add(fechahora, lista_p1);
                                    dic.Add(cups20, o);
                                }
                                else
                                {
                                    List<EndesaEntity.medida.P1> oo;
                                    if (!o.TryGetValue(fechahora, out oo))
                                    {
                                        oo = new List<EndesaEntity.medida.P1>();
                                        oo.Add(c);
                                        o.Add(fechahora, oo);
                                    }
                                    else
                                        oo.Add(c);

                                }


                            }
                            break;
                    }


                }


            }
            db.CloseConnection();

            #endregion

            StreamWriter swa = new StreamWriter(fileInfo.FullName, false);
            foreach(KeyValuePair<string, Dictionary<DateTime,List<EndesaEntity.medida.P1>>> p in dic)
                foreach(KeyValuePair<DateTime, List<EndesaEntity.medida.P1>> pp in p.Value)
                {
                    foreach(EndesaEntity.medida.P1 y in pp.Value)
                        linea = y.fecha_hora.Date.ToString("dd/MM/yyyy");
                }
                
            
            swa.Close();


            MessageBox.Show("Informe terminado.",
                  "Informe",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
        }



    }
}
