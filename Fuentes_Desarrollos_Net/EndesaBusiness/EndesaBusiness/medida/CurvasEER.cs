using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;

namespace EndesaBusiness.medida
{
    public class CurvasEER 
    {

        Dictionary<string, EndesaEntity.medida.DicCurva> dic;
        public List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
        calendarios.UtilidadesCalendario utilCal;
        public int numPeriodosMedidaHorario { get; set; }
        public int numPeriodosMedidaCuartoHorario { get; set; }

        public bool curvaCompleta { get; set; }

        public double[] curvaCuartoHorariaActiva { get; set; }
        public double[] curvaCuartoHorariaReactiva { get; set; }
        public double[] curvaCuartoHorariaReactiva_R1 { get; set; }
        public double[] curvaCuartoHorariaReactiva_R4 { get; set; }

        public DateTime[] curvaHorariaDias { get; set; }
        public DateTime[] curvaCuartoHorariaDias { get; set; }
        

        public double[] curvaHorariaActiva { get; set; }
        public double[] curvaHorariaReactiva { get; set; }
        public double[] curvaCuartoHorariaPotencias { get; set; }
        public string[] curvaCuartoHorariaEstacion { get; set; }
        public string[] curvaCuartoHorariaFuente { get; set; }
        public double totalEnergiaActiva { get; set; }
        public double totalEnergiaReactiva { get; set; }
        public string cups20 { get; set; }

        //public double[] vectorExcesos;
        //public double[] vectorPotenciasMaximasRegistradas;

        DateTime _fd = new DateTime();
        DateTime _fh = new DateTime();

        //calendarios.Calendario calendario;


        public CurvasEER()
        {

        }

        public CurvasEER(EndesaEntity.punto_suministro.PuntoSuministro ps, DateTime fd, DateTime fh)
        {

            this.cups20 = ps.cups20;

            //calendario = new calendarios.Calendario(fd, fh);

            utilCal = new calendarios.UtilidadesCalendario();
            numPeriodosMedidaHorario = (Convert.ToInt32((fh - fd).TotalDays + 1) * 24) + utilCal.CorreccionCambioHorario(fd, fh);
            numPeriodosMedidaCuartoHorario = (numPeriodosMedidaHorario * 4);

            curvaHorariaActiva = new double[numPeriodosMedidaHorario + 1];
            curvaHorariaReactiva = new double[numPeriodosMedidaHorario + 1];
            curvaCuartoHorariaActiva = new double[numPeriodosMedidaCuartoHorario + 1];

            curvaCuartoHorariaReactiva = new double[numPeriodosMedidaCuartoHorario + 1];
            curvaCuartoHorariaReactiva_R1 = new double[numPeriodosMedidaCuartoHorario + 1];
            curvaCuartoHorariaReactiva_R4 = new double[numPeriodosMedidaCuartoHorario + 1];

            curvaHorariaDias = new DateTime[numPeriodosMedidaHorario + 1];
            curvaCuartoHorariaDias = new DateTime[numPeriodosMedidaCuartoHorario + 1];

            curvaCuartoHorariaPotencias = new double[numPeriodosMedidaCuartoHorario + 1];
            curvaCuartoHorariaEstacion = new string[numPeriodosMedidaCuartoHorario + 1];
            curvaCuartoHorariaFuente = new string[numPeriodosMedidaCuartoHorario + 1];

        }
                

        public CurvasEER(EndesaEntity.punto_suministro.PuntoSuministro ps, DateTime fd, DateTime fh, string estado, bool vertical, bool temporal)
        {
            _fd = fd;
            _fh = fh;

            //calendario = new calendarios.Calendario(fd, fh);

            utilCal = new calendarios.UtilidadesCalendario();
            numPeriodosMedidaHorario = (Convert.ToInt32((fh - fd).TotalDays + 1) * 24) + utilCal.CorreccionCambioHorario(fd, fh);
            numPeriodosMedidaCuartoHorario = (numPeriodosMedidaHorario * 4);

            curvaHorariaActiva = new double[numPeriodosMedidaHorario + 1];
            curvaHorariaReactiva = new double[numPeriodosMedidaHorario + 1];
            curvaCuartoHorariaActiva = new double[numPeriodosMedidaCuartoHorario + 1];

            curvaCuartoHorariaReactiva = new double[numPeriodosMedidaCuartoHorario + 1];
            curvaCuartoHorariaReactiva_R1 = new double[numPeriodosMedidaCuartoHorario + 1];
            curvaCuartoHorariaReactiva_R4 = new double[numPeriodosMedidaCuartoHorario + 1];

            curvaHorariaDias = new DateTime[numPeriodosMedidaHorario + 1];
            curvaCuartoHorariaDias = new DateTime[numPeriodosMedidaCuartoHorario + 1];            

            curvaCuartoHorariaPotencias = new double[numPeriodosMedidaCuartoHorario + 1];

            //vectorExcesos = CargaExcesos(calendario,ps);

            // Determinamos si la lista de cups es de cups13 o cups20
            dic = new Dictionary<string, EndesaEntity.medida.DicCurva>();            
            lista = new List<EndesaEntity.medida.CurvaDeCarga>();
            if (!vertical)
            {
                //GetCurvaHorizontal(lista_cups22, fd, fh, estado);            
            }
            else
                GetCurvaVertical(ps.cups20, fd, fh, temporal);

        }


        private void GetCurvaVertical(string cups20, DateTime fd, DateTime fh, bool temporal)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;


            int numeroPeriodosHorarios = 0;
            int parcialPeriodosHorarios = 0;
            int numeroPeriodosCuartoHorarios = 0;
            double energiaActiva = 0;
            double energiaReactiva = 0;
            int hora = 0;
            int minutos = 0;
            
            try
            {
                calendarios.UtilidadesCalendario util = new calendarios.UtilidadesCalendario();

                //for(DateTime d = fd; d <= fh; d = d.AddDays(1))
                //{

                //}


                #region Query

                strSql = "SELECT cups20, fecha, hora, ae, r1, r4";
                strSql += " from " + (temporal ? "cont.eer_curvas_cuarto_horarias_tmp" : "cont.eer_curvas_cuarto_horarias")
                    + " WHERE cups20 = '" + cups20 + "'"                
                    + " and (fecha >= '" + fd.ToString("yyyy-MM-dd")
                    + "' and fecha <= '" + fh.ToString("yyyy-MM-dd") + "')"                    
                    + " ORDER BY fecha, hora";

                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    if (r["ae"] != System.DBNull.Value)
                    {
                        numeroPeriodosCuartoHorarios++;
                        parcialPeriodosHorarios++;
                        totalEnergiaActiva = totalEnergiaActiva + Convert.ToDouble(r["ae"]);

                        if(fd >= new DateTime(2022, 01, 01))
                        {
                            if (r["r1"] != System.DBNull.Value && r["r4"] != System.DBNull.Value)
                            {
                                totalEnergiaReactiva = totalEnergiaReactiva + (Convert.ToDouble(r["r1"]) - Convert.ToDouble(r["r4"]));
                                energiaReactiva = energiaReactiva + (Convert.ToDouble(r["r1"]) - Convert.ToDouble(r["r4"]));
                            }
                        }
                        else
                        {
                            if (r["r1"] != System.DBNull.Value)
                            {
                                totalEnergiaReactiva = totalEnergiaReactiva + Convert.ToDouble(r["r1"]);
                                energiaReactiva = energiaReactiva + Convert.ToDouble(r["r1"]);
                            }
                        }

                                     

                        hora = Convert.ToInt32(r["hora"].ToString().Substring(0,2));
                        minutos = Convert.ToInt32(r["hora"].ToString().Substring(3, 2));

                        curvaCuartoHorariaDias[numeroPeriodosCuartoHorarios] = Convert.ToDateTime(r["fecha"]);
                        curvaCuartoHorariaDias[numeroPeriodosCuartoHorarios] =
                            curvaCuartoHorariaDias[numeroPeriodosCuartoHorarios].AddHours(hora).AddMinutes(minutos);

                        curvaCuartoHorariaActiva[numeroPeriodosCuartoHorarios] = Convert.ToDouble(r["ae"]);

                        curvaCuartoHorariaReactiva_R1[numeroPeriodosCuartoHorarios] = Convert.ToDouble(r["r1"]);
                        curvaCuartoHorariaReactiva_R4[numeroPeriodosCuartoHorarios] = Convert.ToDouble(r["r4"]);

                        if (fd >= new DateTime(2022, 01, 01))                        
                            curvaCuartoHorariaReactiva[numeroPeriodosCuartoHorarios] =
                                (Convert.ToDouble(r["r1"]) - Convert.ToDouble(r["r4"]));
                        
                            
                        else
                            curvaCuartoHorariaReactiva[numeroPeriodosCuartoHorarios] = Convert.ToDouble(r["r1"]);

                        curvaCuartoHorariaPotencias[numeroPeriodosCuartoHorarios] = Convert.ToDouble(r["ae"]) * 4;
                        energiaActiva = energiaActiva + Convert.ToDouble(r["ae"]);
                        

                        if (parcialPeriodosHorarios == 4)
                        {
                            numeroPeriodosHorarios++;                            
                            curvaHorariaActiva[numeroPeriodosHorarios] = energiaActiva;
                            curvaHorariaReactiva[numeroPeriodosHorarios] = energiaReactiva;
                            curvaHorariaDias[numeroPeriodosHorarios] = Convert.ToDateTime(r["fecha"]);
                            curvaHorariaDias[numeroPeriodosHorarios] =
                                curvaHorariaDias[numeroPeriodosHorarios].AddHours(hora);

                            energiaActiva = 0;
                            energiaReactiva = 0;
                            parcialPeriodosHorarios = 0;
                        }

                    }
                    
                }

                db.CloseConnection();

                curvaCompleta = numeroPeriodosCuartoHorarios == numPeriodosMedidaCuartoHorario;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }


        private void GetCurvaHorizontal(List<string> lista_cups22, DateTime fd, DateTime fh, string estado)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            

            int numeroPeriodos = 0;
            int pt = 0;
            DateTime fechaHora = new DateTime();

            try
            {                
                calendarios.UtilidadesCalendario util = new calendarios.UtilidadesCalendario();

                #region Query

                strSql = "SELECT cups22, fecha, estado, version, totala, totalr";
                // HORA ACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,a" + j;

                // HORA REACTIVA
                for (int j = 1; j <= 25; j++)
                    strSql += " ,r" + j;              

                // CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 100; j++)                    
                    strSql += " ,v" + j;
                    

                strSql += " from cont.eer_cc "
                    + " WHERE cups22 in ('" + lista_cups22[0] + "'";

                for (int i = 1; i < lista_cups22.Count; i++)
                    strSql += ",'" + lista_cups22[i] + "'";

                strSql += ") and (fecha >= '" + fd.ToString("yyyy-MM-dd")
                    + "' and fecha <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and estado = '" + estado + "'"
                    + " ORDER BY cups22, fecha;";

                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.medida.CurvaDeCarga c = new EndesaEntity.medida.CurvaDeCarga();


                    fechaHora = Convert.ToDateTime(r["fecha"]);
                    numeroPeriodos = util.NumPeriodosHorarios(fechaHora);

                    c.fecha = fechaHora;
                    c.numPeriodosHorarios = numeroPeriodos;
                    c.numPeriodosCuartoHorarios = c.numPeriodosHorarios * 4;
                    
                    for(int i = 1; i < 26; i++)
                    {
                        if(r["a" + i] != System.DBNull.Value)
                            c.horaria_activa[i-1] = Convert.ToDouble(r["a" + i]);
                    }

                     
                    for (int i = 1; i < 101; i++)
                    {                        
                        if (r["v" + i] != System.DBNull.Value)                        
                            c.cuartohoraria_activa[i - 1] = Convert.ToDouble(r["v" + i]) / 4;                       
                            
                    }


                    
                    for (int i = 1; i <= c.numPeriodosCuartoHorarios; i++)
                    {
                        pt++;
                        if (r["v" + i] != System.DBNull.Value)
                        {
                            if(c.numPeriodosCuartoHorarios == 92 && i >= 9)
                                curvaCuartoHorariaActiva[pt] = curvaCuartoHorariaActiva[pt] + Convert.ToDouble(r["v" + (i + 4)]);
                            else
                                curvaCuartoHorariaActiva[pt] = curvaCuartoHorariaActiva[pt] + Convert.ToDouble(r["v" + i]);
                        }
                         
                    }



                    EndesaEntity.medida.DicCurva o;
                    if (!dic.TryGetValue(r["cups22"].ToString().Substring(0, 20), out o))
                    {
                        EndesaEntity.medida.DicCurva d = new EndesaEntity.medida.DicCurva();
                        d.dic.Add(c.fecha, c);
                        dic.Add(r["cups22"].ToString().Substring(0, 20), d);
                    }
                    else
                    {
                        EndesaEntity.medida.CurvaDeCarga x;
                        if (!o.dic.TryGetValue(c.fecha, out x))
                        {
                            o.dic.Add(c.fecha, c);
                        }
                    }

                    lista.Add(c);


                }

                db.CloseConnection();
                

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);             
            }


        }

        public bool CurvaCompleta(string cups20)
        {
            Dictionary<string, EndesaEntity.medida.DicCurva> matches =
                dic.Where(z => z.Key == cups20).ToDictionary(z => z.Key, z => z.Value);

            foreach(KeyValuePair<string, EndesaEntity.medida.DicCurva> p in matches)
            {
                return p.Value.dic.Count() == ((_fh - _fd).TotalDays + 1);
            }
            return false;
        }
                

        public List<EndesaEntity.medida.CurvaDeCarga> GetCurva(string cups13, DateTime fd, DateTime fh)
        {
            bool firstOnly = true;
            List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
            Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> d = new Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga>();

            Dictionary<string, EndesaEntity.medida.DicCurva> matches =
                dic.Where(z => z.Key == cups13).ToDictionary(z => z.Key, z => z.Value);

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

        public Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> GetCurva(string cups20)
        {
            EndesaEntity.medida.DicCurva o;
            Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> d
                = new Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga>();

            if (dic.TryGetValue(cups20, out o))
            {
                foreach (KeyValuePair<DateTime, EndesaEntity.medida.CurvaDeCarga> p in o.dic)
                    d.Add(p.Key, p.Value);

                return d;
            }
            else
                return null;
            
        }
        public List<double> GetCurvaVertical(string cups20)
        {
            calendarios.UtilidadesCalendario utilFecha = new calendarios.UtilidadesCalendario();
            List<double> l = new List<double>();
            List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
            Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> d = new Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga>();

            Dictionary<string, EndesaEntity.medida.DicCurva> matches =
                dic.Where(z => z.Key == cups20).ToDictionary(z => z.Key, z => z.Value);

            foreach (KeyValuePair<string, EndesaEntity.medida.DicCurva> p in matches)
                foreach (KeyValuePair<DateTime, EndesaEntity.medida.CurvaDeCarga> pp in p.Value.dic)
                    for (int i = 0; i < utilFecha.NumPeriodosCuartoHorarios(pp.Value.fecha); i++)
                        l.Add(pp.Value.cuartohoraria_activa[i]);



            return l;
        }

        public double TotalActiva(string cups20, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.medida.CurvaDeCarga> lista = GetCurva(cups20, fd, fh);
            if (lista != null)
                total = lista.Sum(z => z.total_activa);

            return total;
        }

        public double TotalReactiva(string cups20, DateTime fd, DateTime fh)
        {
            double total = 0;
            List<EndesaEntity.medida.CurvaDeCarga> lista = GetCurva(cups20, fd, fh);
            if (lista != null)
                total = lista.Sum(z => z.total_reactiva);

            return total;
        }

        //private double[] CargaExcesos(EndesaBusiness.calendarios.Calendario cal, EndesaEntity.punto_suministro.PuntoSuministro ps)
        //{
        //    int pt = 0;
        //    double[] vectorExcesos;
        //    double potencia_contratada = 0;
        //    double diferencia = 0;

        //    vectorExcesos = new double[(numPeriodosMedidaHorario + 1) * 4];

        //    for (int i = 1; i <= numPeriodosMedidaCuartoHorario; i++)
        //    {
        //        pt = cal.vectorPeriodosTarifariosCuartoHorarios[i];
        //        potencia_contratada = ps.potecias_contratadas[pt];

        //        if (curvaCuartoHorariaPotencias[i] > potencia_contratada)
        //        {
        //            diferencia = curvaCuartoHorariaPotencias[i] - potencia_contratada;
        //            vectorExcesos[i] = vectorExcesos[i] + diferencia;
        //            vectorExcesos[i] = (vectorExcesos[i] * vectorExcesos[i]);
        //        }


        //    }
        //    return vectorExcesos;
        //}


    }
}

