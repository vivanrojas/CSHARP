using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.calendarios
{
    public class Calendario
    {
        public int numPeriodosMedidaHorario { get; set; }
        public int numPeriodosMedidaCuartoHorario { get; set; }
        public double[] vectorPeriodoMedida { get; set; }
        public int[] vectorPeriodosTarifarios { get; set; }
        public DateTime[] vectorFechas { get; set; }
        public DateTime[] vectorFechasCuartoHorarias { get; set; }
        public int[] vectorPeriodosTarifariosCuartoHorarios { get; set; }
        public DateTime[] vectorDias { get; set; }
        public int[] vectorHoras { get; set; }
        public DateTime[] vectorHorasCuartoHorarias { get; set; }

        List<DateTime> lf = new List<DateTime>(); // lista festivos
        List<Calendario> lc = new List<Calendario>(); // lista calendarios
        Dictionary<string, string> dic_territorios = new Dictionary<string, string>();

        UtilidadesCalendario utilCal = new UtilidadesCalendario();
        EndesaEntity.punto_suministro.PuntoSuministro ps;
        CargaCalendarios cal;
        EndesaBusiness.punto_suministro.Territorios territorio;

        public double[] energiaActivaPorPeriodo { get; set; }
        public double[] energiaReactivaPorPeriodo { get; set; }
        public double[] potenciasMaximasRegistradas { get; set; }
        public double[] potenciasaFacturar { get; set; }
        public double[] potenciasAbsorbilbles { get; set; }


        public Calendario(DateTime fd, DateTime fh)
        {
            numPeriodosMedidaHorario = (Convert.ToInt32((fh - fd).TotalDays + 1) * 24) + utilCal.CorreccionCambioHorario(fd, fh);            
            vectorHoras = new int[numPeriodosMedidaHorario + 1];
            numPeriodosMedidaCuartoHorario = numPeriodosMedidaHorario * 4;

            vectorPeriodosTarifarios = new int[numPeriodosMedidaHorario + 1];
            vectorFechas = new DateTime[numPeriodosMedidaHorario + 1];
            vectorFechasCuartoHorarias =  new DateTime[numPeriodosMedidaCuartoHorario + 1];
            vectorPeriodoMedida = new double[numPeriodosMedidaHorario + 1];
            vectorPeriodosTarifariosCuartoHorarios = new int[numPeriodosMedidaCuartoHorario + 1];
            vectorDias = new DateTime[Convert.ToInt32((fh - fd).TotalDays + 1)];
            vectorHoras = new int[numPeriodosMedidaHorario + 1];
            vectorHorasCuartoHorarias = new DateTime[numPeriodosMedidaCuartoHorario + 1];
            territorio = new punto_suministro.Territorios();

            // CargaTerritorios();
            CargaListaFestivos(fd, fh);
            cal = new CargaCalendarios(fd, fh);
            ps = new EndesaEntity.punto_suministro.PuntoSuministro();
        }

        public Calendario(DateTime fd, DateTime fh, string tarifa, string territorio)
        {

            int j = 1;
            numPeriodosMedidaHorario = (Convert.ToInt32((fh - fd).TotalDays + 1) * 24) + utilCal.CorreccionCambioHorario(fd, fh);
            vectorHoras = new int[numPeriodosMedidaHorario + 1];
            numPeriodosMedidaCuartoHorario = numPeriodosMedidaHorario * 4;

            vectorPeriodosTarifarios = new int[numPeriodosMedidaHorario + 1];
            vectorFechas = new DateTime[numPeriodosMedidaHorario + 1];
            vectorPeriodoMedida = new double[numPeriodosMedidaHorario + 1];
            vectorPeriodosTarifariosCuartoHorarios = new int[numPeriodosMedidaCuartoHorario + 1];
            vectorDias = new DateTime[Convert.ToInt32((fh - fd).TotalDays + 1)];
            vectorFechasCuartoHorarias = new DateTime[numPeriodosMedidaCuartoHorario + 1];
            vectorHorasCuartoHorarias = new DateTime[numPeriodosMedidaCuartoHorario + 1];


            CargaListaFestivos(fd, fh);
            cal = new CargaCalendarios(fd, fh);


            for (DateTime d = fd; d <= fh; d = d.AddDays(1))
            {
                for (int i = 1; i <= utilCal.NumPeriodosHorarios(d); i++)
                {

                    vectorPeriodoMedida[j] = i;
                    vectorHoras[j] = i;
                    vectorFechas[j] = d;
                    vectorPeriodosTarifarios[j] = cal.GetPeriodoTarifa(tarifa, territorio, d, i);

                    
                    vectorFechasCuartoHorarias[(j * 4) - 3] = d;
                    vectorFechasCuartoHorarias[(j * 4) - 2] = d;
                    vectorFechasCuartoHorarias[(j * 4) - 1] = d;
                    vectorFechasCuartoHorarias[(j * 4)] = d;

                    vectorHorasCuartoHorarias[(j * 4) - 3] = d.AddHours(i - 1);
                    vectorHorasCuartoHorarias[(j * 4) - 2] = vectorHorasCuartoHorarias[(j * 4) - 3].AddMinutes(15);
                    vectorHorasCuartoHorarias[(j * 4) - 1] = vectorHorasCuartoHorarias[(j * 4) - 2].AddMinutes(15);
                    vectorHorasCuartoHorarias[(j * 4)] = vectorHorasCuartoHorarias[(j * 4) - 1].AddMinutes(15);



                    vectorPeriodosTarifariosCuartoHorarios[(j * 4) - 3] = vectorPeriodosTarifarios[j];
                    vectorPeriodosTarifariosCuartoHorarios[(j * 4) - 2] = vectorPeriodosTarifarios[j];
                    vectorPeriodosTarifariosCuartoHorarios[(j * 4) - 1] = vectorPeriodosTarifarios[j];
                    vectorPeriodosTarifariosCuartoHorarios[(j * 4)] = vectorPeriodosTarifarios[j];
    

                    j++;
                }

              
            }


        }


        public void CalculaDatosMedida(EndesaEntity.punto_suministro.PuntoSuministro vps, 
            double[] cuartoHorariaActiva, double[] cuartoHorariaReactiva,
            double[] cuartoHorariaPotencia, DateTime[] cuartoHorariaFechaHora)
        {
            ps = vps;
            energiaActivaPorPeriodo = new double[ps.tarifa.numPeriodosTarifarios + 1];
            energiaReactivaPorPeriodo = new double[ps.tarifa.numPeriodosTarifarios + 1];
            potenciasMaximasRegistradas = new double[ps.tarifa.numPeriodosTarifarios + 1];
            potenciasaFacturar = new double[ps.tarifa.numPeriodosTarifarios + 1];
            potenciasAbsorbilbles = new double[ps.tarifa.numPeriodosTarifarios + 1];

            CargaEnergiaActivaPorPeriodo(cuartoHorariaActiva, cuartoHorariaReactiva, cuartoHorariaFechaHora);
            CargaPotenciasMaxPorPeriodo(cuartoHorariaPotencia);

        }


        


        private void CargaEnergiaActivaPorPeriodo(double[] cuartoHorariaActiva, double[] cuartoHorariaReactiva, DateTime[] cuartoHorariaFechaHora)
        {
            // DateTime dia = new DateTime();
            int periodo = 0;
            int j = 0;
            string tarifa = ps.tarifa.tarifa;
            string _territorio = territorio.GetTerritorio(ps.direccion.codigo_postal);
            int parcialPeriodosHorarios = 0;
            int numeroPeriodosHorarios = 0;
            

            try
            {
                

                for (int i = 1; i < cuartoHorariaActiva.Count(); i++)
                {

                    parcialPeriodosHorarios++;
                    periodo = cal.GetPeriodoTarifa(tarifa, _territorio, cuartoHorariaFechaHora[i], (cuartoHorariaFechaHora[i].Hour) + 1);
                    vectorPeriodosTarifariosCuartoHorarios[i] = periodo;

                    energiaActivaPorPeriodo[periodo] = energiaActivaPorPeriodo[periodo] + cuartoHorariaActiva[i];
                    energiaReactivaPorPeriodo[periodo] = energiaReactivaPorPeriodo[periodo] + cuartoHorariaReactiva[i];

                    if (parcialPeriodosHorarios == 4)
                    {
                        numeroPeriodosHorarios++;
                        vectorPeriodosTarifarios[numeroPeriodosHorarios] = periodo;
                        parcialPeriodosHorarios = 0;
                    }
                                       

                }
                
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            


        }

        



        //private void CargaEnergiaActivaPorPeriodo()
        //{
        //    DateTime dia = new DateTime();
        //    int periodo = 0;
        //    int j = 0;
        //    string tarifa = ps.tarifa.tarifa;
        //    string _territorio = territorio.GetTerritorio(ps.direccion.codigo_postal);
        //    double totalEnergia = 0;

        //    foreach (KeyValuePair<DateTime, EndesaEntity.medida.CurvaDeCarga> p in ps.dic_cc)
        //    {

        //        totalEnergia += p.Value.total_activa;
        //        dia = p.Key;
        //        for(int i = 1; i <= p.Value.numPeriodosHorarios; i++)
        //        {
        //            j++;
        //            periodo = cal.GetPeriodoTarifa(tarifa, _territorio, dia, );
        //            // Vectores Tarifarios
        //            VectorPeriodosTarifarios[j] = periodo;
        //            vectorPeriodosTarifariosCuartoHorarios[(j * 4) - 3] = periodo;
        //            vectorPeriodosTarifariosCuartoHorarios[(j * 4) - 2] = periodo;
        //            vectorPeriodosTarifariosCuartoHorarios[(j * 4) - 1] = periodo;
        //            vectorPeriodosTarifariosCuartoHorarios[(j * 4)] = periodo;


        //            if (p.Value.numPeriodosHorarios == 23 && i >= 3)
        //                energiaActivaPorPeriodo[periodo] = energiaActivaPorPeriodo[periodo] + p.Value.horaria_activa[i];
        //            else
        //                energiaActivaPorPeriodo[periodo] = energiaActivaPorPeriodo[periodo] + p.Value.horaria_activa[i-1];
        //        }

        //    }

        //}       

        private void CargaPotenciasMaxPorPeriodo(double[] cc)
        {
            double pot_85 = 0;
            double pot_105 = 0;
            double potMax = 0;
            int j = 0;
            int periodo = 0;
            double pc_85 = 0;
            double pd = 0;
            double pc = 0;
            double pa = 0;


            
            
            for (int i = 1; i <= numPeriodosMedidaCuartoHorario; i++)
            {                
                periodo = vectorPeriodosTarifariosCuartoHorarios[i];
                if (potenciasMaximasRegistradas[periodo] < cc[i])
                    potenciasMaximasRegistradas[periodo] = cc[i];
            }
            
            
            if (ps.tarifa.tarifa == "3.1A" || ps.tarifa.tarifa == "3.0A")
            {
                for(int i = 1; i <= ps.tarifa.numPeriodosTarifarios; i++)
                {
                    pot_85 = ps.potecias_contratadas[i]  * 0.85;
                    pot_105 = ps.potecias_contratadas[i] * 1.05;
                    potMax = potenciasMaximasRegistradas[i];

                    if (potMax < pot_85)
                        potenciasaFacturar[i] = pot_85;
                    else if (potMax > pot_105)
                        potenciasaFacturar[i] = potMax + (2 * pot_105);
                    else
                        potenciasaFacturar[i] = potMax;

                }
            }else if(ps.tarifa.tarifa.Substring(0,1)  == "6" || ps.tarifa.tarifa == "3.0TD")
            {
                for (int i = 1; i <= ps.tarifa.numPeriodosTarifarios; i++)
                {
                    pc_85 = ps.potecias_contratadas[i] * 0.85;
                    pd = potenciasMaximasRegistradas[i];
                    pc = ps.potecias_contratadas[i];
                    pa = potenciasAbsorbilbles[i];                    

                    if (pc_85 >= pd)
                        potenciasaFacturar[i] = (0.1 * (pa - pc) + pc_85);                   
                    else
                        potenciasaFacturar[i] = (0.1 * (pa - pc) + pd);

                }
            }
        }

        
        private void CargaTerritorios()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "Select CodigoPostal, Territorio from fact_territorios";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    dic_territorios.Add(r["CodigoPostal"].ToString().Substring(0, 2), r["Territorio"].ToString());
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "ClaseCalendario - CargaTerritorios",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            }
        }

        private void CargaListaFestivos(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";            

            try
            {
                strSql = "select FechaFestivo from fact.fact_diasfestivos where"
                    + " FechaFestivo >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FechaFestivo <= '" + fh.ToString("yyyy-MM-dd") + "';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    lf.Add(Convert.ToDateTime(r["FechaFestivo"]));
                }

                r.Close();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la carga de festivos electricos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }




    }
}

