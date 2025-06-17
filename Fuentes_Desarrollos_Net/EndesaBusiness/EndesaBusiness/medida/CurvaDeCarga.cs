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
    public class CurvaDeCarga
    {
        List<EndesaEntity.medida.CurvaDeCarga> lista_curvas;

        public double[] energia_activa_por_periodo { get; set; }
        public double[] potencia_maxima_por_periodo { get; set; }

        public CurvaDeCarga(EndesaEntity.punto_suministro.PuntoSuministro ps, 
                            EndesaBusiness.calendarios.Calendario calendario,
                            DateTime fd, DateTime fh)
        {
            #region Inicializaciones

            lista_curvas = new List<EndesaEntity.medida.CurvaDeCarga>();
            
            energia_activa_por_periodo = new double[ps.tarifa.numPeriodosTarifarios];
            potencia_maxima_por_periodo = new double[ps.tarifa.numPeriodosTarifarios];
            CargaCurva(calendario, ps, fd, fh, "F");
            #endregion

        }


        private void CargaCurva(EndesaBusiness.calendarios.Calendario calendario,
            EndesaEntity.punto_suministro.PuntoSuministro ps,
             DateTime fd, DateTime fh, string estado)
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
                List<EndesaEntity.medida.PuntoMedida> lista_cups22= ps.lista_puntos_medida_principales;
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

                    for (int i = 1; i < 26; i++)
                    {
                        if (r["a" + i] != System.DBNull.Value)
                        {
                            c.horaria_activa[i - 1] = Convert.ToDouble(r["a" + i]);
                            c.total_activa = c.total_activa + Convert.ToDouble(r["a" + i]);
                        }
                            
                        
                    }


                    for (int i = 1; i < 101; i++)
                    {
                        if (r["v" + i] != System.DBNull.Value)
                            c.cuartohoraria_activa[i - 1] = Convert.ToDouble(r["v" + i]) / 4;

                    }

                    for (int i = 1; i <= c.numPeriodosCuartoHorarios; i++)
                    {
                        
                        if (r["v" + i] != System.DBNull.Value)
                        {
                            if (c.numPeriodosCuartoHorarios == 92 && i >= 9)
                                
                                c.potencias_cuartohorarias[i] = c.potencias_cuartohorarias[i] + Convert.ToDouble(r["v" + (i + 4)]);
                            else
                                c.potencias_cuartohorarias[i] = c.potencias_cuartohorarias[i] + Convert.ToDouble(r["v" + i]);
                        }

                    }





                    lista_curvas.Add(c);


                }

                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GetCurva", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }
    }
}
