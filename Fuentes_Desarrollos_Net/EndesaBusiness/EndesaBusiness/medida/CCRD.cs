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
    public class CCRD
    {

        public Dictionary<string, EndesaEntity.medida.DicCurva> dic;
        public List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
        DateTime _fd = new DateTime();
        DateTime _fh = new DateTime();




        public CCRD(List<string> lista_cups13, DateTime fd, DateTime fh)
        {
            _fd = fd;
            _fh = fh;

            // Determinamos si la lista de cups es de cups13 o cups20
            dic = new Dictionary<string, EndesaEntity.medida.DicCurva>();
            medida.PuntosMedidaPrincipalesVigentes pmed = new PuntosMedidaPrincipalesVigentes(lista_cups13);
            GetCurvaRedShift(pmed.lista_cups15, fd, fh, "R");
        }

        public CCRD(List<string> lista_cups20, DateTime fd, DateTime fh, string estado)
        {
            _fd = fd;
            _fh = fh;

            // Determinamos si la lista de cups es de cups13 o cups20
            dic = new Dictionary<string, EndesaEntity.medida.DicCurva>();
            medida.PuntosMedidaPrincipalesVigentes pmed = new PuntosMedidaPrincipalesVigentes(lista_cups20);
            lista = new List<EndesaEntity.medida.CurvaDeCarga>();
            GetCurvaRedShift(pmed.lista_cups15, fd, fh, estado);            
            
        }


        private void GetCurvaRedShift(List<string> lista_cups15, DateTime fd, DateTime fh, string estado)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql;

            int numeroPeriodos = 0;
            int year = 0;
            int month = 0;
            int day = 0;
            int y = 0;
            int z = 0;
            DateTime fechaHora = new DateTime();

            try
            {                
                calendarios.UtilidadesCalendario util = new calendarios.UtilidadesCalendario();                

                #region Query

                strSql = "SELECT cd_punto_med as cups15, cd_cups_ext as cups22, fh_lect_registro as fecha,  cd_estado_curva as estado ";
                //// HORA ACTIVA
                //for (int j = 1; j <= 25; j++)
                //    strSql += " ,NM_ENER_AC_H" + j + " as a" + j;

                //// HORA REACTIVA
                //for (int j = 1; j <= 25; j++)
                //    strSql += " ,NM_ENER_R1_H" + j + " as r" + j;

                //// FUENTE HORA ACTIVA
                //for (int j = 1; j <= 25; j++)
                //    strSql += " ,CD_FUENTE_HOR_AC_H" + j + " as fh" + j;

                // CUARTOHORARIA ACTIVA
                for (int j = 1; j <= 25; j++)
                    for (int w = 1; w <= 4; w++)
                    {
                        z++;
                        strSql += " ,NM_POT_AC_H" + j + "_CUAD" + w + " as v" + z;
                    }


                // FUENTE CUARTOHORARIA ACTIVA
                //for (int j = 1; j <= 25; j++)
                //    strSql += " ,CD_FUENTE_CUARTH_AC_H" + j + " as fch" + j;



                strSql += " from metra_owner.t_ed_h_curvas "
                    + " WHERE cd_punto_med in ('" + lista_cups15[0] + "'";

                for (int i = 1; i < lista_cups15.Count; i++)
                    strSql += ",'" + lista_cups15[i] + "'";

                strSql += ") and (fh_lect_registro >= " + fd.ToString("yyyyMMdd")
                    + " and fh_lect_registro <= " + fh.ToString("yyyyMMdd") + ")"
                    + " and cd_estado_curva = '" + estado + "'"
                    + " ORDER BY cd_punto_med, fh_lect_registro;";
               
                #endregion

                db = new RedShiftServer(RedShiftServer.Entornos.PROD);
                command = new OdbcCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.medida.CurvaDeCarga c = new EndesaEntity.medida.CurvaDeCarga();

                    year = Convert.ToInt32(r["FECHA"].ToString().Substring(0, 4));
                    month = Convert.ToInt32(r["FECHA"].ToString().Substring(4, 2));
                    day = Convert.ToInt32(r["FECHA"].ToString().Substring(6, 2));

                    fechaHora = new DateTime(year, month, day, 0, 0, 0);
                    numeroPeriodos = util.NumPeriodosHorarios(fechaHora);

                    c.fecha = fechaHora;
                    c.numPeriodosHorarios = numeroPeriodos;
                    c.numPeriodosCuartoHorarios = c.numPeriodosHorarios * 4;
                    

                    for (int i = 1; i < 101; i++)
                    {
                        if (r["v" + i] != System.DBNull.Value)
                            c.cuartohoraria_activa[i - 1] = Convert.ToDouble(r["v" + i]) / 4;
                    }


                    EndesaEntity.medida.DicCurva o;
                    if (!dic.TryGetValue(r["cups15"].ToString().Substring(0,13), out o))
                    {
                        EndesaEntity.medida.DicCurva d = new EndesaEntity.medida.DicCurva();
                        d.dic.Add(c.fecha, c);
                        dic.Add(r["cups15"].ToString().Substring(0, 13), d);
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

        public bool CurvaCompleta(string cups13)
        {
            Dictionary<string, EndesaEntity.medida.DicCurva> matches =
                dic.Where(z => z.Key == cups13).ToDictionary(z => z.Key, z => z.Value);

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
        public List<double> GetCurvaVertical(string cups13)
        {
            calendarios.UtilidadesCalendario utilFecha = new calendarios.UtilidadesCalendario();
            List<double> l = new List<double>();
            List<EndesaEntity.medida.CurvaDeCarga> lista = new List<EndesaEntity.medida.CurvaDeCarga>();
            Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> d = new Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga>();

            Dictionary<string, EndesaEntity.medida.DicCurva> matches =
                dic.Where(z => z.Key == cups13).ToDictionary(z => z.Key, z => z.Value);

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

    }
}
