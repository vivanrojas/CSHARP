using EndesaBusiness.medida.Redshift;
using EndesaBusiness.servidores;
using EndesaBusiness.utilidades;
using EndesaEntity;
using EndesaEntity.medida;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class CurvaResumenBI : EndesaEntity.medida.CurvaResumenTabla
    {
        public Dictionary<string, List<EndesaEntity.medida.CurvaResumenTabla>> dic { get; set; }

        public CurvaResumenBI()
        {
            dic = new Dictionary<string, List<CurvaResumenTabla>>();
        }

        public CurvaResumenBI(List<string> lista_cups20, DateTime fd, DateTime fh)
        {
            Redshift.Estados_Curvas estados_curvas = new Redshift.Estados_Curvas();

            dic = new Dictionary<string, List<EndesaEntity.medida.CurvaResumenTabla>>();

            BuscaCurvaResumen(lista_cups20, fd, fh, estados_curvas.estados_facturados);
            BuscaCurvaResumen(lista_cups20, fd, fh, estados_curvas.estados_registrados);
        }
              

        private void BuscaCurvaResumen(List<string> lista_cups20, DateTime fd, DateTime fh, List<string> lista_estados)
        {

            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            int j = 0;

            try
            {

                for (int i = 0; i < lista_cups20.Count;   i++)
                {
                    j++;
                    if (firstOnly)
                    {
                        sb.Append("select r.cd_cups_ext, r.cd_cups_ext_20, r.fh_fact_desde, r.fh_fact_hasta,");
                        sb.Append(" r.cd_estado_resumen, r.de_estado_resumen, r.nm_cons_total_act, r.nm_cons_total_reac,");
                        sb.Append(" r.nm_curvas_recibidas, r.de_tp_fuente_horaria, r.de_tp_fuente_cuartoh");
                        sb.Append(" from metra_owner.t_ed_h_rcurvas r");                        
                        sb.Append(" where r.cd_cups_ext in (");
                        sb.Append("'").Append(lista_cups20[i]).Append("'");
                        firstOnly = false;
                    }
                    else
                        sb.Append(",'").Append(lista_cups20[i]).Append("'");

                    if (j == 250)
                    {
                        sb.Append(")");
                        sb.Append(" and (r.fh_fact_desde >= ").Append(fd.ToString("yyyyMMdd"));
                        sb.Append(" and r.fh_fact_hasta <= ").Append(fh.ToString("yyyyMMdd")).Append(")");
                        sb.Append(" and r.cd_estado_resumen in ('").Append(lista_estados[0]).Append("'");

                        for (int y = 1; y < lista_estados.Count; y++)
                            sb.Append(",'").Append(lista_estados[y]).Append("'");

                        sb.Append(") order by r.cd_cups_ext_20, r.fh_fact_desde, r.cd_estado_resumen");
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
                    sb.Append(" and (r.fh_fact_desde >= ").Append(fd.ToString("yyyyMMdd"));
                    sb.Append(" and r.fh_fact_hasta <= ").Append(fh.ToString("yyyyMMdd")).Append(")");
                    sb.Append(" and r.cd_estado_resumen in ('").Append(lista_estados[0]).Append("'");

                    for (int y = 1; y < lista_estados.Count; y++)
                        sb.Append(",'").Append(lista_estados[y]).Append("'");

                    sb.Append(") order by r.cd_cups_ext_20, r.fh_fact_desde, r.cd_estado_resumen");
                    j = 0;
                    firstOnly = true;
                    RunQuery(sb.ToString());
                    sb = null;
                    sb = new StringBuilder();
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "CurvaResumenBI.BuscaCurvaResumen",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        private void RunQuery(string q)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;


            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(q, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.medida.CurvaResumenTabla re = new EndesaEntity.medida.CurvaResumenTabla();
                re.cups22 = r["cd_cups_ext"].ToString();
                re.cups20 = r["cd_cups_ext_20"].ToString();

                re.fecha_desde = new DateTime(Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(0, 4)),
                       Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(4, 2)),
                       Convert.ToInt32(r["fh_fact_desde"].ToString().Substring(6, 2)));

                re.fecha_hasta = new DateTime(Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(0, 4)),
                       Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(4, 2)),
                       Convert.ToInt32(r["fh_fact_hasta"].ToString().Substring(6, 2)));

                re.estado = r["de_estado_resumen"].ToString();

                if (r["nm_cons_total_act"] != System.DBNull.Value) 
                    re.activa = Convert.ToDouble(r["nm_cons_total_act"]);
                if (r["nm_cons_total_reac"] != System.DBNull.Value)
                    re.reactiva = Convert.ToDouble(r["nm_cons_total_reac"]);
                if(r["nm_curvas_recibidas"] != System.DBNull.Value)
                    re.dias = Convert.ToInt32(r["nm_curvas_recibidas"]);

                if (r["de_tp_fuente_horaria"] != System.DBNull.Value)
                    re.fuente = r["de_tp_fuente_horaria"].ToString();
                else if(r["de_tp_fuente_cuartoh"] != System.DBNull.Value)
                    re.fuente = r["de_tp_fuente_cuartoh"].ToString();

                List<EndesaEntity.medida.CurvaResumenTabla> rr;
                
                if (!dic.TryGetValue(re.cups22, out rr))
                {
                    rr = new List<EndesaEntity.medida.CurvaResumenTabla>();
                    rr.Add(re);
                    dic.Add(re.cups22, rr);
                }
                else
                {
                    GetCurva(re.cups22, re.fecha_desde, re.fecha_hasta);
                    //if (!this.existe_curva) 18/06/2024 GUS: modificamos condición
                    if (!this.existe_curva)
                    {
                        rr.Add(re);
                    }
                    else if (re.estado == "REGISTRADA" && (this.estado == "SEGUNDA RECEPCION" || this.estado == "FACTURADA"))
                    {
                        re.estado = "SEGUNDA RECEPCION";
                        rr.Add(re);
                    }
                    //else //DEBE SER REGISTRADA, PUEDE HABER MÁS DE UNA, POR EJEMPLO PERIODOS PARTIDOS 1-14 + 15-31
                    //{
                    //    rr.Add(re);
                    //}
                }
                


               
                    
                

            }

            db.CloseConnection();
        }

        public void GetCurva(string cups22, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.medida.CurvaResumenTabla> lista;
            if (dic.TryGetValue(cups22, out lista))
            {
                List<EndesaEntity.medida.CurvaResumenTabla> sublista = lista.Where(z => (z.fecha_desde >= fd && z.fecha_hasta <= fh)).ToList();
                if (sublista.Count > 0)
                    for (int i = 0; i < sublista.Count(); i++)
                    {
                        this.existe_curva = true;
                        if (i == 0)
                        {
                            this.estado = sublista[i].estado;
                            this.fecha_desde = sublista[i].fecha_desde;
                            this.fecha_hasta = sublista[i].fecha_hasta;
                            this.activa = sublista[i].activa;
                            this.reactiva = sublista[i].reactiva;
                            this.fuente = sublista[i].fuente;
                            this.dias = sublista[i].dias;
                        }
                        else if(this.estado == sublista[i].estado)
                        {
                            this.estado = sublista[i].estado;
                            this.fecha_desde = sublista[i].fecha_desde;
                            this.fecha_hasta = sublista[i].fecha_hasta;
                            this.activa += sublista[i].activa;
                            this.reactiva += sublista[i].reactiva;
                            this.fuente = sublista[i].fuente;
                            this.dias += sublista[i].dias;
                        }
                        else
                        {
                            this.estado = sublista[i].estado;
                            this.fecha_desde = sublista[i].fecha_desde;
                            this.fecha_hasta = sublista[i].fecha_hasta;
                            this.activa = sublista[i].activa;
                            this.reactiva = sublista[i].reactiva;
                            this.fuente = sublista[i].fuente;
                            this.dias = sublista[i].dias;
                        }

                    }
                else
                {
                    this.existe_curva = false;

                    this.dias = 0;
                }     

            }            
            else
            {
                this.existe_curva = false;

                this.dias = 0;


            }
        }

    }
}
