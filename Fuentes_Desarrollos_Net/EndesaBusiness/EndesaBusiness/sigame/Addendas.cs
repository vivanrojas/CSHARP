using EndesaBusiness.contratacion.gestionATRGas;
using EndesaBusiness.servidores;
using EndesaEntity.contratacion.gas;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.sigame
{
    public class Addendas
    {
        Dictionary<string, List<EndesaEntity.contratacion.gas.Addenda>> dic;

        public Addendas()
        {
            dic = Carga(DateTime.MinValue, DateTime.MinValue);
        }

        public Addendas(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }


        private Dictionary<string, List<EndesaEntity.contratacion.gas.Addenda>> Carga(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.contratacion.gas.Addenda>> d =
                new Dictionary<string, List<EndesaEntity.contratacion.gas.Addenda>>();

            try
            {
                strSql = "select T_SGM_M_ADDENDA_CTT_DIST.ID_PS, T_SGM_G_PS.CD_CUPS,"
                    + " T_SGM_P_TIPO_TARIFA.DE_TIPO_TARIFA as Tarifa_Peaje, T_SGM_P_PEAJE_DURACION.DE_PEAJE_DURACION as Duracion_Peaje,"
                    + " T_SGM_M_ADDENDA_CTT_DIST.NM_CDC as Qd,"
                    + " T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO, T_SGM_M_ADDENDA_CTT_DIST.FH_FIN"
                    + " from T_SGM_M_ADDENDA_CTT_DIST"
                    + " LEFT JOIN T_SGM_P_TIPO_TARIFA ON"
                    + " T_SGM_P_TIPO_TARIFA.ID_TIPO_TARIFA = T_SGM_M_ADDENDA_CTT_DIST.ID_TIPO_TARIFA"
                    + " LEFT JOIN T_SGM_P_PEAJE_DURACION ON"
                    + " T_SGM_P_PEAJE_DURACION.ID_PEAJE_DURACION = T_SGM_M_ADDENDA_CTT_DIST.ID_PEAJE_DURACION"
                    + " LEFT JOIN T_SGM_G_PS on"
                    + " T_SGM_G_PS.ID_PS = T_SGM_M_ADDENDA_CTT_DIST.ID_PS";

                if(fd != DateTime.MinValue && fh != DateTime.MinValue) 
                {
                    strSql += " WHERE T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO <= '" + fh.ToString("yyyy-MM-dd") + "' AND"
                        + " (T_SGM_M_ADDENDA_CTT_DIST.FH_FIN >= '" + fd.ToString("yyyy-MM-dd") + "' OR"
                        + " T_SGM_M_ADDENDA_CTT_DIST.FH_FIN is null)";

                }
                strSql += " ORDER BY T_SGM_G_PS.CD_CUPS, T_SGM_M_ADDENDA_CTT_DIST.FH_INICIO";

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.contratacion.gas.Addenda c = new Addenda();

                    c.id_ps = Convert.ToInt32(r["ID_PS"]);
                    c.cups = r["CD_CUPS"].ToString();

                    if (r["Tarifa_Peaje"] != System.DBNull.Value)
                        c.tarifa_peaje = r["Tarifa_Peaje"].ToString();

                    if (r["Duracion_Peaje"] != System.DBNull.Value)
                    {
                        if(r["Duracion_Peaje"].ToString() == "Larga")
                            c.duracion_peaje = "INDEFINIDO";
                        else
                            c.duracion_peaje = r["Duracion_Peaje"].ToString().ToUpper();
                    }
                        

                    if (r["Qd"] != System.DBNull.Value)
                        c.qd = Convert.ToDouble(r["Qd"]);

                    if (r["FH_INICIO"] != System.DBNull.Value)
                        c.fecha_desde = (Convert.ToDateTime(r["FH_INICIO"])).Date; // [03/04/2025 GUS] Establecemos la hora a 00:00:00

                    if (r["FH_FIN"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["FH_FIN"]);                    

                    List<EndesaEntity.contratacion.gas.Addenda> o;
                    if (!d.TryGetValue(c.cups, out o))
                    {
                        o = new List<Addenda>();
                        o.Add(c);
                        d.Add(c.cups, o);
                    }
                    else
                        o.Add(c);


                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                return null;
            }
        }

        public List<EndesaEntity.contratacion.gas.Addenda> GetAddendas(string cups)
        {
            List<EndesaEntity.contratacion.gas.Addenda> o;
            if (dic.TryGetValue(cups, out o))
                return o.OrderBy(z => z.fecha_desde).ToList();
            else
                return null;
            
        }

        public List<EndesaEntity.contratacion.gas.Addenda> GetAddendas(string cups, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.contratacion.gas.Addenda> o;
            if (dic.TryGetValue(cups, out o))
                return o.Where(z => z.fecha_desde <= fd && (z.fecha_hasta >= fh || z.fecha_hasta == DateTime.MinValue)).ToList();
            else
                return null;

        }

        public List<EndesaEntity.contratacion.gas.Addenda> GetAddendasReunibacionesSinFechaFin(string cups, DateTime fd)
        {
            List<EndesaEntity.contratacion.gas.Addenda> lista_addendas = new List<Addenda>();

            List<EndesaEntity.contratacion.gas.Addenda> o;
            if (dic.TryGetValue(cups, out o))
            {
                foreach(EndesaEntity.contratacion.gas.Addenda p in o)
                {
                    if(p.fecha_desde <= fd && p.fecha_hasta == DateTime.MinValue)
                        lista_addendas.Add(p);
                }

                return lista_addendas;
            }                
            else
                return null;

        }

        public List<EndesaEntity.contratacion.gas.Addenda> GetAddendasReunibacionesConFechaFin(string cups, DateTime fd)
        {
            List<EndesaEntity.contratacion.gas.Addenda> lista_addendas = new List<Addenda>();

            List<EndesaEntity.contratacion.gas.Addenda> o;
            if (dic.TryGetValue(cups, out o))
            {
                foreach (EndesaEntity.contratacion.gas.Addenda p in o)
                {
                    if (p.fecha_desde <= fd && p.fecha_hasta >= fd)
                        lista_addendas.Add(p);
                }

                return lista_addendas;
            }
            else
                return null;

        }

        public List<EndesaEntity.contratacion.gas.Addenda> GetAddendasReunibaciones(string cups, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.contratacion.gas.Addenda> o;
            if (dic.TryGetValue(cups, out o))
                return o.Where(z => z.fecha_desde <= fd && (z.fecha_hasta >= fd || z.fecha_hasta == DateTime.MinValue)).ToList();
            else
                return null;

        }

        public string UltimaTarifaAddendas(string cups)
        {
            string tarifa = null;
            List<EndesaEntity.contratacion.gas.Addenda> o;
            if (dic.TryGetValue(cups, out o))
                tarifa = o.Last().tarifa_peaje;

            return tarifa;
        }
        
        public bool ExisteCUPS(string cups) 
        {
            List<EndesaEntity.contratacion.gas.Addenda> o;
            return dic.TryGetValue(cups, out o);
        }

        public void ActualizaAddendas(Dictionary<string, List<ContratoGasDetalle>> dic_contratosDetalle)
        {

            EndesaBusiness.contratacion.gestionATRGas.ContratosGasDetalle cgd =
                new ContratosGasDetalle();

            List<EndesaEntity.contratacion.gas.Addenda> lista_addendas;

            foreach (KeyValuePair < string, List<ContratoGasDetalle>> p in dic_contratosDetalle)
            {
                lista_addendas = this.GetAddendas(p.Value[0].cups20);
                if(lista_addendas != null)                
                    for (int x = 0; x < lista_addendas.Count; x++)
                    {
                        if (lista_addendas[x].duracion_peaje.ToUpper() == "DIARIO")
                        {
                            for (DateTime j = lista_addendas[x].fecha_desde;
                                j <= lista_addendas[x].fecha_hasta; j = j.AddDays(1))
                            {
                                
                                    cgd.nif = p.Value[0].vatnum;
                                    cgd.qd = lista_addendas[x].qd;
                                    cgd.tarifa = lista_addendas[x].tarifa_peaje;
                                    cgd.cups20 = lista_addendas[x].cups;
                                    cgd.tipo = "DIARIO";
                                    cgd.fecha_inicio = j;
                                    cgd.fecha_fin = j;
                                    cgd.comentario = "ADDENDA SIGAME";
                                    cgd.Replace();
                                                            
                            }

                        }
                        else
                        {
                            
                            
                                cgd.nif = p.Value[0].vatnum;
                                cgd.qd = lista_addendas[x].qd;
                                cgd.tarifa = lista_addendas[x].tarifa_peaje;
                                cgd.cups20 = lista_addendas[x].cups;
                                cgd.tipo = lista_addendas[x].duracion_peaje;
                                cgd.fecha_inicio = lista_addendas[x].fecha_desde;
                                cgd.fecha_fin = lista_addendas[x].fecha_hasta;
                                cgd.comentario = "ADDENDA SIGAME";

                                if (cgd.Existe_ContratoDetalle_ConFinMundo(cgd.nif,
                                    cgd.cups20,
                                    cgd.tipo, cgd.fecha_inicio))

                                    cgd.Del(cgd.nif,
                                    cgd.cups20,
                                    cgd.tipo,
                                    cgd.fecha_inicio,
                                    new DateTime(4999, 12, 31).Date);

                                cgd.Replace();
                            
                        
                        }

                    }
            }




        }

        public Dictionary<string, List<ContratoGasDetalle>> CompletaInfo(Dictionary<string, List<ContratoGasDetalle>> dic_contratosDetalle, DateTime fd)
        {

           

            List<EndesaEntity.contratacion.gas.Addenda> lista_addendas;

            try
            {
                foreach (KeyValuePair<string, List<ContratoGasDetalle>> p in dic_contratosDetalle)
                {
                    // Buscamos todas las addendas del CUPS en SIGAME
                    lista_addendas = this.GetAddendas(p.Value[0].cups20);


                    // Para los contratos detalle que existen
                    // actualizamos datos con la addenda
                    foreach (ContratoGasDetalle pp in p.Value)
                    {

                        EndesaEntity.contratacion.gas.Addenda addenda = lista_addendas.
                            Find(z => z.fecha_desde == pp.fecha_inicio && z.duracion_peaje == pp.tipo);

                        if (addenda != null)
                        {
                            pp.tarifa = addenda.tarifa_peaje;

                            if(pp.fecha_fin != addenda.fecha_hasta)
                                pp.fecha_fin = addenda.fecha_hasta;

                            pp.qd = addenda.qd;
                            
                        }

                    }

                    // Añadimos las Addendas que no existan en contratos detalle

                    foreach (EndesaEntity.contratacion.gas.Addenda pp in lista_addendas)
                    {

                        //if(pp.fecha_desde <= fd && (pp.fecha_hasta >= fd || pp.fecha_hasta == DateTime.MinValue))
                        {
                           ContratoGasDetalle cgd = p.Value.
                           Find(z => z.fecha_inicio == pp.fecha_desde
                           && z.tipo == pp.duracion_peaje
                           && z.tarifa == pp.tarifa_peaje);

                            if (cgd == null)
                            {

                                ContratoGasDetalle c = new ContratoGasDetalle();
                                c.nif = p.Value[0].vatnum;
                                c.vatnum = p.Value[0].vatnum;
                                c.customer_name = p.Value[0].customer_name;
                                c.cups20 = p.Value[0].cups20;
                                c.fecha_inicio = pp.fecha_desde;
                                c.fecha_fin = pp.fecha_hasta;
                                c.tipo = pp.duracion_peaje;
                                c.qd = pp.qd;
                                c.tarifa = pp.tarifa_peaje;
                                p.Value.Add(c);
                            }
                        }

                       
                    }
                }

                return dic_contratosDetalle;
            }
            catch(Exception e)
            {
                return null;
            }

            
            
        }

    }
}
