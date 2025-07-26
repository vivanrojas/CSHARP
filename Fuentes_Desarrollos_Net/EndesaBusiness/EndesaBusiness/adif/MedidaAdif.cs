using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.adif
{
    public class MedidaAdif
    {
        public Dictionary<string, EndesaEntity.medida.AdifInformeMedida> dic_informe { get; set; }

        public medida.MedidaFunciones m { get; set; }

        public MedidaAdif()
        {

        }

        public MedidaAdif(DateTime fd, DateTime fh, Dictionary<string, EndesaEntity.medida.AdifInventario> dic_inventario)
        {
            dic_informe = new Dictionary<string, EndesaEntity.medida.AdifInformeMedida>();


            ConsultaMedidaProporcionadaPorAdif("A", fd, fh, dic_inventario);
            ConsultaMedidaProporcionadaPorAdif("R", fd, fh, dic_inventario);

            // Como debemos mostrar todo el inventario y su situación
            // primero creamos tantos registros de informe como registros de inventario
            this.VuelcaInventario(dic_inventario);

            // Anexamos la medida de ADIF a todos los puntos de inventario
            // this.VuelcaMedidaAdif(dic_inventario);

            // Anexamos la medida del SCE 
            m = new medida.MedidaFunciones(dic_inventario.Select(z => z.Value.cups13).ToList(), fd, fh);
            this.VuelcaMedidaSCE(m);
        }

        public void ExportaMedidaADIF()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            DateTime fecha_hora = new DateTime();
            
            StreamWriter sw;
            string nombreArchivo = @"c:\Temp\CC_ADIF.CSV";
            string linea = "";
            string cups22 = "";
            bool firstOnly = true;

            List<string> lista_cups = new List<string>();

            medida.PuntosMedidaPrincipalesVigentes pm;


            Dictionary<string, List<EndesaEntity.medida.P1>> dic =
               new Dictionary<string, List<EndesaEntity.medida.P1>>();

            strSql = "select CUPS20 from adif_borrar_puntos_pendientes"
                + " group by CUPS20";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                lista_cups.Add(r["CUPS20"].ToString());
            }
            db.CloseConnection();

            pm = new medida.PuntosMedidaPrincipalesVigentes(lista_cups);


            StreamWriter swa = new StreamWriter(@"c:\Temp\Multipuntos.txt", false);
            foreach (KeyValuePair<string, EndesaEntity.medida.PuntoSuministro>  p in pm.dic)
            {
                
                if(p.Value.cups22.Count > 1)
                {
                    linea = p.Key ;
                    for (int i = 0; i < p.Value.cups22.Count; i++)
                        linea = linea + ";" + p.Value.cups22[i];

                    swa.WriteLine(linea);
                }
                
            }
            swa.Close();


            #region Activa
            strSql = "SELECT p.CUPS20, m.Fecha,"
               + " Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9, Value10,"
               + " Value11, Value12, Value13, Value14, Value15, Value16, Value17, Value18, Value19, Value20,"
               + " Value21, Value22, Value23, Value24, Value25"
               + " FROM med.adif_borrar_puntos_pendientes p inner join"
               + " med.adif_medida_horaria_adif m on"
               + " m.CCOUNIPS = p.CUPS13 AND"
               + " m.TipoEnergia = 'A' and"
               + " (year(m.Fecha) = substr(p.aaaammPdte, 1, 4) AND"
               + " MONTH(m.Fecha) = substr(p.aaaammPdte, 5, 2))"
               + " inner join (select ID, CCOUNIPS, max(FechaCarga) FechaCarga, Fecha from adif_medida_horaria_adif b where"
               + " b.TipoEnergia = 'A' group by b.ID, b.Fecha) b on"
               + " b.ID = m.ID and"
               + " b.FechaCarga = m.FechaCarga and"
               + " b.Fecha = m.Fecha"
               + " order by p.CUPS20 , m.Fecha";

            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                firstOnly = true;
                fecha_hora = Convert.ToDateTime(r["Fecha"]).Date;

                EndesaEntity.medida.PuntoSuministro o;
                if (pm.dic.TryGetValue(r["CUPS20"].ToString(),out o))
                {
                    cups22 = o.cups22[0];
                }


                List<EndesaEntity.medida.P1> oo;
                if(!dic.TryGetValue(cups22, out oo))
                {
                    oo = new List<EndesaEntity.medida.P1>();
                    for (int i = 1; i <= 25; i++)
                    {
                        
                        if(r["Value" + i] != System.DBNull.Value)
                        {
                            EndesaEntity.medida.P1 c = new EndesaEntity.medida.P1();
                            c.cups22 = cups22;
                            c.ae = Convert.ToDouble(r["Value" + i]);
                            if (firstOnly)
                            {
                                c.fecha_hora = Convert.ToDateTime(r["Fecha"]).Date;
                                firstOnly = false;
                            }
                            else
                                c.fecha_hora = fecha_hora;

                            oo.Add(c);
                        }
                        
                        fecha_hora = fecha_hora.AddHours(1);

                        
                                                
                    }

                    dic.Add(cups22, oo);

                }
                else
                {
                    for (int i = 1; i <= 25; i++)
                    {
                       
                        if (r["Value" + i] != System.DBNull.Value)
                        {
                            EndesaEntity.medida.P1 c = new EndesaEntity.medida.P1();
                            c.cups22 = cups22;
                            c.ae = Convert.ToDouble(r["Value" + i]);
                            if (firstOnly)
                            {
                                c.fecha_hora = Convert.ToDateTime(r["Fecha"]).Date;
                                firstOnly = false;
                            }
                            else
                                c.fecha_hora = fecha_hora;

                            oo.Add(c);
                        }
                        
                        fecha_hora = fecha_hora.AddHours(1);
                    }
                    
                }

                
            }
            db.CloseConnection();
            #endregion


            #region Reactiva
            strSql = "SELECT p.CUPS20, m.Fecha,"
                + " Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9, Value10,"
                + " Value11, Value12, Value13, Value14, Value15, Value16, Value17, Value18, Value19, Value20,"
                + " Value21, Value22, Value23, Value24, Value25"
                + " FROM med.adif_borrar_puntos_pendientes p inner join"
                + " med.adif_medida_horaria_adif m on"
                + " m.CCOUNIPS = p.CUPS13 AND"
                + " m.TipoEnergia = 'R' and"
                + " (year(m.Fecha) = substr(p.aaaammPdte, 1, 4) AND"
                + " MONTH(m.Fecha) = substr(p.aaaammPdte, 5, 2))"
                + " inner join (select ID, CCOUNIPS, max(FechaCarga) FechaCarga, Fecha from adif_medida_horaria_adif b where"               
                + " b.TipoEnergia = 'R' group by b.ID, b.Fecha) b on"
                + " b.ID = m.ID and"
                + " b.FechaCarga = m.FechaCarga and"
                + " b.Fecha = m.Fecha"
                + " order by p.CUPS20 , m.Fecha";

            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();

            while (r.Read())
            {
                firstOnly = true;
                fecha_hora = Convert.ToDateTime(r["Fecha"]).Date;

                EndesaEntity.medida.PuntoSuministro o;
                if (pm.dic.TryGetValue(r["CUPS20"].ToString(), out o))
                {
                    cups22 = o.cups22[0];
                }


                List<EndesaEntity.medida.P1> oo;
                if (dic.TryGetValue(cups22, out oo))
                {
                    for (int i = 1; i <= 25; i++)
                    {
                        if (r["Value" + i] != System.DBNull.Value)
                        {
                            if (firstOnly)
                            {
                                fecha_hora = Convert.ToDateTime(r["Fecha"]).Date;
                                firstOnly = false;
                            }
                            else
                                fecha_hora = fecha_hora.AddHours(1);

                            for (int j = 0; j < oo.Count(); j++)
                                if (oo[j].fecha_hora == fecha_hora)
                                    oo[j]._as = Convert.ToDouble(r["Value" + i]);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("error");
                }

                    
               

                    
                


            }
            db.CloseConnection();
            #endregion


            sw = new StreamWriter(nombreArchivo, false);
            foreach (KeyValuePair<string,List<EndesaEntity.medida.P1>> p in dic)
            {
                foreach(EndesaEntity.medida.P1 pp in p.Value)
                {
                    linea = pp.fecha_hora.ToString("dd/MM/yyyy") + ";"
                        + Convert.ToInt32(pp.fecha_hora.ToString("HH")) + ":" + pp.fecha_hora.ToString("mm") + ";"
                        + Convert.ToInt32(pp.ae) + ";"
                        + Convert.ToInt32(pp._as) + ";"
                        + "0" + ";"
                        + p.Key;

                    sw.WriteLine(linea);
                }
            }
            sw.Close();

           

        }



        private void ConsultaMedidaProporcionadaPorAdif(string tipoEnergia, DateTime fd, DateTime fh, Dictionary<string, EndesaEntity.medida.AdifInventario> dic_inventario)
        {

            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            List<int> lista_inventario_id = new List<int>();
            int j = 0;

            try
            {
                lista_inventario_id = dic_inventario.Select(z => z.Value.id_cups).ToList();
                for (int i = 0; i < lista_inventario_id.Count; i++)
                {
                    j++;
                    if (firstOnly)
                    {
                        // sb.Append("select c.ID_CUPS, c.CUPS13, c.CUPS20 , sum(a.Total) as total ,a.FechaCarga, a.Fecha");
                        sb.Append("select c.ID_CUPS, c.CUPS13, c.CUPS20 , a.Total as total ,a.FechaCarga, a.Fecha");

                        for (int x = 1; x <= 25; x++)
                            sb.Append(" ,a.Value" + x);

                        sb.Append(" from adif_cups c inner join");
                        sb.Append(" adif_medida_horaria_adif a on");
                        sb.Append(" c.ID_CUPS = a.ID");
                        sb.Append(" inner join (select ID, CCOUNIPS, max(FechaCarga)FechaCarga, Fecha from adif_medida_horaria_adif b where");
                        sb.Append(" (b.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' and b.Fecha <= '").Append(fh.ToString("yyyy-MM-dd")).Append("')");
                        sb.Append(" and b.TipoEnergia = '").Append(tipoEnergia).Append("' group by b.ID, month(Fecha)) b on");
                        sb.Append(" b.ID = a.ID and");
                        sb.Append(" b.FechaCarga = a.FechaCarga");
                        sb.Append(" where");
                        sb.Append(" c.ID_CUPS in (");
                        sb.Append(lista_inventario_id[i]);

                        firstOnly = false;
                    }
                    else
                        sb.Append(" ,").Append(lista_inventario_id[i]);

                    if (j == 250)
                    {
                        sb.Append(")");
                        sb.Append(" and (a.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' AND a.Fecha <= '");
                        sb.Append(fh.ToString("yyyy-MM-dd")).Append("') AND a.TipoEnergia = '").Append(tipoEnergia).Append("'");
                        sb.Append(" group by c.ID_CUPS, month(a.Fecha)");
                        j = 0;
                        firstOnly = true;
                        this.RunQuery(fd, fh, tipoEnergia, sb.ToString(), dic_inventario);
                        sb = null;
                        sb = new StringBuilder();
                    }

                }

                if (j > 0)
                {
                    sb.Append(")");
                    sb.Append(" and (a.Fecha >= '").Append(fd.ToString("yyyy-MM-dd")).Append("' AND a.Fecha <= '");
                    sb.Append(fh.ToString("yyyy-MM-dd")).Append("') AND a.TipoEnergia = '").Append(tipoEnergia).Append("'");
                    sb.Append(" order by c.ID_CUPS, a.Fecha");
                    j = 0;
                    firstOnly = true;
                    this.RunQuery(fd, fh, tipoEnergia, sb.ToString(), dic_inventario);
                    sb = null;
                    sb = new StringBuilder();
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                 "InformeErse.ConsultaMedida",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }
        }

        private void RunQuery(DateTime fd, DateTime fh, string tipoEnergia, string q, Dictionary<string, EndesaEntity.medida.AdifInventario> dic_inventario)
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string key;

            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(q, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                key = r["CUPS20"].ToString() + fd.ToString("yyyyMMdd") + fh.ToString("yyyyMMdd");
                EndesaEntity.medida.AdifInventario i;
                if (tipoEnergia == "A")
                {
                    if (dic_inventario.TryGetValue(key, out i))
                        i.adif_a = Convert.ToInt32(r["total"]);
                }
                else
                {
                    if (dic_inventario.TryGetValue(key, out i))
                        i.adif_r = Convert.ToInt32(r["total"]);
                }

            }
            db.CloseConnection();
        }

        private void VuelcaInventario(Dictionary<string, EndesaEntity.medida.AdifInventario> dic_inventario)
        {

            foreach (KeyValuePair<string, EndesaEntity.medida.AdifInventario> p in dic_inventario)
            {

                EndesaEntity.medida.AdifInformeMedida inf = new EndesaEntity.medida.AdifInformeMedida();
                inf.cups13 = p.Value.cups13;
                inf.cups20 = p.Value.cups20;
                inf.ffactdes = p.Value.ffactdes;
                inf.ffacthas = p.Value.ffacthas;
                inf.tarifa = p.Value.tarifa;
                inf.lote = p.Value.lote;
                inf.adif_a = p.Value.adif_a;
                inf.adif_r = p.Value.adif_r;
                inf.estado_contrato = p.Value.estado_contrato;
                dic_informe.Add(p.Key, inf);

            }
        }

        //private void VuelcaMedidaAdif(Dictionary<string, EndesaEntity.medida.AdifInventario> dic_inventario)
        //{
        //    foreach (KeyValuePair<string, EndesaEntity.medida.AdifInformeMedida> p in dic_informe)
        //    {
        //        EndesaEntity.medida.AdifInventario inv = new EndesaEntity.medida.AdifInventario();
        //        inv = dic_inventario.Where(z => z.Value.cups13 == p.Value.cups13).SingleOrDefault().Value;
        //        p.Value.adif_a = inv.adif_a;
        //        p.Value.adif_r = inv.adif_r;
        //    }

        //}




        private void VuelcaMedidaSCE(medida.MedidaFunciones m)
        {
            string key;
            double calculo;
            double denominador;
            double umbral;
            Param pp = new Param();

            umbral = Convert.ToDouble(pp.GetParam("Umbral_diferencia"));

            foreach (KeyValuePair<string, EndesaEntity.medida.AdifInformeMedida> p in dic_informe)
            {
                EndesaEntity.Medida medida;
                key = p.Value.cups13 + p.Value.ffactdes.ToString("yyyyMMdd") + p.Value.ffacthas.ToString("yyyyMMdd");
                if (m.dicMedida.TryGetValue(key, out medida))
                {
                    p.Value.ffactdes = medida.fromdate;
                    p.Value.ffacthas = medida.todate;
                    p.Value.fuente = medida.fuente;
                    p.Value.sce_a = medida.activa;
                    p.Value.sce_r = medida.reactiva;
                    p.Value.dias = medida.dias;
                    p.Value.estado_curva = medida.estado;
                    p.Value.resumen_sce_a = medida.activa;
                    p.Value.resumen_sce_r = medida.reactiva;
                    p.Value.dif_sce_adif_a = medida.activa - p.Value.adif_a;
                    p.Value.dif_sce_adif_r = medida.reactiva - p.Value.adif_r;

                    if (p.Value.sce_a == 0)
                        denominador = 1;
                    else
                        denominador = p.Value.sce_a;

                    calculo = (Math.Abs(p.Value.sce_a - p.Value.adif_a) / denominador) * 100;

                    if (p.Value.sce_a == p.Value.adif_a)
                    {
                        p.Value.resultado = "SCE=ADIF";
                        p.Value.enviado_facturar = "OK";
                    }
                    else if (umbral > calculo || Math.Abs(p.Value.sce_a - p.Value.adif_a) < 100)
                    {
                        p.Value.resultado = "%ERROR ASUMIBLE";
                        p.Value.enviado_facturar = "OK";
                    }
                    else
                    {
                        p.Value.resultado = "REVISAR";
                        p.Value.enviado_facturar = "REVISAR";
                    }
                }
            }
        }

        private void Obtener_Medida_ADIF_FTP(DateTime f)
        {


            string dirFTP = "";
            string[] lista_carpetas;
            List<string> rutas_validas = new List<string>();

            string[] lista_archivos;

            string mes_letra = "";
            string year = "";
            string subruta = "";

            utilidades.Fechas utilFecha = new utilidades.Fechas();

            try
            {
                mes_letra = utilFecha.ConvierteMes_a_Letra(f);
                year = f.Year.ToString();

                utilidades.Param param = new utilidades.Param("adif_param", MySQLDB.Esquemas.MED);

                utilidades.FTP ftpClient = new utilidades.FTP(param.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                    param.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                    param.GetValue("ftp_pass", DateTime.Now, DateTime.Now));



                dirFTP = param.GetValue("ftp_main_folder", DateTime.Now, DateTime.Now);
                lista_carpetas = ftpClient.DirectoryListSimple(dirFTP);
                for (int i = 0; i < lista_carpetas[i].Count(); i++)
                {
                    if (lista_carpetas[i].Contains("ADIF_GRUPO_"))
                    {
                        subruta = lista_carpetas[i].Replace(dirFTP, "");
                        subruta = subruta.Replace("/", "");
                        subruta = subruta.Replace("ADIF_", "") + "_CURVAS_ADIF";

                        rutas_validas.Add(lista_carpetas[i] + "/"
                            + subruta + "/"
                            + year + "/"
                            + subruta
                            + "_MES_" + mes_letra + "_" + year + "/");
                    }

                }

                for (int i = 0; i < rutas_validas.Count(); i++)
                {
                    lista_archivos = ftpClient.DirectoryListSimple(rutas_validas[i]);
                    for (int j = 0; j < lista_archivos.Count(); j++)
                    {
                        if (lista_archivos[j].Contains("_CC1_"))
                            ftpClient.Download(rutas_validas[i] + lista_archivos[j], @"c:\Temp\adif\" + lista_archivos[j]);
                    }
                }








                //strFichero1 = param.GetValue("strFichero1", DateTime.Now, DateTime.Now) +
                //    ultimoDiaHabil.ToString("yyyyMMdd") + "." +
                //    param.GetValue("extensionFicheros", DateTime.Now, DateTime.Now);

                //strFichero2 = param.GetValue("strFichero3", DateTime.Now, DateTime.Now) +
                //    ultimoDiaHabil.ToString("yyyyMMdd") + "." +
                //    param.GetValue("extensionFicheros", DateTime.Now, DateTime.Now);

                //ftpClient.Download(dirFTP + strFichero1, param.GetValue("DirectorioLocal", DateTime.Now, DateTime.Now) + strFichero1);
                //ftpClient.Download(dirFTP + strFichero1, param.GetValue("DirectorioLocal", DateTime.Now, DateTime.Now) + strFichero2);


                //FileInfo file = new FileInfo(param.GetValue("DirectorioLocal", DateTime.Now, DateTime.Now) + strFichero1);

                //if (file.Exists)
                //{
                //    this.ImportarPdteWeb(file.FullName,
                //        Convert.ToChar(param.GetValue("SeparadorFicheros", DateTime.Now, DateTime.Now)),
                //        Convert.ToInt32(param.GetValue("LineaLectura", DateTime.Now, DateTime.Now)),
                //        ultimoDiaHabil);

                //}



            }
            catch (Exception e)
            {
                // ficheroLog.Add("ObtenerPDTESWeb --> " + e.Message);
            }
        }
    }
}
