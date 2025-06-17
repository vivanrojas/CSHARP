using EndesaBusiness.servidores;
using EndesaEntity;
using Microsoft.SharePoint.Client;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot.Types;
using static Microsoft.Exchange.WebServices.Data.SearchFilter;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;

namespace EndesaBusiness.adif
{
    public class InventarioFunciones : EndesaEntity.medida.AdifInventario
    {
        // Key = cups20+fd+fh --> fechas (yyyymmdd)
        public Dictionary<string, EndesaEntity.medida.AdifInventario> dic_inventario { get; set; }
        EndesaBusiness.adif.CierresEnergia funcion_cierres_energia;


        public InventarioFunciones(DateTime fd, DateTime fh, string cups20, List<string> lista_lotes)
        {
            int numMeses = 0;
            bool firstOnly = true;
            int anio = 0;
            int mes = 0;
            int dias_del_mes = 0;

            dic_inventario = new Dictionary<string, EndesaEntity.medida.AdifInventario>();

            funcion_cierres_energia = new EndesaBusiness.adif.CierresEnergia();

            if (fd.Month != fh.Month)
            {
                numMeses = (((fh.Year - fd.Year) * 12) + fh.Month - fd.Month) + 1;
                for (int i = 0; i < numMeses; i++)
                {
                    if (firstOnly)
                    {
                        firstOnly = false;
                    }
                    else
                        fd = fd.AddMonths(1);

                    anio = fd.Year;
                    mes = fd.Month;
                    dias_del_mes = DateTime.DaysInMonth(anio, mes);
                    fh = new DateTime(anio, mes, dias_del_mes);
                    fd = new DateTime(fh.Year, fh.Month, 1);
                    CargaInventario(fd, fh, cups20, lista_lotes);
                }
            }
            else
                CargaInventario(fd, fh, cups20, lista_lotes);
        }

        
        public InventarioFunciones()
        {
            dic_inventario = new Dictionary<string, EndesaEntity.medida.AdifInventario>();
        }

        public void GetInventario(string cups20)
        {

            EndesaEntity.medida.AdifInventario inv = dic_inventario.Where(z => z.Value.cups20 == cups20).SingleOrDefault().Value;
            this.Set(inv);
        }


        private void Set(EndesaEntity.medida.AdifInventario inv)
        {
            
            this.cups20 = inv.cups20;
            this.lote = inv.lote;
            this.vigencia_desde = inv.vigencia_desde;
            this.vigencia_hasta = inv.vigencia_hasta;
            this.ffactdes = inv.ffactdes;
            this.ffacthas = inv.ffacthas;
            this.tarifa = inv.tarifa;
            this.tension = inv.tension;
            this.zona = inv.zona;
            this.codigo = inv.codigo;
            this.nombre_punto_suministro = inv.nombre_punto_suministro;
            this.distribuidora = inv.distribuidora;
            this.comentarios = this.comentarios;
            this.medida_en_baja = inv.medida_en_baja;
            this.devolucion_de_energia = inv.devolucion_de_energia;
            this.cierres_energia = inv.cierres_energia;
            this.provincia = inv.provincia;
            this.comunidad_autonoma = inv.comunidad_autonoma;
            this.sitema_traccion = inv.sitema_traccion;
            this.multipunto_num_principales = inv.multipunto_num_principales;
            this.grupo = inv.grupo;
            this.perdidas = inv.perdidas;
            this.valor_kvas = inv.valor_kvas;

            this.p = inv.p;
            this.comentarios = inv.comentarios;

        }

        //public void GetInventario(string cups20)
        //{
        //    EndesaEntity.medida.InventarioAdif inv;
        //    if (dic_inventario.TryGetValue(cups20, out inv))
        //        this.Set(inv);
        //    else
        //        this.id_cups = 0;

        //}


        private void CargaInventario(DateTime fd, DateTime fh, string cups20, List<string> lista_lotes)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int num = 0;
            bool firstOnly = true;

            try
            {

                strSql = "select l.*, if(e.Descripcion is null, 'BAJA', e.Descripcion) as Estado"
                   + " from med.adif_lotes l"
                   + " left outer join cont.PS_AT ps on"
                   + " ps.CUPS20 = l.CUPS20"
                   + " left outer join cont.cont_estadoscontrato e on"
                   + " e.Cod_Estado = ps.estadoCont"
                   + " where"
                   + " l.FECHA_DESDE <= '" + fh.ToString("yyyy-MM-dd") + "' AND "
                   + " l.FECHA_HASTA >= '" + fd.ToString("yyyy-MM-dd") + "'";                   


                if (lista_lotes.Count() > 0)
                {
                    for (int i = 0; i < lista_lotes.Count(); i++)
                    {
                        if (firstOnly)
                        {
                            strSql += " AND l.LOTE in (" + lista_lotes[i];
                            firstOnly = false;
                        }
                        else
                        {
                            strSql += " ," + lista_lotes[i];
                        }
                    }

                    strSql += ")";
                }
                                
                if (cups20 != null)
                {
                    strSql += " and l.CUPS20 = '" + cups20 + "'";
                }


                strSql += " order by l.lote,l.CUPS20";

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.AdifInventario inv = new EndesaEntity.medida.AdifInventario();
                    num++;
                    inv.num = num;
                    
                    inv.vigencia_desde = Convert.ToDateTime(r["FECHA_DESDE"]);
                    inv.vigencia_hasta = Convert.ToDateTime(r["FECHA_HASTA"]);
                    inv.ffactdes = fd;
                    inv.ffacthas = fh;                    
                    inv.cups20 = r["CUPS20"].ToString();
                    inv.lote = Convert.ToInt32(r["LOTE"]);
                    inv.tarifa = r["TARIFA"].ToString();
                    inv.estado_contrato = r["Estado"].ToString();
                    inv.tension = r["TENSION"] != System.DBNull.Value ? Convert.ToInt32(r["TENSION"]) : 0;
                    inv.estado_curva = "SIN CURVA REGISTRADA";
                    inv.zona = r["ZONA"] != System.DBNull.Value ? r["ZONA"].ToString() : null;
                    inv.codigo = r["CODIGO"] != System.DBNull.Value ? r["CODIGO"].ToString() : null;
                    inv.distribuidora = r["DISTRIBUIDORA"] != System.DBNull.Value ? r["DISTRIBUIDORA"].ToString() : null;
                    inv.nombre_punto_suministro = r["NOMBRE_PUNTO_SUMINISTRO"] != System.DBNull.Value ? r["NOMBRE_PUNTO_SUMINISTRO"].ToString() : null;
                    inv.comentarios = r["COMENTARIOS"] != System.DBNull.Value ? r["COMENTARIOS"].ToString() : null;
                    inv.cierres_energia = funcion_cierres_energia.ExisteCierre(inv.cups20, fd, fh);
                    inv.medida_en_baja = r["MEDIDA_EN_BAJA"].ToString() == "S";
                    inv.devolucion_de_energia = r["DEVOLUCION_ENERGIA"].ToString() == "S";

                    if (r["PROVINCIA"] != System.DBNull.Value)
                        inv.provincia = r["PROVINCIA"].ToString();
                    if (r["COMUNIDAD_AUTONOMA"] != System.DBNull.Value)
                        inv.comunidad_autonoma = r["COMUNIDAD_AUTONOMA"].ToString();
                    if (r["SISTEMA_TRACCION"] != System.DBNull.Value)
                        inv.sitema_traccion = r["SISTEMA_TRACCION"].ToString();
                    if (r["GRUPO"] != System.DBNull.Value)
                        inv.grupo = r["GRUPO"].ToString();
                    if (r["VALOR_KVA"] != System.DBNull.Value)
                        inv.valor_kvas = Convert.ToDouble(r["VALOR_KVA"]);
                    if (r["PERDIDAS"] != System.DBNull.Value)
                    {
                        inv.perdidas = Convert.ToDouble(r["PERDIDAS"]);
                        inv.tiene_perdidas = inv.perdidas > 0;
                            
                    }
                        
                    if (r["MULTIPUNTO"] != System.DBNull.Value)
                        inv.multipunto_num_principales = Convert.ToInt32(r["MULTIPUNTO"]);


                    for (int j = 0; j < 6; j++)
                    {
                        if (r["P"+ (j + 1)] != System.DBNull.Value)
                            inv.p[j] = Convert.ToDouble(r["P" + (j + 1)]);
                    }


                    EndesaEntity.medida.AdifInventario i;

                    if (!dic_inventario.TryGetValue(inv.cups20 + inv.ffactdes.ToString("yyyyMMdd") + inv.ffacthas.ToString("yyyyMMdd"), out i))                        
                        dic_inventario.Add(inv.cups20 + inv.ffactdes.ToString("yyyyMMdd") + inv.ffacthas.ToString("yyyyMMdd"), inv);
                    
                        


                }


            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message,
              "InventarioFunciones.CargaInventario",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

        public void CompruebaInventario()
        {
            bool hayCambios = false;
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.medida.AdifInventario> inv = new List<EndesaEntity.medida.AdifInventario>();

            try
            {
                // Comprobamos el inventario a nivel de facturas
                //strSql = "select l.CUPS20,  l.LOTE lote_anterior, f.LOTE lote_nuevo "
                //    + " from med.adif_lotes l"                    
                //    + " left outer join fact.adif_facturas f on"
                //    + " f.CUPSREE = l.CUPS20"
                //    + " where l.LOTE = 0 and f.LOTE is not null"
                //    + " group by l.CUPS20";

                strSql = "SELECT l.CUPS20, l.LOTE lote_anterior, f.LOTE lote_nuevo"
                    + " from med.adif_lotes l"
                    + " LEFT OUTER JOIN fact.adif_ficheros_facturas f ON"
                    + " f.cups20 = l.CUPS20 WHERE l.LOTE = 0 AND f.lote IS NOT NULL GROUP BY l.CUPS20";

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    hayCambios = true;
                    this.UpdateLote(Convert.ToInt32(r["lote_nuevo"]), r["CUPS20"].ToString());
                }

                // Comprobamos el inventario a nivel de tablas PS_AT

                if (hayCambios)
                {
                    MessageBox.Show("Se han detectado cambios en el inventario.",
                    "Comprobación de inventario",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "InventarioFunciones.CompruebaInventario",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }


        private void UpdateLote(int lote, string cups20)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            strSql = "update adif_lotes set LOTE = " + lote
                        + " where CUPS20 = '" + cups20 + "'"
                        + " and LOTE = 0";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha asignado el cups " + cups20 + " con lote 0 "
                + " por el lote " + lote + ".",
            "Comprobación de inventario",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        }

        public void Save()
        {
            try
            {

                EndesaEntity.medida.AdifInventario inv = new EndesaEntity.medida.AdifInventario();
                inv = dic_inventario.Where(z => z.Value.cups20 == this.cups20).SingleOrDefault().Value;

                if (inv != null)
                    Update();
                else
                    New();

               

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "InventarioFunciones - Save",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
                   

        }

        private void New()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "insert into adif_lotes (CUPS20, LOTE, FECHA_DESDE, FECHA_HASTA,"
                    + " ZONA, CODIGO, NOMBRE_PUNTO_SUMINISTRO, DISTRIBUIDORA, COMENTARIOS, TARIFA,"
                    + " TENSION, P1, P2, P3, P4, P5, P6, DEVOLUCION_ENERGIA, MEDIDA_EN_BAJA,"
                    + " PROVINCIA, COMUNIDAD_AUTONOMA, SISTEMA_TRACCION, GRUPO, VALOR_KVA,"
                    + " PERDIDAS, MULTIPUNTO, "
                    + " created_by, created_date) values ";

                if (this.cups20 != null)
                    strSql += "('" + this.cups20 + "'";
                else
                    strSql += "(null";

                if (this.lote >= 0)
                    strSql += ",'" + this.lote + "'";
                else
                    strSql += ",null";

                if (this.ffactdes != DateTime.MinValue)
                    strSql += ",'" + this.ffactdes.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.ffacthas != DateTime.MinValue)
                    strSql += ",'" + this.ffacthas.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.zona != null)
                    strSql += ",'" + this.zona + "'";
                else
                    strSql += ",null";

                if (this.codigo != null)
                    strSql += ",'" + this.codigo + "'";
                else
                    strSql += ",null";

                if (this.nombre_punto_suministro != null)
                    strSql += ",'" + this.nombre_punto_suministro + "'";
                else
                    strSql += ",null";

                if (this.distribuidora != null)
                    strSql += ",'" + this.distribuidora + "'";
                else
                    strSql += ",null";

                if (this.comentarios != null)
                    strSql += ",'" + this.comentarios + "'";
                else
                    strSql += ",null";

                if (this.tarifa != null)
                    strSql += ",'" + this.tarifa + "'";
                else
                    strSql += ",null";

                if (this.tension > 0)
                    strSql += "," + this.tension;
                else
                    strSql += ",null";

                for(int i = 0; i < 6; i++)
                    if (this.p[i] > 0)
                        strSql += "," + this.p[i];
                    else
                        strSql += ",null";

                if(this.devolucion_de_energia)
                    strSql += ",'S'";
                else
                    strSql += ",'N'";

                if (this.medida_en_baja)
                    strSql += ",'S'";
                else
                    strSql += ",'N'";

                if (this.provincia != null)
                    strSql += ",'" + this.provincia + "'";

                if (this.comunidad_autonoma != null)
                    strSql += ",'" + this.comunidad_autonoma + "'";

                if (this.sitema_traccion != null)
                    strSql += ",'" + this.sitema_traccion + "'";

                if (this.grupo != null)
                    strSql += ",'" + this.grupo + "'";

                if (this.valor_kvas > 0)
                    strSql = "," + this.valor_kvas.ToString().Replace(",", ".");

                if (this.perdidas > 0)
                    strSql = "," + this.perdidas.ToString().Replace(",", ".");

                if (this.multipunto_num_principales > 0)
                    strSql = "," + this.multipunto_num_principales;

                if (this.created_by != null)
                    strSql += ",'" + this.created_by + "'";
                else
                    strSql += ",'" + System.Environment.UserName + "'";

                strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                MessageBox.Show("Se ha añadido correctamente el CUPS: " + this.cups20,
                 "Inventario ADIF",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "CierresEnergia - New",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }

        private void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update adif_lotes set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                + " ,LOTE = " + this.lote;
                

            if (this.ffactdes != null)
                strSql += " ,FECHA_DESDE = '" + this.ffactdes.ToString("yyyy-MM-dd") + "'";

            if (this.ffacthas != null)
                strSql += " ,FECHA_HASTA = '" + this.ffacthas.ToString("yyyy-MM-dd") + "'";

            if (this.zona != null)
                strSql += " ,ZONA = '" + this.zona + "'";

            if (this.codigo != null)
                strSql += " ,CODIGO = '" + this.codigo + "'";

            if (this.nombre_punto_suministro != null)
                strSql += " ,NOMBRE_PUNTO_SUMINISTRO = '" + this.nombre_punto_suministro + "'";

            if (this.distribuidora != null)
                strSql += " ,DISTRIBUIDORA = '" + this.distribuidora + "'";

            if (this.comentarios != null)
                strSql += " ,COMENTARIOS = '" + this.comentarios + "'";

            if (this.tarifa != null)
                strSql += " ,TARIFA = '" + this.tarifa + "'";

            if (this.tension != 0)
                strSql += " ,TENSION = " + this.tension ;

            if (this.devolucion_de_energia)
                strSql += " ,DEVOLUCION_ENERGIA = 'S'";
            else
                strSql += " ,DEVOLUCION_ENERGIA = 'N'";

            if (this.medida_en_baja)
                strSql += " ,MEDIDA_EN_BAJA = 'S'";
            else
                strSql += " ,MEDIDA_EN_BAJA = 'N'";

            for(int i = 0; i < 6;  i++)
                if (this.p[i] != 0)
                    strSql += " ,P" + (i + 1) +" = " + this.p[i];

            if (this.provincia != null)
                strSql += " ,PROVINCIA = '" + this.provincia + "'";

            if (this.comunidad_autonoma != null)
                strSql += " ,COMUNIDAD_AUTONOMA = '" + this.comunidad_autonoma + "'";

            if (this.sitema_traccion != null)
                strSql += " ,SISTEMA_TRACCION = '" + this.sitema_traccion + "'";

            if (this.grupo != null)
                strSql += " ,GRUPO = '" + this.grupo + "'";

            
            strSql += " ,VALOR_KVA = " + this.valor_kvas.ToString().Replace(",",".");            
            strSql += " ,PERDIDAS = " + this.perdidas.ToString().Replace(",", ".");            
            strSql += " ,MULTIPUNTO = " + this.multipunto_num_principales;

            strSql += " where CUPS20 = '" + this.cups20 + "' and"
                + " FECHA_DESDE = '" + this.ffactdes.ToString("yyyy-MM-dd") + "'";

            

            db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el inventario para el CUPS: " + this.cups20,
            "Inventario ADIF",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        }

        private void Del()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                DialogResult result = MessageBox.Show("Warning:" + System.Environment.NewLine
                    + "¿Desea borrar el cups " + this.cups20 + "?"
                    , "Borrado del cups " + this.cups20,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    strSql = "delete from adif_lotes where cups20 = '" + this.cups20 + "'"
                        + " and fecha_desde = '" + this.ffactdes.ToString("yyyy-MM-dd") + "'";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("registro borrado.",
                      "Borrado inventario",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "InventarioFunciones - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }


        public int GetID_from_Cups13(string cups13)
        {
            EndesaEntity.medida.AdifInventario inv = new EndesaEntity.medida.AdifInventario();
            inv = dic_inventario.Where(z => z.Value.cups13 == cups13).SingleOrDefault().Value;
            if (inv != null)
            {
                this.Set(inv);
                return inv.id_cups;
            }
            return 0;

        }

        public void GetRegistroCUPS20(string cups20)
        {
            EndesaEntity.medida.AdifInventario inv = new EndesaEntity.medida.AdifInventario();
            inv = dic_inventario.Where(z => z.Value.cups20 == cups20).SingleOrDefault().Value;
            if (inv != null)
            {
                this.Set(inv);                
            }


        }

        public void GetRegistro(string cups20, DateTime fd, DateTime fh)
        {
            //string key = cups20 + fd.ToString("yyyyMMdd")
            //    + fh.ToString("yyyyMMdd");

            //foreach(KeyValuePair<string, EndesaEntity.medida.AdifInventario>)

            //if (dic_inventario.TryGetValue(key, out o))
            //{
            //    this.cups20 = o.cups20;
            //    this.lote = o.lote;
            //    this.ffactdes = o.ffactdes;
            //    this.ffacthas = o.ffacthas;
            //    this.zona = o.zona;
            //    this.codigo = o.codigo;
            //    this.nombre_punto_suministro = o.nombre_punto_suministro;
            //    this.distribuidora = o.distribuidora;
            //    this.comentarios = o.comentarios;
            //    this.tarifa = o.tarifa;
            //    this.tension = o.tension;
            //    this.p = o.p;
            //    this.devolucion_de_energia = o.devolucion_de_energia;
            //    this.medida_en_baja = o.medida_en_baja;                

            //}
        }

    }
}
