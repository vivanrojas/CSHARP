using EndesaBusiness.calendarios;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EndesaBusiness.eer
{
    public class FacturasEER : EndesaEntity.eer.Factura
    {

        public Dictionary<int, EndesaEntity.eer.Factura> dic;
        public string estado { get; set; }
        public bool existe { get; set; }
        utilidades.Param param;

        public FacturasEER()
        {
            param = new utilidades.Param("eer_param", MySQLDB.Esquemas.CON);
        }


        public FacturasEER(DateTime fd, DateTime fh)
        {
            param = new utilidades.Param("eer_param", MySQLDB.Esquemas.CON);
            dic = Carga(fd, fh);
        }

        private Dictionary<int, EndesaEntity.eer.Factura> Carga(DateTime fd, DateTime fh)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            bool firstOnly = true;
            


            Dictionary<int, EndesaEntity.eer.Factura> d = new Dictionary<int, EndesaEntity.eer.Factura>();

            try
            {
                strSql = "select id_factura, nif, razon_social, cups20," 
                    + " fecha_consumo_desde, fecha_consumo_hasta, codigo_factura, fecha_factura," 
                    + " consumo_activa, consumo_reactiva, tarifa, direccion_fiscal," 
                    + " direccion_suministro, direccion_envio, base_ise, impuesto_electricidad," 
                    + " base_ise_reducido, impuesto_electricidad_reducido, base_imponible, iva," 
                    + " total_factura, termino_energia, descuento_energia, facturacion_potencia," 
                    + " recargo_excesos, reactiva, suplemento_territorial, servicio_gestion_preferente," 
                    + " alquiler, ruta_factura, f_ult_mod"                
                    + " from eer_facturas where"
                    + " (fecha_consumo_desde >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_consumo_hasta <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " order by cups20, fecha_factura";
                    
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.eer.Factura c = new EndesaEntity.eer.Factura();
                    c.id_factura = Convert.ToInt32(r["id_factura"]);
                    c.cupsree = r["cups20"].ToString();
                    c.nif = r["nif"].ToString();
                    c.nombre_cliente = r["razon_social"].ToString();

                    c.fecha_consumo_desde = Convert.ToDateTime(r["fecha_consumo_desde"]);
                    c.fecha_consumo_hasta = Convert.ToDateTime(r["fecha_consumo_hasta"]);

                    if (r["codigo_factura"] != System.DBNull.Value)
                        c.codigo_factura = r["codigo_factura"].ToString();

                    if (r["fecha_factura"] != System.DBNull.Value)
                        c.fecha_factura = Convert.ToDateTime(r["fecha_factura"]);

                    if (r["consumo_activa"] != System.DBNull.Value)
                        c.consumo_activa = Convert.ToDouble(r["consumo_activa"]);

                    if (r["consumo_reactiva"] != System.DBNull.Value)
                        c.consumo_reactiva = Convert.ToDouble(r["consumo_reactiva"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["direccion_fiscal"] != System.DBNull.Value)
                        c.direccion_facturacion = r["direccion_fiscal"].ToString();

                    if (r["direccion_suministro"] != System.DBNull.Value)
                        c.direccion_suministro = r["direccion_suministro"].ToString();

                    if (r["direccion_envio"] != System.DBNull.Value)
                        c.direccion_envio = r["direccion_envio"].ToString();

                    if (r["base_ise"] != System.DBNull.Value)
                        c.base_ise = Convert.ToDouble(r["base_ise"]);

                    if (r["impuesto_electricidad"] != System.DBNull.Value)
                        c.impuesto_electricidad = Convert.ToDouble(r["impuesto_electricidad"]);

                    if (r["base_ise_reducido"] != System.DBNull.Value)
                        c.base_ise_reducido = Convert.ToDouble(r["base_ise_reducido"]);

                    if (r["impuesto_electricidad_reducido"] != System.DBNull.Value)
                        c.impuesto_electricidad_reducido = Convert.ToDouble(r["impuesto_electricidad_reducido"]);

                    if (r["base_imponible"] != System.DBNull.Value)
                        c.base_imponible = Convert.ToDouble(r["base_imponible"]);

                    if (r["iva"] != System.DBNull.Value)
                        c.iva = Convert.ToDouble(r["iva"]);

                    if (r["recargo_excesos"] != System.DBNull.Value)
                        c.recargo_excesos = Convert.ToDouble(r["recargo_excesos"]);

                    if (r["total_factura"] != System.DBNull.Value)
                        c.total_factura = Convert.ToDouble(r["total_factura"]);

                    if (r["termino_energia"] != System.DBNull.Value)
                        c.termino_energia = Convert.ToDouble(r["termino_energia"]);

                    if (r["descuento_energia"] != System.DBNull.Value)
                        c.descuento_energia = Convert.ToDouble(r["descuento_energia"]);

                    if (r["facturacion_potencia"] != System.DBNull.Value)
                        c.facturacion_potencia = Convert.ToDouble(r["facturacion_potencia"]);

                    if (r["recargo_excesos"] != System.DBNull.Value)
                        c.recargo_excesos = Convert.ToDouble(r["recargo_excesos"]);

                    if (r["suplemento_territorial"] != System.DBNull.Value)
                        c.suplemento_territorial = Convert.ToDouble(r["suplemento_territorial"]);

                    if (r["alquiler"] != System.DBNull.Value)
                        c.alquiler = Convert.ToDouble(r["alquiler"]);


                    EndesaEntity.eer.Factura o;
                    if (!d.TryGetValue(c.id_factura, out o))                    
                        d.Add(c.id_factura, c);
                    

                }
                db.CloseConnection();

                
                if(d.Count > 0)
                {
                    strSql = "SELECT id_factura, linea, producto, concepto, calculo, cantidad, unidad_cantidad, "
                   + " precio, unidad_precio, total"
                   + " FROM cont.eer_facturasdetalle where id_factura in (";
                    foreach (KeyValuePair<int, EndesaEntity.eer.Factura> p in d)
                    {
                        if (firstOnly)
                        {
                            strSql += p.Key;
                            firstOnly = false;
                        }
                        else
                            strSql += ", " + p.Key;
                    }
                    strSql += ")";
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        EndesaEntity.eer.FacturaDetalle fad = new EndesaEntity.eer.FacturaDetalle();
                        fad.id_factura = Convert.ToInt32(r["id_factura"]);
                        fad.linea_factura = Convert.ToInt32(r["linea"]);
                        fad.producto = r["producto"].ToString();

                        if (r["concepto"] != System.DBNull.Value)
                            fad.concepto = r["concepto"].ToString();

                        if (r["total"] != System.DBNull.Value)
                            fad.total = Convert.ToDouble(r["total"]);

                        EndesaEntity.eer.Factura f;
                        if (d.TryGetValue(fad.id_factura, out f))
                            f.lista_factura_detalle.Add(fad);

                    }
                    db.CloseConnection();
                }
                              

                return d;
            }catch(Exception e)
            {
                return null;
            }

        }

        public void GetInvoice(string cups20, DateTime fd, DateTime fh)
        {
            EndesaEntity.eer.Factura o;
            this.existe = false;
            this.estado = "Pdte. Calcular";

            foreach (KeyValuePair<int, EndesaEntity.eer.Factura> d in dic)
            {
                        

                if (d.Value.cupsree == cups20 &&
                    d.Value.fecha_consumo_desde == fd && 
                    d.Value.fecha_consumo_hasta == fh)
                {
                    this.codigo_factura = d.Value.codigo_factura;
                    this.fecha_factura = d.Value.fecha_factura;
                    this.existe = true;

                    if (d.Value.codigo_factura == null)
                        this.estado = "Calculada";
                    else
                        this.estado = "Facturada";

                    break;
                }
            }
    
        }

        private void GuardaNumFacturaTemporal(int numFacturaTemporal)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            
            strSql = "update eer_param set value = '" + numFacturaTemporal + "'"
                + " where code = 'ultima_factura'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void GuardaFactura(EndesaEntity.eer.Factura factura)
        {

            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            double total = 0;
            string strSql = "";
            int dias_anio = 0;

            utilidades.Fechas utilFechas = new utilidades.Fechas();


            try
            {
                // Si existe la factura calculada la BORRAMOS
                // Si existe la factura registrada la añadimos

                dias_anio = utilFechas.EsAnioBisiesto(factura.fecha_consumo_desde) ? 366 : 365;

                strSql = "select id_factura from eer_facturas where"
                    + " cups20 = '" + factura.cupsree + "' and"
                    + " (fecha_consumo_desde = '" + factura.fecha_consumo_desde.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_consumo_hasta = '" + factura.fecha_consumo_hasta.ToString("yyyy-MM-dd") + "') and"
                    + " codigo_factura is null";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    BorraFactura(Convert.ToInt32(r["id_factura"]));
                }
                db.CloseConnection();



                sb.Append("replace into eer_facturas");
                sb.Append(" (id_factura, nif, razon_social, codigo_factura, fecha_factura, fecha_consumo_desde, fecha_consumo_hasta,");
                sb.Append(" consumo_activa, consumo_reactiva, cups20, tarifa,");
                sb.Append(" direccion_fiscal, direccion_suministro, direccion_envio, base_ise, impuesto_electricidad,");
                sb.Append(" base_ise_reducido, impuesto_electricidad_reducido, base_imponible, iva, total_factura,");
                sb.Append(" termino_energia, descuento_energia, facturacion_potencia, recargo_excesos,");
                sb.Append(" reactiva, suplemento_territorial, servicio_gestion_preferente, alquiler");
                sb.Append(") values (");
                sb.Append(factura.id_factura).Append(",");
                sb.Append("'").Append(factura.nif).Append("',");
                sb.Append("'").Append(factura.nombre_cliente).Append("',");                
                sb.Append("null").Append(","); // codigo_factura
                sb.Append("null").Append(","); // fecha_factura
                sb.Append("'").Append(factura.fecha_consumo_desde.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(factura.fecha_consumo_hasta.ToString("yyyy-MM-dd")).Append("',");
                sb.Append(factura.consumo_activa.ToString().Replace(",",".")).Append(",");
                sb.Append(factura.consumo_reactiva.ToString().Replace(",", ".")).Append(",");
                sb.Append("'").Append(factura.cupsree).Append("',");
                sb.Append("'").Append(factura.tarifa).Append("',");

                if (factura.direccion_facturacion != null)
                    sb.Append("'").Append(factura.direccion_facturacion).Append("',");
                else
                    sb.Append(" null,");

                if (factura.direccion_suministro != null)
                    sb.Append("'").Append(factura.direccion_suministro).Append("',");
                else
                    sb.Append(" null,");

                if (factura.direccion_envio != null)
                    sb.Append("'").Append(factura.direccion_envio).Append("',");
                else
                    sb.Append(" null,");




                #region Impuesto Electricidad
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> linea_ise = factura.lista_factura_detalle.FindAll(z => z.producto == "IE");

                if(linea_ise.Count > 0)
                {
                    for (int i = 0; i < linea_ise.Count; i++)
                        total = total + linea_ise[i].total;

                    sb.Append(linea_ise[0].precio.ToString().Replace(",", ".")).Append(",");
                    sb.Append(total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append(" null, null,");


                #endregion

                #region Impuesto Electricidad Reducido

                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> linea_ise_reducido 
                    = factura.lista_factura_detalle.FindAll(z => z.producto == "ISE");

                if (linea_ise_reducido.Count > 0)
                {
                    for (int i = 0; i < linea_ise_reducido.Count; i++)
                        total = total + linea_ise_reducido[i].total;

                    sb.Append(linea_ise_reducido[0].precio.ToString().Replace(",", ".")).Append(",");
                    sb.Append(total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append(" null, null,");

                #endregion


                sb.Append(factura.base_imponible.ToString().Replace(",", ".")).Append(",");
                sb.Append(factura.iva.ToString().Replace(",", ".")).Append(",");
                sb.Append(factura.total_factura.ToString().Replace(",", ".")).Append(",");


                #region Termino Energia Variable              
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> lineas_energia_variable = factura.lista_factura_detalle.FindAll(z => z.producto == "L01");
                if (lineas_energia_variable.Count > 0)
                {
                    for (int i = 0; i < lineas_energia_variable.Count; i++)
                        total = total + lineas_energia_variable[i].total;

                    sb.Append(total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append("null,");

                #endregion

                #region Dto Cliente
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> dto_cliente = factura.lista_factura_detalle.FindAll(z => z.producto == "DTO_TE");
                if(dto_cliente.Count > 0)
                {
                    for (int i = 0; i < dto_cliente.Count; i++)
                        sb.Append(dto_cliente[i].total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append("null,");

                #endregion

                #region Facturacion Potencia Periodos
                total = 0;
                string producto_potencia_periodos = factura.fecha_consumo_desde >= new DateTime(2021, 06, 01) ? "L85" : "L34";
                List<EndesaEntity.eer.FacturaDetalle> lineas_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == producto_potencia_periodos);
                if (lineas_potencia.Count > 0)
                {


                    if (factura.fecha_consumo_desde >= new DateTime(2021, 06, 01))
                    {

                        for (int i = 0; i < lineas_potencia.Count; i++)
                            total = total + lineas_potencia[i].total;

                        sb.Append((total / dias_anio).ToString().Replace(",", ".")).Append(",");
                    }
                    else
                    {
                        for (int i = 0; i < lineas_potencia.Count; i++)
                            total = total + lineas_potencia[i].total;

                        sb.Append((total / 12).ToString().Replace(",", ".")).Append(",");
                    }

                }
                else
                    sb.Append("null,");

                #endregion

                #region Excesos Potencia
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> linea_excesos_potencia = factura.lista_factura_detalle.FindAll(z => z.producto == "REPA");
                if (linea_excesos_potencia.Count > 0)
                {
                    for (int i = 0; i < linea_excesos_potencia.Count; i++)
                        total = total + linea_excesos_potencia[i].total;

                    sb.Append(total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append(" null,");

                #endregion

                #region Excesos Reactiva
                sb.Append("null,");
                #endregion

                #region Suplemento Territorial
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> lineas_sstt = factura.lista_factura_detalle.FindAll(z => z.producto == "SSTT");
                if (lineas_sstt.Count > 0)
                {
                    for (int i = 0; i < lineas_sstt.Count; i++)
                        total = total + lineas_sstt[i].total;

                    sb.Append(total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append(" null,");
                #endregion

                #region Servicio Gestion Preferente
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> gest_pref = factura.lista_factura_detalle.FindAll(z => z.producto == "SERV_PREF");
                if (gest_pref.Count > 0)
                {
                    for (int i = 0; i < gest_pref.Count; i++)
                        total = total + gest_pref[i].total;

                    sb.Append(total.ToString().Replace(",", ".")).Append(",");
                }
                else
                    sb.Append(" null,");
                #endregion

                #region Alquiler
                total = 0;
                List<EndesaEntity.eer.FacturaDetalle> linea_alquiler = factura.lista_factura_detalle.FindAll(z => z.producto == "ALQU");
                if (linea_alquiler.Count > 0)
                {
                    for (int i = 0; i < linea_alquiler.Count; i++)
                        total = total + linea_alquiler[i].total;

                    sb.Append(total.ToString().Replace(",", ".")).Append(")");
                }
                else
                    sb.Append(" null)");
                #endregion

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                GuardaFacturaLineas(factura.id_factura, factura.lista_factura_detalle);
                                
                GuardaNumFacturaTemporal(factura.id_factura);

            }
            catch(Exception e)
            {
                MessageBox.Show("Error a la hora de guardar la factura",
               "Guardado de Factura en BBDD",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void GuardaFacturaLineas(int id_factura, List<EndesaEntity.eer.FacturaDetalle> lista)
        {

            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            int linea_factura = 0;
            double total = 0;

            

            sb.Append("replace into eer_facturasdetalle");
            sb.Append(" (id_factura, linea, producto, concepto, calculo, cantidad,");
            sb.Append(" unidad_cantidad, precio, unidad_precio, total) values ");

            for (int i = 0; i < lista.Count(); i++)
            {
                sb.Append("(").Append(id_factura).Append(",");
                sb.Append(lista[i].linea_factura).Append(",");
                sb.Append("'").Append(lista[i].producto).Append("',");
                sb.Append("'").Append(lista[i].concepto).Append("',");
                sb.Append("'").Append(lista[i].calculo).Append("',");
                sb.Append(lista[i].cantidad.ToString().Replace(",",".")).Append(",");
                sb.Append("'").Append(lista[i].unidad_cantidad).Append("',");
                sb.Append(lista[i].precio.ToString().Replace(",", ".")).Append(",");
                sb.Append("'").Append(lista[i].unidad_precio).Append("',");
                sb.Append(lista[i].total.ToString().Replace(",", ".")).Append("),");
            }

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }


        public void RegistraFactura(string cups20, DateTime fd, DateTime fh, string codigo_factura, DateTime fechaFactura)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update eer_facturas set codigo_factura = '" + codigo_factura + "', "
                + " fecha_factura = '" + fechaFactura.ToString("yyyy-MM-dd") + "'"
                + " where cups20 = '" + cups20 + "' and"
                + " (fecha_consumo_desde = '" + fd.ToString("yyyy-MM-dd") + "'"
                + " and fecha_consumo_hasta = '" + fh.ToString("yyyy-MM-dd") + "')"
                + " and codigo_factura is null and fecha_factura is null";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Factura registrada correctamente",
               "Registro de factura",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information);
            


        }

        private void BorraFactura(int id_factura)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "delete from eer_facturasdetalle where id_factura = " + id_factura;
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "delete from eer_facturas where id_factura = " + id_factura;
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public double ImporteProducto(int id_factura, string producto)
        {
            double importe = 0;

            EndesaEntity.eer.Factura f;
            if (dic.TryGetValue(id_factura, out f))
                for(int i = 0; i < f.lista_factura_detalle.Count; i++)
                {
                    if (f.lista_factura_detalle[i].producto == producto)
                        importe += f.lista_factura_detalle[i].total;
                }

            return importe;
        }

        public int GetLastID()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int ultimo_id = 0;

            strSql = "SELECT MAX(id_factura) id_factura FROM eer_facturas";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                ultimo_id = Convert.ToInt32(r["id_factura"]);
            }
            db.CloseConnection();

            return ultimo_id + 1;
        }


        public void GuardaFacturaManual(EndesaEntity.eer.Factura factura)
        {
            
            MySQLDB db;
            MySqlCommand command;

            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Append("replace into eer_facturas");
                sb.Append(" (id_factura, nif, razon_social, codigo_factura, fecha_factura, fecha_consumo_desde, fecha_consumo_hasta,");
                sb.Append(" consumo_activa, consumo_reactiva, cups20, tarifa,");
                sb.Append(" direccion_fiscal, direccion_suministro, direccion_envio, base_ise, impuesto_electricidad,");
                sb.Append(" base_ise_reducido, impuesto_electricidad_reducido, base_imponible, iva, total_factura,");
                sb.Append(" termino_energia, descuento_energia, facturacion_potencia, recargo_excesos,");
                sb.Append(" reactiva, suplemento_territorial, servicio_gestion_preferente, alquiler");
                sb.Append(") values (");
                sb.Append(factura.id_factura).Append(",");
                sb.Append("'").Append(factura.nif).Append("',");
                sb.Append("'").Append(factura.nombre_cliente).Append("',");
                sb.Append("'").Append(factura.codigo_factura).Append("',");
                sb.Append("'").Append(factura.fecha_factura.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(factura.fecha_consumo_desde.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(factura.fecha_consumo_hasta.ToString("yyyy-MM-dd")).Append("',");
                sb.Append(factura.consumo_activa.ToString().Replace(",", ".")).Append(",");
                sb.Append(factura.consumo_reactiva.ToString().Replace(",", ".")).Append(",");
                sb.Append("'").Append(factura.cupsree).Append("',");
                sb.Append("'").Append(factura.tarifa).Append("',");

                if (factura.direccion_facturacion != null)
                    sb.Append("'").Append(factura.direccion_facturacion).Append("',");
                else
                    sb.Append(" null,");

                if (factura.direccion_suministro != null)
                    sb.Append("'").Append(factura.direccion_suministro).Append("',");
                else
                    sb.Append(" null,");

                if (factura.direccion_envio != null)
                    sb.Append("'").Append(factura.direccion_envio).Append("',");
                else
                    sb.Append(" null,");

                if (factura.base_ise != 0)
                    sb.Append(factura.base_ise.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.impuesto_electricidad != 0)
                    sb.Append(factura.impuesto_electricidad.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.base_ise_reducido != 0)
                    sb.Append(factura.base_ise_reducido.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.impuesto_electricidad_reducido != 0)
                    sb.Append(factura.impuesto_electricidad_reducido.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.base_imponible != 0)
                    sb.Append(factura.base_imponible.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.iva != 0)
                    sb.Append(factura.iva.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.total_factura != 0)
                    sb.Append(factura.total_factura.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.termino_energia != 0)
                    sb.Append(factura.termino_energia.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.descuento_energia != 0)
                    sb.Append(factura.descuento_energia.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.facturacion_potencia != 0)
                    sb.Append(factura.facturacion_potencia.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.recargo_excesos != 0)
                    sb.Append(factura.recargo_excesos.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.recargo_excesos_reactiva != 0)
                    sb.Append(factura.recargo_excesos_reactiva.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                if (factura.suplemento_territorial != 0)
                    sb.Append(factura.suplemento_territorial.ToString().Replace(",", ".")).Append(",");
                else
                    sb.Append(" null,");

                // Servicio servicio_gestion_preferente
                sb.Append(" null,");

                if (factura.alquiler != 0)
                    sb.Append(factura.alquiler.ToString().Replace(",", ".")).Append(")");
                else
                    sb.Append(" null)");


                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                MessageBox.Show("Factura registrada correctamente",
                    "Registro de factura",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

            }
            catch(Exception ex)
            {
                MessageBox.Show("Factura No registrada",
              "Registro de factura",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

           
        }
    }
}
