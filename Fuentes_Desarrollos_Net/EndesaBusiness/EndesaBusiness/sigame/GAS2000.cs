using EndesaBusiness.servidores;
using EndesaEntity.facturacion;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaBusiness.sigame
{
    public class GAS2000
    {

        public string motivo_pendiente { get; set; }
        public bool telemedida { get; set; }
        public bool top { get; set; }
        public bool existe_id_pto_suministro { get; set; }
        public string nombre_cliente { get; set; }
        public string mes_medida { get; set; }
        public DateTime fecha_alta { get; set; }


        Dictionary<int, EndesaEntity.medida.GAS2000_cliente> dic_clientes;
        Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas> dic_medidas_facturas;
        public GAS2000()
        {
            CopiaClientes();
            dic_clientes = CargaClientes();
            dic_medidas_facturas = CargaMedidasyFacturas();
        }
        
        public void CopiaClientes()
        {
            OleDbDataReader rs;
            servidores.AccessDB acc;
            string strSql = "";

            List<EndesaEntity.medida.GAS2000_cliente> lista = new List<EndesaEntity.medida.GAS2000_cliente>();
            try
            {
                strSql = "select IdPtoSuministro, CUPS, CLIENTE,  Clientes.[Fecha Alta], Clientes.[Fecha Baja], UM, TELEMEDIDA,"
                    + " [Motivo Pendiente], TOP"
                    + " from Clientes";
                acc = new servidores.AccessDB(@"\\espmad01nas1.esp.e-corpnet.org\mad_grp30\GAS_AP\GAS2000.accdb");
                OleDbCommand cmd = new OleDbCommand(strSql, acc.con);
                rs = cmd.ExecuteReader();
                while (rs.Read())
                {
                    EndesaEntity.medida.GAS2000_cliente c = new EndesaEntity.medida.GAS2000_cliente();

                    c.id_pto_suministro = Convert.ToInt32(rs["IdPtoSuministro"]);

                    if (rs["CUPS"] != System.DBNull.Value)
                        c.cups = rs["CUPS"].ToString();

                    if (rs["CLIENTE"] != System.DBNull.Value)
                        c.cliente = rs["CLIENTE"].ToString();

                    c.telemedida = rs["TELEMEDIDA"].ToString() == "True";

                    if (rs["Fecha Alta"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(rs["Fecha Alta"]);


                    if (rs["TOP"] != System.DBNull.Value)
                        c.top = rs["TOP"].ToString().ToUpper() != "NO";
                    else
                        c.top = false;

                    if (rs["Motivo Pendiente"] != System.DBNull.Value)
                        c.motivo_pendiente = rs["Motivo Pendiente"].ToString();

                    lista.Add(c);

                }
                acc.CloseConnection();

                GuardarEnMySQL(lista);
            }catch(Exception e)
            {
                
                Console.WriteLine(e.Message);
            }
            
        }

        private void GuardarEnMySQL(List<EndesaEntity.medida.GAS2000_cliente> l)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int registros = 0;
            int totalregistros = 0;            
            
            try
            {
                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand("delete from cm_clientes_gas2000", db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                Console.WriteLine("Guardando registros en SIOPE.fact.cm_clientes_gas2000");
                foreach (EndesaEntity.medida.GAS2000_cliente p in l)
                {
                    registros++;
                    totalregistros++;
                    if (firstOnly)
                    {
                        sb.Append("replace into cm_clientes_gas2000");
                        sb.Append(" (cliente, cif, id_pto_suministro, cups,");
                        sb.Append(" cisternas, consumoscliente, qcontratado1,");
                        sb.Append(" fecha_alta, fecha_baja, um, telemedida, top, motivo_pendiente) values ");
                        firstOnly = false;                        
                    }

                    if (p.cliente != null)
                        sb.Append("('").Append(utilidades.FuncionesTexto.ArreglaAcentos(p.cliente)).Append("',");
                    else
                        sb.Append("(null,");                    

                    if (p.cif != null)
                        sb.Append("'").Append(p.cif).Append("',");
                    else
                        sb.Append("null,");

                    if (p.id_pto_suministro != 0)
                        sb.Append(p.id_pto_suministro).Append(",");
                    else
                        sb.Append("null,");

                    if (p.cups != null)
                        sb.Append("'").Append(p.cups).Append("',");
                    else
                        sb.Append("null,");

                    sb.Append(p.cisternas ? "'S'" : "'N'").Append(",");
                    sb.Append(p.consumoscliente ? "'S'" : "'N'").Append(",");

                    if (p.qcontratado1 != 0)
                        sb.Append(p.qcontratado1).Append(",");
                    else
                        sb.Append("null,");

                    if (p.fecha_alta > DateTime.MinValue)
                        sb.Append("'").Append(p.fecha_alta.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.fecha_baja > DateTime.MinValue)
                        sb.Append("'").Append(p.fecha_baja.ToString("yyyy-MM-dd")).Append("',");
                    else
                        sb.Append("null,");

                    if (p.um != 0)
                        sb.Append(p.um).Append(",");
                    else
                        sb.Append("null,");

                    sb.Append(p.telemedida ? "'S'" : "'N'").Append(",");
                    sb.Append(p.top ? "'S'" : "'N'").Append(",");

                    if (p.motivo_pendiente != null)
                        sb.Append("'").Append(p.motivo_pendiente).Append("'),");
                    else
                        sb.Append("null),");

                    if (registros == 100)
                    {
                        Console.CursorLeft = 0;
                        Console.Write("Guardando " + totalregistros + " de " + l.Count);
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        registros = 0;
                    }
                }

                if (registros > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write("Guardando " + totalregistros + " de " + l.Count);
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    registros = 0;
                }


            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private Dictionary<int, EndesaEntity.medida.GAS2000_cliente> CargaClientes()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<int, EndesaEntity.medida.GAS2000_cliente> d = new Dictionary<int, EndesaEntity.medida.GAS2000_cliente>();

            try
            {
                Console.WriteLine("Cargando datos de fact.cm_clientes_gas2000");
                strSql = "SELECT cliente, cif, id_pto_suministro, cups, cisternas, consumoscliente, qcontratado1,"
                    + " fecha_alta, fecha_baja, um, telemedida, top, motivo_pendiente"
                    + " FROM fact.cm_clientes_gas2000";
                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.GAS2000_cliente c = new EndesaEntity.medida.GAS2000_cliente();

                    c.id_pto_suministro = Convert.ToInt32(r["id_pto_suministro"]);
                    if (r["cliente"] != System.DBNull.Value)
                        c.cliente = r["cliente"].ToString();

                    if (r["cif"] != System.DBNull.Value)
                        c.cif = r["cif"].ToString();

                    if (r["cups"] != System.DBNull.Value)
                        c.cups = r["cups"].ToString();

                    if (r["cisternas"] != System.DBNull.Value)
                        c.cisternas = r["cisternas"].ToString() == "S";

                    if (r["consumoscliente"] != System.DBNull.Value)
                        c.consumoscliente = r["consumoscliente"].ToString() == "S";

                    if (r["qcontratado1"] != System.DBNull.Value)
                        c.qcontratado1 = Convert.ToInt32(r["qcontratado1"]);

                    if (r["fecha_alta"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["fecha_alta"]);

                    if (r["fecha_baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["fecha_baja"]);

                    if (r["um"] != System.DBNull.Value)
                        c.um = Convert.ToInt32(r["um"]);

                    if (r["telemedida"] != System.DBNull.Value)
                        c.telemedida = r["telemedida"].ToString() == "S";

                    if (r["top"] != System.DBNull.Value)
                        c.top = r["top"].ToString() == "S";

                    if (r["motivo_pendiente"] != System.DBNull.Value)
                        c.motivo_pendiente = r["motivo_pendiente"].ToString();

                    
                    EndesaEntity.medida.GAS2000_cliente o;
                    if (!d.TryGetValue(c.id_pto_suministro, out o))                        
                        d.Add(c.id_pto_suministro, c);
                    
                }
                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }

        }

        public void GetId_pto_suministro(int id_pto_suministro)
        {
            EndesaEntity.medida.GAS2000_cliente o;
            if (dic_clientes.TryGetValue(id_pto_suministro, out o))
            {

                motivo_pendiente = o.motivo_pendiente;
                telemedida = o.telemedida;
                existe_id_pto_suministro = true;
                nombre_cliente = o.cliente;
                top = o.top;
                fecha_alta = o.fecha_alta;

                EndesaEntity.medida.GAS2000_Medidas_y_Facturas oo;
                if (dic_medidas_facturas.TryGetValue(id_pto_suministro, out oo))
                {
                    mes_medida = oo.mes;
                }
            }
            else
                existe_id_pto_suministro = false;

        }

        private Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas> _CargaMedidasyFacturas()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas> d = new Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas>();
            try
            {

                Console.WriteLine("Cargando datos de fact.cm_medidas_facturas");
                strSql = "select a.`IdPto Medida` as ID, a.Mes, a.Comentario, a.Facturacion"
                    + " from cm_medidas_facturas a inner join"
                    + " (select b.`IdPto Medida`, max(Mes) Mes from cm_medidas_facturas b group by b.`IdPto Medida`) as b"
                    + " on a.`IdPto Medida` = b.`IdPto Medida` and"
                    + " a.mes = b.mes"
                    + " where a.Mes > date_format(last_day((now() - interval 2 year)),'%Y%m');";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.GAS2000_Medidas_y_Facturas c = new EndesaEntity.medida.GAS2000_Medidas_y_Facturas();
                    c.id_pto_suministro = Convert.ToInt32(r["ID"]);
                    if (r["Mes"] != System.DBNull.Value)
                        c.mes = r["Mes"].ToString();
                    if (r["Comentario"] != System.DBNull.Value)
                        c.comentario = r["Comentario"].ToString();
                    if (r["Facturacion"] != System.DBNull.Value)
                        c.facturacion = r["Facturacion"].ToString();

                    EndesaEntity.medida.GAS2000_Medidas_y_Facturas o;
                    if (!d.TryGetValue(c.id_pto_suministro, out o))
                        d.Add(c.id_pto_suministro, c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas> CargaMedidasyFacturas()
        {

            OleDbDataReader r;
            servidores.AccessDB acc;
            string strSql = "";
            
            Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas> d = new Dictionary<int, EndesaEntity.medida.GAS2000_Medidas_y_Facturas>();
            try
            {

                Console.WriteLine("Cargando datos de fact.cm_medidas_facturas");
                strSql = "select [Medidas y Facturas].[Medidas y Facturas] as id_ps,"
                    + " [Medidas y Facturas].Mes, [Medidas y Facturas].Factura"
                    + " [Medidas y Facturas].Comentario, [Medidas y Facturas].Facturacion"
                    + " [Medidas y Facturas].[Desc TUR %] as desc_tur_porcentaje,"
                    + " [Medidas y Facturas].[Desc TUR €] as desc_tur_importe,"
                    + " [Medidas y Facturas].Carga,"
                    + " [Medidas y Facturas].[Clausula Flexibilidad] as clausula_flexibilidad"
                    + " from [Medidas y Facturas]";
                    
                   
                acc = new servidores.AccessDB(@"\\espmad01nas1.esp.e-corpnet.org\mad_grp30\GAS_AP\GAS2000.accdb");
                OleDbCommand cmd = new OleDbCommand(strSql, acc.con);
                r = cmd.ExecuteReader();
                while (r.Read())                   
                {
                    EndesaEntity.medida.GAS2000_Medidas_y_Facturas c = new EndesaEntity.medida.GAS2000_Medidas_y_Facturas();
                    c.id_pto_suministro = Convert.ToInt32(r["id_ps"]);
                    if (r["Mes"] != System.DBNull.Value)
                        c.mes = r["Mes"].ToString();
                    if (r["Comentario"] != System.DBNull.Value)
                        c.comentario = r["Comentario"].ToString();
                    if (r["Facturacion"] != System.DBNull.Value)
                        c.facturacion = r["Facturacion"].ToString();

                    EndesaEntity.medida.GAS2000_Medidas_y_Facturas o;
                    if (!d.TryGetValue(c.id_pto_suministro, out o))
                        d.Add(c.id_pto_suministro, c);
                }
                acc.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
}
