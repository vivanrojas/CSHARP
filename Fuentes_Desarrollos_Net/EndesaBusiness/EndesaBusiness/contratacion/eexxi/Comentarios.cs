using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EndesaBusiness.servidores;
using EndesaEntity.contratacion.xxi;
using MySql.Data.MySqlClient;

namespace EndesaBusiness.contratacion.eexxi
{
    public class Comentarios : Comentario  
    {
        public Dictionary<string, List<Comentario>> dic_comentarios { get; set; }

        public enum EditStatus
        {
            NEW,
            EDIT
        }
       
        public Comentarios()
        {
            dic_comentarios = CargaDatos(null);
        }

       

        public Comentarios(string codigo)
        {
            List<string> l = new List<string>();
            l.Add(codigo);
            dic_comentarios = CargaDatos(l);        
        }

        private Dictionary<string, List<Comentario>> CargaDatos(List<string> lista_codigos)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;

            Dictionary<string, List<Comentario>> d = new Dictionary<string, List<Comentario>>();

            try
            {

                strSql = "select c.id_comentario, c.linea, c.comentario,"
                + " c.created_by, c.created_date, c.last_update_by, c.last_update_date"
                + " from eexxi_comentarios c";
                
                if(lista_codigos != null)
                    for(int i = 0; i < lista_codigos.Count; i++)
                    {
                        if (firstOnly)
                        {
                            strSql += " where c.id_comentario in ("
                                + "'" + lista_codigos[i] + "'";
                            firstOnly = false;
                        }
                        else
                            strSql += ",'" + lista_codigos[i] + "'";
                    }

                if (lista_codigos != null)
                    strSql += ") ORDER BY c.id_comentario, c.linea, c.created_date desc";
                else
                    strSql += " ORDER BY c.id_comentario, c.linea, c.created_date desc";


                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Comentario c = new Comentario();

                    
                    c.id_comentario = r["id_comentario"].ToString();
                    c.linea = Convert.ToInt32(r["linea"]);

                    if (r["comentario"] != System.DBNull.Value)
                        c.comentario = r["comentario"].ToString();

                    if (r["created_by"] != System.DBNull.Value)
                        c.created_by = r["created_by"].ToString();                    

                    if (r["created_date"] != System.DBNull.Value)
                        c.created_date = Convert.ToDateTime(r["created_date"]);

                    if (r["last_update_by"] != System.DBNull.Value)
                        c.last_update_by = r["last_update_by"].ToString();

                    if (r["last_update_date"] != System.DBNull.Value)
                        c.last_update_date = Convert.ToDateTime(r["last_update_date"]);

                    List<Comentario> o;
                    if (!d.TryGetValue(c.id_comentario, out o))
                    {
                        o = new List<Comentario>();
                        o.Add(c);
                        d.Add(c.id_comentario, o);
                    }
                    else
                        o.Add(c);
                    
                   

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Comentarios - CargaDatos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
                return null; 
            }
        }

        public string UltimoComentario(string cups, string codigodesolicitud)
        {
            List<Comentario> o;
            if (dic_comentarios.TryGetValue(cups + codigodesolicitud, out o))
                return o.Last().comentario;
            else
                return null;
        }

        public int UltimaLinea(string id_comentario)
        {
            List<Comentario> o;
            if (dic_comentarios.TryGetValue(id_comentario, out o))
                return o.Last().linea;
            else
                return 0;
        }

        public List<Comentario> Lista_Comentarios(string id_comentario)
        {
            List<Comentario> o;
            if (dic_comentarios.TryGetValue(id_comentario, out o))
                return o;
            else
                return null;
        }

        public void Del(string codigoSolicitud, int linea)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "delete from eexxi_comentarios where id_comentario = '" + codigoSolicitud + "'"
                    + " and linea = " + linea;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el borrado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Add()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "replace into eexxi_comentarios (id_comentario, linea, comentario,"
                    + " created_by, created_date) values ";

                strSql += "('" + this.id_comentario + "',"                    
                    + this.linea + ","
                    + "'" + this.comentario + "',"
                    + "'" + System.Environment.UserName + "',"                    
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el guardado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update eexxi_comentarios set"
                    + " comentario = '" + this.comentario + "',"
                    + " last_update_by = '" + System.Environment.UserName + "'"
                    + " where id_comentario = '" + this.id_comentario + "' and"                    
                    + " linea = " + this.linea;

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la actualización de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
    }
}
