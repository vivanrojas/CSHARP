﻿using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static EndesaBusiness.medida.Kee_Extraccion_Formulas;

namespace EndesaBusiness.medida
{
    public class EstadosSubestadosKronos : EndesaEntity.medida.Pendiente
    {
        public Dictionary<string, EndesaEntity.medida.Pendiente> dic { get; set; }
        public bool existe { get; set; }
        public EstadosSubestadosKronos()
        {
            dic = Carga();
        }

        public EstadosSubestadosKronos(string strBTN )
        {
            if (strBTN != "")
            {
                dic = CargaBTN();
            }
    
           
        }


        private Dictionary<string, EndesaEntity.medida.Pendiente> CargaBTN()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.medida.Pendiente> d = new Dictionary<string, EndesaEntity.medida.Pendiente>();
            try
            {
                strSql = "SELECT codigo_estado, estado_periodo, area_responsable,"
                    + " prioridad,  estado, codigo_subestado,subestado, ESTADO_GLOBAL, ESTADO_GLOBAL_A_REPORTAR,comentario"
                    + " FROM estados_kee_param_BTN "
                    + " order by codigo_estado";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_estado = r["codigo_estado"].ToString();
                    c.estado_periodo = r["estado_periodo"].ToString();
                    c.area_responsable = r["area_responsable"].ToString();
                    ////c.prioridad = Convert.ToInt32(r["prioridad"]);
                    c.cod_subestado = r["codigo_subestado"].ToString();
                    c.descripcion_subestado = r["subestado"].ToString();
                    c.descripcion_estado = r["estado"].ToString();
                    c.ESTADO_GLOBAL = r["ESTADO_GLOBAL"].ToString();
                    c.ESTADO_GLOBAL_A_REPORTAR = r["ESTADO_GLOBAL_A_REPORTAR"].ToString();

                    if (r["comentario"] != System.DBNull.Value)
                        c.comentario_revision_medida = r["comentario"].ToString();

                    d.Add(c.estado_periodo.ToUpper() , c);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                  "EstadosSubestadosKronos.Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);

                return null;
            }
        }
        private Dictionary<string, EndesaEntity.medida.Pendiente> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.medida.Pendiente> d = new Dictionary<string, EndesaEntity.medida.Pendiente>();
            try
            {
                strSql = "SELECT codigo_estado, estado_periodo, area_responsable,"
                    + " prioridad, subestado, estado, comentario_revision_medida"
                    + " FROM estados_kee_param order by prioridad";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_estado = r["codigo_estado"].ToString();
                    c.estado_periodo = r["estado_periodo"].ToString();
                    c.area_responsable = r["area_responsable"].ToString();
                    c.prioridad = Convert.ToInt32(r["prioridad"]);
                    c.descripcion_subestado = r["subestado"].ToString();
                    c.descripcion_estado = r["estado"].ToString();

                    if (r["comentario_revision_medida"] != System.DBNull.Value)
                        c.comentario_revision_medida = r["comentario_revision_medida"].ToString();

                    d.Add(c.estado_periodo.ToUpper(), c);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                  "EstadosSubestadosKronos.Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);

                return null;
            }
        }

        public void Save()
        {
            try
            {
                EndesaEntity.medida.Pendiente o;
                if (!dic.TryGetValue(this.cod_estado, out o))
                    this.New();
                else
                    this.Update();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "EstadosSubestadosKronos.Save",
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
                strSql = "insert into t_ed_p_subestado_sap_pendiente_facturar (cd_subestado, de_subestado,"
                    + " area_responsable, created_by, created_date) values ";

                if (this.cod_subestado != null)
                    strSql += "('" + this.cod_subestado + "'";
                else
                    strSql += "(null";

                if (this.descripcion_subestado != null)
                    strSql += ",'" + this.descripcion_subestado + "'";
                else
                    strSql += ",null";

                if (this.area_responsable != null)
                    strSql += ",'" + this.area_responsable + "'";
                else
                    strSql += ",null";

                if (this.created_by != null)
                    strSql += ",'" + this.created_by + "'";
                else
                    strSql += ",'" + System.Environment.UserName + "'";

                strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                db.MySQLTransaction();
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.MySQLCommit();
                db.CloseConnection();

                MessageBox.Show("Se ha añadido correctamente el registro con estado: " + this.cod_estado,
                 "Subestados",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "EstadosSubestadosKronos.New",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }

        private void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update estados_kee_param set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            if (this.cod_estado != null)
                strSql += " ,codigo_estado = '" + this.cod_estado + "'";

            if (this.estado_periodo != null)
                strSql += " ,estado_periodo = '" + this.estado_periodo + "'";

            if (this.area_responsable != null)
                strSql += " ,area_responsable = '" + this.area_responsable + "'";

            if (this.prioridad != 0)
                strSql += " ,prioridad = " + this.prioridad;

            if (this.descripcion_subestado != null)
                strSql += " ,subestado = '" + this.descripcion_subestado + "'";

            if (this.descripcion_estado != null)
                strSql += " ,estado = '" + this.descripcion_estado + "'";

            if (this.comentario_revision_medida != null)
                strSql += " ,comentario_revision_medida = '" + this.comentario_revision_medida + "'";

            strSql += " where codigo_estado = '" + this.cod_estado + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
            db.MySQLTransaction();
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.MySQLCommit();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el subestado: " + this.cod_subestado,
            "EstadosSubestadosKronos",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        }

        public void Del()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                DialogResult result = MessageBox.Show("Warning:" + System.Environment.NewLine
                    + "¿Desea borrar el registro para el código " + this.cod_estado + "?"
                    , "Borrado del registro con código " + this.cod_estado,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    strSql = "delete from estados_kee_param where cd_subestado = '" + this.cod_subestado + "'";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("registro borrado.",
                      "Subestados",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Pendiente_Subestados - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }


        private int GetPrioridad(string estado_periodo)
        {
            int prioridad = 99999999;
            EndesaEntity.medida.Pendiente o;

            if (estado_periodo.IndexOf("Discrepancia") >= 0)
            {
                String[] cadena;
                cadena = estado_periodo.Split(';');

                if (dic.TryGetValue(cadena[1].ToUpper(), out o))
                    return o.prioridad;
                else
                    return prioridad;

            }
            else {
                if (dic.TryGetValue(estado_periodo.ToUpper(), out o))
                    return o.prioridad;
                else
                    return prioridad;

            }

           
        }
        private int GetPrioridadMultipunto(string estado_periodo)
        {
            int prioridad = 0;
            EndesaEntity.medida.Pendiente o;

            if (estado_periodo.IndexOf("Discrepancia") >= 0)
            {
                String[] cadena;
                cadena = estado_periodo.Split(';');

                if (dic.TryGetValue(cadena[1].ToUpper(), out o))
                    return o.prioridad;
                else
                    return prioridad;

            }
            else
            {
                if (dic.TryGetValue(estado_periodo.ToUpper(), out o))
                    return o.prioridad;
                else
                    return prioridad;

            }


        }

        public void GetEstadoKEE(string estado_periodo)
        {

            EndesaEntity.medida.Pendiente o;

            if (dic.TryGetValue(estado_periodo.ToUpper(), out o))
            {
                existe = true;
                this.cod_estado = o.cod_estado;
                this.estado_periodo = o.estado_periodo;
                this.area_responsable = o.area_responsable;
                this.prioridad = o.prioridad;
                this.descripcion_subestado = o.descripcion_subestado;
                this.descripcion_estado = o.descripcion_estado;
            }
            else
                existe = false;

          
        }

        public void GetEstadoKEE(List<string> lista_estado_periodo)
        {
            bool firstOnly = true;
            int prioridad = 99999999;
            EndesaEntity.medida.Pendiente pendiente
                = new EndesaEntity.medida.Pendiente();

            foreach (string p in lista_estado_periodo)
            {
                if (firstOnly)
                {
                    GetEstadoKEE(p);
                    firstOnly = false;
                    prioridad = this.prioridad;
                }
                else
                {
                    if (GetPrioridad(p) < prioridad)
                        GetEstadoKEE(p);

                }



            }


        }


        public void GetEstadoKEEDetalle(List<string> lista_estado_periodo)
        {
            bool firstOnly = true;
            int prioridad = 99999999;
            EndesaEntity.medida.Pendiente pendiente
                = new EndesaEntity.medida.Pendiente();

            foreach (string p in lista_estado_periodo)
            {
                if (firstOnly)
                {
                    GetEstadoKEEDetalle(p);
                    firstOnly = false;
                    prioridad = this.prioridad;
                }
                else
                {
                    if (GetPrioridad(p) < prioridad)
                        GetEstadoKEEDetalle(p);
                }
            }
        }

        public void GetEstadoKEEDetalleBTN(List<string> lista_estado_periodo)
        {
            bool firstOnly = true;
            int prioridad = 99999999;
            EndesaEntity.medida.Pendiente pendiente
                = new EndesaEntity.medida.Pendiente();


            foreach (string p in lista_estado_periodo)
            {

                if (firstOnly)
                {
                    GetEstadoKEEDetalleBTN(p);
                    firstOnly = false;
                    prioridad = this.prioridad;
                }
                else
                {
                    if (GetPrioridad(p) < prioridad)
                        GetEstadoKEEDetalleBTN(p);
                }
            }
        }



        public void GetEstadoKEEDetalleMultipunto(List<string> lista_estado_periodo)
        {
            bool firstOnly = true;
            int prioridad = 0;

           
            EndesaEntity.medida.Pendiente pendiente
                = new EndesaEntity.medida.Pendiente();

            foreach (string p in lista_estado_periodo)
            {
                if (firstOnly)
                {
                    GetEstadoKEEDetalle(p);
                    firstOnly = false;
                    prioridad = this.prioridad;
                }
                else
                {
                    if (GetPrioridadMultipunto(p) > prioridad)
                        GetEstadoKEEDetalle(p);
                }
            }
        }


        public void GetEstadoKEEDetalleBTN(string estado_periodo)
        {
            EndesaEntity.medida.Pendiente o;

            string comparar;
            string[] cadena;

            cadena = estado_periodo.Split(';');
            //////comparar = cadena[1];
            //11-04-2025
            if (cadena[0] == "Pendiente Segmentos periodo") {
                cadena[0] = "Pendiente Segmentos periodos";
            }
            if (cadena[0] == "Validada pendiente de enrutar")
            {
                cadena[0] = "Validada pendiente de enturar";
            }
            ///FIn 11-04-2025
            if (dic.TryGetValue(cadena[0].ToUpper(), out o))
            {
                existe = true;
                this.cod_estado = o.cod_estado;
                this.estado_periodo = o.estado_periodo;
                this.area_responsable = o.area_responsable;
                this.prioridad = o.prioridad;
                this.descripcion_subestado = o.cod_subestado + " "+ o.descripcion_subestado;  
                this.descripcion_estado =  o.descripcion_estado;
                this.ESTADO_GLOBAL = o.ESTADO_GLOBAL;
                this.ESTADO_GLOBAL_A_REPORTAR = o.ESTADO_GLOBAL_A_REPORTAR;

                string[] cadenaAux = cadena[1].Split('/');
                this.fh_desde = Convert.ToDateTime(cadenaAux[0].Substring(1, 8).Substring(6, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(4, 2) + "/" + cadenaAux[0].Substring(1, 8).Substring(0, 4));
                this.fh_hasta = Convert.ToDateTime(cadenaAux[1].Substring(0, 8).Substring(6, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(4, 2) + "/" + cadenaAux[1].Substring(0, 8).Substring(0, 4));

                cadenaAux = cadena[2].Split('/');
                this.fecha_modificacion = Convert.ToDateTime(cadenaAux[0].Substring(0, 4) + "/" + cadenaAux[0].Substring(4, 2) + "/" + cadenaAux[0].Substring(6, 2));

                //Paco 09/02/2024
                this.multimedida = o.multimedida;

            }
            else

                //Para controlar si se fuerza la etiqueta (columna=Discrepancia)
                if (estado_periodo.IndexOf("Discrepancia") >= 0 || estado_periodo.IndexOf(";") >= 0)
                {
                existe = true;
                //Paco 08/03/2024
                ////////if (estado_periodo.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0)
                ////////{
                ////////    string[] cadena = estado_periodo.Split(';');
                ////////    if (dic.TryGetValue(cadena[1].ToUpper().Trim(), out o))
                ////////    {
                ////////        this.estado_periodo = o.estado_periodo;
                ////////        this.area_responsable = o.area_responsable;
                ////////        // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                ////////        //en lugar de en descripcion_estado
                ////////        this.descripcion_subestado = o.descripcion_estado;
                ////////        this.temporal = o.descripcion_subestado;
                ////////        //Paco 09/02/2024
                ////////        this.multimedida = o.multimedida;
                ////////    }
                ////////}
                ////////else
                ////////{
                ////////    if (estado_periodo.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || estado_periodo.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)
                ////////    {
                ////////        string[] cadena = estado_periodo.Split(';');
                ////////        if (dic.TryGetValue(cadena[1].ToUpper().Trim(), out o))
                ////////        {
                ////////            this.estado_periodo = o.estado_periodo;
                ////////            this.area_responsable = o.area_responsable;
                ////////            // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                ////////            //en lugar de en descripcion_estado
                ////////            this.descripcion_subestado = o.descripcion_estado;
                ////////            //Paco 09/02/2024
                ////////            this.multimedida = o.multimedida;
                ////////            // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el subestado de KEE lo he guardado en la variable temporal
                ////////            //en lugar de en descripcion_estado
                ////////            this.temporal = o.descripcion_subestado;
                ////////            this.prioridad = o.prioridad;
                ////////        }
                ////////    }
                
         
                //fin Paco 08/03/2024
                this.descripcion_estado = estado_periodo;
            }
            else
            {
                existe = false;
            }

        }

        public void GetEstadoKEEDetalle(string estado_periodo)
        {
            EndesaEntity.medida.Pendiente o;
            if (dic.TryGetValue(estado_periodo.ToUpper(), out o))
            {
                existe = true;
                this.cod_estado = o.cod_estado;
                this.estado_periodo = o.estado_periodo;
                this.area_responsable = o.area_responsable;
                this.prioridad = o.prioridad;
                this.descripcion_subestado = o.descripcion_subestado;
                this.descripcion_estado = o.descripcion_estado;

                //Paco 09/02/2024
                this.multimedida= o.multimedida;

            }
            else

                //Para controlar si se fuerza la etiqueta (columna=Discrepancia)
                if (estado_periodo.IndexOf("Discrepancia") >= 0 || estado_periodo.IndexOf(";")>=0)
                {
                    existe = true;
                    //Paco 08/03/2024
                    if (estado_periodo.IndexOf("Periodo SAP encontrado parcialmente en el informe pendiente de KEE") >= 0 )
                    {
                        string[] cadena = estado_periodo.Split(';');
                        if (dic.TryGetValue(cadena[1].ToUpper().Trim(), out o))
                        {
                            this.estado_periodo = o.estado_periodo;
                            this.area_responsable = o.area_responsable;
                            // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                            //en lugar de en descripcion_estado
                            this.descripcion_subestado = o.descripcion_estado;
                            this.temporal = o.descripcion_subestado;
                            //Paco 09/02/2024
                            this.multimedida = o.multimedida;
                        }
                    }
                    else {
                        if (estado_periodo.IndexOf("Periodos del pendiente de KEE contenidos en las fechas de SAP") >= 0 || estado_periodo.IndexOf("Periodo SAP contenido dentro de un informe pendiente de KEE") >= 0)
                        {
                            string[] cadena = estado_periodo.Split(';');
                            if (dic.TryGetValue(cadena[1].ToUpper().Trim(), out o))
                            {
                                this.estado_periodo = o.estado_periodo;
                                this.area_responsable = o.area_responsable;
                                // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el estado de KEE lo he guardado en descripcion_subestado
                                //en lugar de en descripcion_estado
                                this.descripcion_subestado = o.descripcion_estado;
                                //Paco 09/02/2024
                                this.multimedida = o.multimedida;
                                // Como estoy forzando y pongo a pelo la descripcion subestado y tengo la variable ocupada, el subestado de KEE lo he guardado en la variable temporal
                                //en lugar de en descripcion_estado
                                this.temporal = o.descripcion_subestado;
                                this.prioridad = o.prioridad;
                            }
                        }
                    }
                    //fin Paco 08/03/2024
                    this.descripcion_estado = estado_periodo;                
                }
                else {
                    existe = false;
                }

        }


    }
}
