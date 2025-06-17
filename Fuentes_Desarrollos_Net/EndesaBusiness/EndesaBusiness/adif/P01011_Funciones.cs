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
    public class P01011_Funciones
    {
        Dictionary<string, EndesaEntity.Adif_Medida_Horaria> dic_adif = 
            new Dictionary<string, EndesaEntity.Adif_Medida_Horaria>();
        CupsFunciones cf;


        public P01011_Funciones()
        {
            cf = new CupsFunciones();
        }

        public void CargaP01011(string archivo)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            FileInfo f = new FileInfo(archivo);

            strSql = "LOAD DATA LOCAL INFILE '" + archivo.Replace(@"\", "\\\\")
                + " ' REPLACE INTO TABLE adif_PO1011_temp"
                + " FIELDS TERMINATED BY ';' LINES TERMINATED BY '\n'"
                + " (@CUPSREE,@VAL,@FECHAHORA,@ESTACION,@AE,@CAL_AE,"
                + " @AES,@CAL_AES,@R1,@CAL_R1,@R2,@CAL_R2,@R3,@CAL_R3,"
                + " @R4,@CAL_R4,@RESERVA1,@CAL_RESERVA1,@RESERVA2,@CAL_RESERVA2) SET"
                + " CUPSREE = @CUPSREE,"
                + " VAL = @VAL,"
                + " FECHAHORA = @FECHAHORA,"
                + " ESTACION = @ESTACION,"
                + " AE = @AE,"
                + " CAL_AE = @CAL_AE,"
                + " AES = @AES,"
                + " CAL_AES = @CAL_AES,"
                + " R1 = @R1,"
                + " CAL_R1 = @CAL_R1,"
                + " R2 = @R2,"
                + " CAL_R2 = @CAL_R2,"
                + " R3 = @R3,"
                + " CAL_R3 = @CAL_R3,"
                + " R4 = @R4,"
                + " CAL_R4 = @CAL_R4,"
                + " RESERVA1 = @RESERVA1,"
                + " CAL_RESERVA1 = @CAL_RESERVA1,"
                + " RESERVA2 = @RESERVA2,"
                + " CAL_RESERVA2 = @CAL_RESERVA2,"
                + " FechaCarga = '" + DateTime.Now.ToString("yyyy-MM-dd") + "',"
                + " Fichero = '" + f.Name + "',"
                + " User = '" + System.Environment.UserName + "';";

            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void GuardaP01011(string archivo)
        {
            System.IO.StreamReader fileStream;
            string line;
            string[] c;
            EndesaEntity.Adif_Medida_Horaria adif;
            int hora = 0;
            bool invierno = false;
            string cups20 = "";

            DateTime fecha = new DateTime();
            int periodo = 0;

            try
            {
                FileInfo file = new FileInfo(archivo);
                fileStream = new System.IO.StreamReader(file.FullName, System.Text.Encoding.GetEncoding(1252));
                while ((line = fileStream.ReadLine()) != null)
                {
                    c = line.Split(';');
                    if (c.Count() > 10)
                    {
                        cups20 = c[0];
                        //cf.GetFromCups20(cups20);
                        fecha = new DateTime(Convert.ToInt32(c[2].Substring(0, 4)), Convert.ToInt32(c[2].Substring(5, 2)), Convert.ToInt32(c[2].Substring(8, 2)));
                        hora = Convert.ToInt32(c[2].Substring(11, 2));
                        if (hora == 0)
                        {
                            hora = 23;
                            fecha = fecha.AddDays(-1);
                        }
                        else
                            hora = hora - 1;
                        invierno = Convert.ToInt32(c[3]) == 0;
                        periodo = this.GetPeriodoHorario(fecha, hora, invierno);

                        // Activa
                        #region Activa
                        EndesaEntity.Adif_Medida_Horaria aa;
                        if (!dic_adif.TryGetValue(cups20 + fecha.ToString("yyyyMMdd") + "A", out aa))
                        {

                            adif = new EndesaEntity.Adif_Medida_Horaria();
                            //adif.id = cf.id;
                            //adif.cups13 = cf.cups13;
                            adif.cup20 = cups20;
                            adif.fecha = fecha;
                            adif.fechaCarga = DateTime.Now;
                            adif.tipoEnergia = "A";
                            adif.numPeriodos = this.NumPeriodosHorarios(fecha);
                            adif.fichero = file.Name;
                            adif.value[periodo] = Convert.ToInt32(c[4]);
                            adif.total += Convert.ToInt32(c[4]);
                            adif.c[periodo] = c[5].ToString();
                            dic_adif.Add(cups20 + fecha.ToString("yyyyMMdd") + "A", adif);
                        }
                        else
                        {
                            aa.value[periodo] = Convert.ToInt32(c[4]);
                            aa.total += Convert.ToInt32(c[4]);
                            aa.c[periodo] = c[5].ToString();
                        }
                        #endregion

                        // Reactiva
                        #region Reactiva
                        if (!dic_adif.TryGetValue(cups20 + fecha.ToString("yyyyMMdd") + "R", out aa))
                        {
                            adif = new EndesaEntity.Adif_Medida_Horaria();
                            //adif.id = cf.id;
                            //adif.cups13 = cf.cups13;
                            adif.cup20 = cups20;
                            adif.fecha = fecha;
                            adif.fechaCarga = DateTime.Now;
                            adif.tipoEnergia = "R";
                            adif.numPeriodos = this.NumPeriodosHorarios(fecha);
                            adif.fichero = file.Name;
                            adif.value[periodo] = Convert.ToInt32(c[8]);
                            adif.total = Convert.ToInt32(c[8]);
                            adif.c[periodo] = c[9].ToString();
                            dic_adif.Add(cups20 + fecha.ToString("yyyyMMdd") + "R", adif);
                        }
                        else
                        {
                            aa.value[periodo] = Convert.ToInt32(c[8]);
                            aa.total += Convert.ToInt32(c[8]);
                            aa.c[periodo] = c[9].ToString();
                        }
                        #endregion

                    }
                }
                fileStream.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "P01011.GuardaP01011",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }

        }


        

        public void Save()
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int numReg = 0;

            try
            {
                foreach (KeyValuePair<string, EndesaEntity.Adif_Medida_Horaria> p in dic_adif)
                {
                    numReg++;
                    if (firstOnly)
                    {
                        sb.Append("replace into adif_medida_horaria_adif_temp ");
                        sb.Append("(cups20,Fecha,TipoEnergia,Unidad,Total,");
                        sb.Append("Value1,Value2,Value3,Value4,Value5,Value6,Value7,Value8,Value9,Value10,");
                        sb.Append("Value11,Value12,Value13,Value14,Value15,Value16,Value17,Value18,Value19,");
                        sb.Append("Value20,Value21,Value22,Value23,Value24,Value25,");
                        sb.Append("C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,");
                        sb.Append("C17,C18,C19,C20,C21,C22,C23,C24,C25,");
                        sb.Append("FechaCarga,Fichero,User) values ");
                        firstOnly = false;
                    }

                    
                    sb.Append("('").Append(p.Value.cup20).Append("',");
                    sb.Append("'").Append(p.Value.fecha.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(p.Value.tipoEnergia).Append("',");
                    sb.Append("null").Append(",");
                    sb.Append(p.Value.total);

                    for (int j = 1; j <= 25; j++)
                    {
                        if (p.Value.numPeriodos == 24 && j == 25)
                            sb.Append(" ,null");
                        else if (p.Value.numPeriodos == 23 && (j == 2 || j == 25))
                            sb.Append(" ,null");
                        else
                            sb.Append(" ,").Append(p.Value.value[j]);
                    }

                    for (int j = 1; j <= 25; j++)
                    {
                        if (p.Value.numPeriodos == 24 && j == 25)
                            sb.Append(" ,null");
                        else if (p.Value.numPeriodos == 23 && (j == 2 || j == 25))
                            sb.Append(" ,null");
                        else
                            sb.Append(" ,").Append(p.Value.c[j]);

                    }
                    sb.Append(" ,'").Append(p.Value.fechaCarga.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(p.Value.fichero).Append("',");
                    sb.Append("'").Append(System.Environment.UserName).Append("'),");

                    if (numReg == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.MED);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }

                }

                if (numReg > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    numReg = 0;
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "P01011.Save",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
        }

        private int GetPeriodoHorario(DateTime dia, int hora, bool invierno)
        {
            if (NumPeriodosHorarios(dia) == 25 && invierno)
                return hora + 2;
            else
                return hora + 1;

        }

        private int NumPeriodosHorarios(DateTime f)
        {
            return (f.Month == 3 ?
                    f.Equals(UltimoDomingoMarzo(f)) ? 23 : 24 :
                    f.Month == 10 ?
                    f.Equals(UltimoDomingoOctubre(f)) ? 25 : 24 : 24);
        }

        private DateTime UltimoDomingoMarzo(DateTime f)
        {
            return new DateTime(f.Year, 3, 31).AddDays(-(new DateTime(f.Year, 3, 31).DayOfWeek - DayOfWeek.Sunday));
        }

        private DateTime UltimoDomingoOctubre(DateTime f)
        {
            return new DateTime(f.Year, 10, 31).AddDays(-(new DateTime(f.Year, 10, 31).DayOfWeek - DayOfWeek.Sunday));
        }
    }
}
