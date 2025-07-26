using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class Fichero
    {


        
        public static string UltimoViernes_MMDD()
        {
            DateTime dia = new DateTime();
            DateTime inicio = new DateTime();
            DateTime fin = new DateTime();
            //utilidades.Global g = new utilidades.Global();
            try
            {
                inicio = DateTime.Now;
                dia = DateTime.Now;
                dia = dia.AddDays(-1);
                // Mientras sea distinto de viernes
                while ((int)dia.DayOfWeek != 5)
                {
                    dia = dia.AddDays(-1);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                fin = DateTime.Now;
//                g.SaveProcess("UltimoViernes_YYMM", "ERROR: " + e.Message, inicio, inicio, fin);
            }

            return dia.ToString("MMdd");
        }


        public static DateTime SiguienteDiaHabil()
        {
            DateTime siguienteDiaHabil = new DateTime();

            siguienteDiaHabil = UltimoDiaHabil();
            for (int i = 1; i < 10; i++)
            {
                siguienteDiaHabil = siguienteDiaHabil.AddDays(1);
                if ((int)siguienteDiaHabil.DayOfWeek != 0 && (int)siguienteDiaHabil.DayOfWeek != 6 && HayPaseBatch(siguienteDiaHabil))
                {
                    return siguienteDiaHabil;
                }
            }

            return siguienteDiaHabil;
        }



        public static string UltimoDiaHabil_YYMMDD()
        {
            DateTime hoy = new DateTime();
            try
            {
                hoy = DateTime.Now;
                Console.WriteLine(hoy.ToString("yyMMdd") + " " + (int)hoy.DayOfWeek);
                hoy = hoy.AddDays(-1);
                Console.WriteLine(hoy.ToString("yyMMdd") + " " + (int)hoy.DayOfWeek);
                for (int i = 9; i > 0; i--)
                {
                    if ((int)hoy.DayOfWeek != 0 && (int)hoy.DayOfWeek != 6 && HayPaseBatch(hoy))
                    {
                        Console.WriteLine(hoy.ToString("yyMMdd") + " " + (int)hoy.DayOfWeek);
                        return hoy.ToString("yyMMdd");
                    }
                    hoy = hoy.AddDays(-1);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Peta " + e.Message);
            }
            return hoy.ToString("yyMMdd");
        }

        public static bool HayPaseBatch(DateTime fecha)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            bool haypase = true;

            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select FechaFestivo from fact.fact_diasfestivos where " +
                "FechaFestivo = '" + fecha.ToString("yyyy-MM-dd") + "';";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    haypase = false;
                }
                db.CloseConnection();
                return haypase;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


            return haypase;
        }

        public static DateTime UltimoDiaHabil()
        {
            DateTime hoy = new DateTime();
            try
            {
                hoy = DateTime.Now.Date;
                hoy = hoy.AddDays(-1);

                for (int i = 9; i > 0; i--)
                {
                    if ((int)hoy.DayOfWeek != 0 && (int)hoy.DayOfWeek != 6 && HayPaseBatch(hoy))
                    {
                        return hoy.Date;
                    }
                    hoy = hoy.AddDays(-1);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error en último día habil " + e.Message);
            }
            return hoy;
        }



        public static void EjecutaComando(string command)
        {
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = true;
                proc.StartInfo.FileName = command;
                proc.StartInfo.UseShellExecute = false;
                proc.Start();
                proc.WaitForExit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public static void EjecutaComando(string command, string argument)
        {
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = true;
                proc.StartInfo.FileName = command;
                proc.StartInfo.Arguments = argument;
                proc.StartInfo.UseShellExecute = false;
                proc.Start();
                proc.WaitForExit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        public static void ArreglaFichero(String fichero)
        {
            ProcessStartInfo p = new ProcessStartInfo();
            p.FileName = AppDomain.CurrentDomain.BaseDirectory + @"bin\ConversorCodificacionFicheros.exe";
            p.Arguments = fichero;
            p.WindowStyle = ProcessWindowStyle.Normal;
            Process x = Process.Start(p);
            x.WaitForExit();
            BorrarArchivo(fichero);
        }


        public static void ExecuteCommand(string command)
        {
            int exitCode;
            ProcessStartInfo processInfo;
            Process process;

            processInfo = new ProcessStartInfo("cmd.exe", "/c " + command);
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = true;
            // *** Redirect the output ***
            processInfo.RedirectStandardError = true;
            processInfo.RedirectStandardOutput = true;

            process = Process.Start(processInfo);
            process.WaitForExit();

            // *** Read the streams ***
            // Warning: This approach can lead to deadlocks, see Edit #2
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            exitCode = process.ExitCode;

            Console.WriteLine("output>>" + (String.IsNullOrEmpty(output) ? "(none)" : output));
            Console.WriteLine("error>>" + (String.IsNullOrEmpty(error) ? "(none)" : error));
            Console.WriteLine("ExitCode: " + exitCode.ToString(), "ExecuteCommand");
            process.Close();
        }

        public static string checkMD5(string filename)
        {
            Console.WriteLine("Comprobando md5 ...");
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filename))
                {
                    return HashString(Encoding.Default.GetString(md5.ComputeHash(stream)));
                }
            }
        }

        public static void Filetea(string filename, Boolean borrarArchivo)
        {

            StreamWriter writer = null;
            String ficheroDestino = null;
            int numFile = 1;
            try
            {
                using (StreamReader inputfile = new System.IO.StreamReader(filename))
                {
                    int count = 0;
                    string line;
                    while ((line = inputfile.ReadLine()) != null)
                    {


                        if (writer == null || count > 100000)
                        {
                            if (writer != null)
                            {
                                writer.Close();
                                writer = null;
                            }
                            ficheroDestino = Path.GetDirectoryName(filename) + Path.DirectorySeparatorChar +
                                Path.GetFileNameWithoutExtension(filename) + "_" + numFile.ToString() + ".txt";
                            Console.WriteLine("Creando fichero " + ficheroDestino);
                            writer = new System.IO.StreamWriter(ficheroDestino, true);
                            count = 0;
                            numFile++;
                        }

                        writer.WriteLine(line);

                        ++count;
                    }
                }
            }
            finally
            {
                if (writer != null)
                    writer.Close();
                if (borrarArchivo)
                {
                    BorrarArchivo(filename);
                }
            }


        }

        public static void BorrarArchivo(String filename)
        {
            // Delete a file by using File class static method...
            if (System.IO.File.Exists(filename))
            {
                // Use a try block to catch IOExceptions, to
                // handle the case of the file already being
                // opened by another process.
                try
                {
                    System.IO.File.Delete(filename);
                }
                catch (System.IO.IOException e)
                {
                    Console.WriteLine(e.Message);
                    return;
                }
            }
        }

        public static void SplitFile(string inputFile, int chunkSize, string path)
        {
            byte[] buffer = new byte[chunkSize];
            using (Stream input = File.OpenRead(inputFile))
            {
                int index = 0;
                while (input.Position < input.Length)
                {
                    using (Stream output = File.Create(path + "\\" + index))
                    {
                        int chunkBytesRead = 0;
                        while (chunkBytesRead < chunkSize)
                        {
                            int bytesRead = input.Read(buffer,
                                                       chunkBytesRead,
                                                       chunkSize - chunkBytesRead);

                            if (bytesRead == 0)
                            {
                                break;
                            }
                            chunkBytesRead += bytesRead;
                        }
                        output.Write(buffer, 0, chunkBytesRead);
                    }
                    index++;
                }
            }
        }

        private static string HashString(string text)
        {
            const string chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            byte[] bytes = Encoding.UTF8.GetBytes(text);

            SHA256Managed hashstring = new SHA256Managed();
            byte[] hash = hashstring.ComputeHash(bytes);

            char[] hash2 = new char[32];

            // Note that here we are wasting bits of hash! 
            // But it isn't really important, because hash.Length == 32
            for (int i = 0; i < hash2.Length; i++)
            {
                hash2[i] = chars[hash[i] % chars.Length];
            }

            return new string(hash2);
        }

        public static void BorrarArchivos_MenosUltimo(string directorio, string prefijo, string extension_archivo, string ultimo)
        {
            FileInfo archivo;
            string[] listaArchivos = Directory.GetFiles(directorio, prefijo + "*." + extension_archivo);
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                archivo = new FileInfo(listaArchivos[i]);
                if(ultimo != null)
                {
                    if (!archivo.Name.Contains(ultimo))
                        archivo.Delete();
                }
                else
                    archivo.Delete();

            }
        }
    }
}
