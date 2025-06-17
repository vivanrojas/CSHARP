using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using FluentFTP;
using FluentFTP.GnuTLS;
using FluentFTP.GnuTLS.Enums;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Runtime.InteropServices.WindowsRuntime;
//using static Org.BouncyCastle.Math.EC.ECCurve;

namespace EndesaBusiness.utilidades
{
    public class UltimateFTP
    {
        private string host = null;
        private string user = null;
        private string pass = null;
        private int port;

        public UltimateFTP(string hostIP, string userName, string password, string port_number)
        {
            Console.WriteLine("Conecting to " + hostIP + " user: " + userName);
            host = hostIP; user = userName; pass = password; port = Convert.ToInt32(port_number);

        }

        public void Download(string remoteFileName, string localDestinationFileName)
        {
            String Host = host;
            int Port = port;

            using (var sftp = new SftpClient(host, port, user, pass))
            {
                sftp.Connect();

                using (var file = File.OpenWrite(localDestinationFileName))
                {
                    sftp.DownloadFile(remoteFileName, file);
                }

                sftp.Disconnect();
            }
        }

        public void DownloadAll(string remoteDirectory, string localDirectory, int desde_num_dias)
        {
            String Host = host;
            int Port = port;

            using (var sftp = new SftpClient(host, port, user, pass))
            {
                sftp.Connect();
                var files = sftp.ListDirectory(remoteDirectory);
                foreach (var file in files)
                {
                    string remoteFileName = file.Name;
                    if (file.LastWriteTime.Date > DateTime.Today.AddDays(desde_num_dias))
                        Download(remoteDirectory + remoteFileName, localDirectory + remoteFileName);


                }


                sftp.Disconnect();
            }
        }
        public void DownloadAll(string remoteDirectory, string localDirectory, int desde_num_dias, string prefijo)
        {
            String Host = host;
            int Port = port;

            using (var sftp = new SftpClient(host, port, user, pass))
            {
                sftp.Connect();
                var files = sftp.ListDirectory(remoteDirectory);
                foreach (var file in files)
                {
                    string remoteFileName = file.Name;
                    if (file.LastWriteTime.Date > DateTime.Today.AddDays(desde_num_dias) && file.Name.Substring(0, prefijo.Length) == prefijo)
                    {
                        Console.WriteLine("Downloading " + file.FullName + " Creation Date: "
                            + file.LastWriteTime.Date.ToString("yyyyMMdd")
                            + " Size: " + (file.Length / 1024) + " KB");
                        Download(remoteDirectory + remoteFileName, localDirectory + remoteFileName);
                    }



                }


                sftp.Disconnect();
            }
        }


        public List<string> Lista_Contenido_Directorio(string directorio)
        {
                     
            List<string> lista = new List<string>();

            String Host = host;
            int Port = port;

            using (var sftp = new SftpClient(host, port, user, pass))
            {
                sftp.Connect();
                var files = sftp.ListDirectory(directorio);
                foreach (var file in files)
                {
                    string remoteFileName = file.Name;
                    lista.Add(file.Name);
                }


                sftp.Disconnect();
            }

                return lista;
        }

        public void Upload(string remoteFileName, string localDestinationFileName)
        {
            String Host = host;
            int Port = port;

            using (var sftp = new SftpClient(host, port, user, pass))
            {
                sftp.Connect();

                using (var file = File.OpenRead(localDestinationFileName))
                {
                    Console.WriteLine("Uploading " + localDestinationFileName
                        + " to " + remoteFileName);
                    sftp.UploadFile(file, remoteFileName);
                }

                sftp.Disconnect();
            }
        }

        public List<string> Lista_Contenido_Directorio_InsecureFTP(string localPath)
        {
            string url = host;

            NetworkCredential credentials;
            try
            {
                credentials = new NetworkCredential(user, pass);

                FtpWebRequest listRequest = (FtpWebRequest)WebRequest.Create(url);
                listRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                listRequest.Credentials = credentials;

                List<string> lines = new List<string>();

                using (var listResponse = (FtpWebResponse)listRequest.GetResponse())
                using (Stream listStream = listResponse.GetResponseStream())
                using (var listReader = new StreamReader(listStream))
                {
                    while (!listReader.EndOfStream)
                    {
                        lines.Add(listReader.ReadLine());
                    }
                }
                return lines;                
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void DownloadAllInSecureFTP(string localPath)
        {
            string url = host;
            
            NetworkCredential credentials;
            try
            {
                credentials = new NetworkCredential(user, pass);

                FtpWebRequest listRequest = (FtpWebRequest)WebRequest.Create(url);
                listRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                listRequest.Credentials = credentials;

                List<string> lines = new List<string>();

                using (var listResponse = (FtpWebResponse)listRequest.GetResponse())
                using (Stream listStream = listResponse.GetResponseStream())
                using (var listReader = new StreamReader(listStream))
                {
                    while (!listReader.EndOfStream)
                    {
                        lines.Add(listReader.ReadLine());
                    }
                }

                foreach (string line in lines)
                {
                    string[] tokens =
                        line.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                    string name = tokens[8];
                    string permissions = tokens[0];

                    string localFilePath = Path.Combine(localPath, name);
                    string fileUrl = url + name;

                    if (permissions[0] == 'd')
                    {
                        if (!Directory.Exists(localFilePath))
                        {
                            Directory.CreateDirectory(localFilePath);
                        }

                        //DownloadFtpDirectory(fileUrl + "/", credentials, localFilePath);
                    }
                    else
                    {
                        FtpWebRequest downloadRequest =
                            (FtpWebRequest)WebRequest.Create(fileUrl);
                        downloadRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                        downloadRequest.Credentials = credentials;

                        using (FtpWebResponse downloadResponse =
                                  (FtpWebResponse)downloadRequest.GetResponse())
                        using (Stream sourceStream = downloadResponse.GetResponseStream())
                        using (Stream targetStream = File.Create(localFilePath))
                        {
                            byte[] buffer = new byte[10240];
                            int read;
                            while ((read = sourceStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                targetStream.Write(buffer, 0, read);
                            }
                        }
                    }


                }
            }
            catch (Exception ex)
            {

            }
        }

        public void UploadInSecureFTP(string remoteFileName, string localDestinationFileName)
        {
            String Host = host;
            int Port = port;



            using (WebClient client = new WebClient())
            {
                client.Credentials = new NetworkCredential(user, pass);
                client.UploadFile(remoteFileName, WebRequestMethods.Ftp.UploadFile, localDestinationFileName);
            }
        }
        public bool UploadFTPS(string remoteFileName, string localDestinationFileName)
        {
            //String Host = host;
            //int Port = port;
            bool exito_subida = false;


            //using (WebClient client = new WebClient())
            //{
            //    client.Credentials = new NetworkCredential(user, pass);
            //    client.UploadFile(remoteFileName, WebRequestMethods.Ftp.UploadFile, localDestinationFileName);
            //}

            try
            {
                Uri extremadura_server_uri = new Uri(host);

                FtpClient conn = new FtpClient(extremadura_server_uri.Host, user, pass);
                conn.Config.EncryptionMode = FtpEncryptionMode.Explicit;
                conn.Config.ValidateAnyCertificate = true;
            
                //conn.Config.DataConnectionType = FtpDataConnectionType.AutoPassive;
                //conn.Config.DataConnectionEncryption = true;
                //conn.Config.SslProtocols = System.Security.Authentication.SslProtocols.Tls12;
                //conn.ValidateCertificate += new FtpSslValidation(delegate (FluentFTP.Client.BaseClient.BaseFtpClient control, FtpSslValidationEventArgs e) { e.Accept = true; });
                conn.Config.CustomStream = typeof(GnuTlsStream);
                conn.Config.CustomStreamConfig = new GnuConfig()
                    {
                        // optional configuration
                        SecuritySuite = GnuSuite.Secure128,
                        HandshakeTimeout = 50000
                        
                    };

                conn.Connect();

                FtpStatus resultado = conn.UploadFile(localDestinationFileName, remoteFileName);
                Console.WriteLine("Resultado de subir el fichero al FTPS: " + resultado.ToString());
                if (resultado == FtpStatus.Success)
                    exito_subida = true;


                conn.Disconnect();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ha ocurrido un error al intentar subir el fichero al FTPS: " + ex.Message);
                if (ex.InnerException != null)
                    Console.WriteLine("Inner exception: {0}", ex.InnerException);
            }

            return exito_subida;
        }

        public void DownloadInSecureFTP()
        {
            String Host = host;
            int Port = port;
            using (WebClient client = new WebClient())
            {
                //client.Credentials = new NetworkCredential(user, pass);
                //client.DownloadFile()
            }
        }

        public bool isValidConnection()
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(host);
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(user, pass);
                ////Pruebas GUS
                //request.EnableSsl = true;
                //ServicePointManager.ServerCertificateValidationCallback +=
                //       (s, certificate, chain, sslPolicyErrors) => true;
                //ServicePointManager.Expect100Continue = false;


                request.GetResponse();
            }
            catch (WebException ex)
            {
                return false;
            }
            return true;
        }
        public bool isValidConnectionFTPS()
        {
            try
            {
                Uri extremadura_server_uri = new Uri(host);

                FtpClient conn = new FtpClient(extremadura_server_uri.Host, user, pass);
                conn.Config.EncryptionMode = FtpEncryptionMode.Explicit;
                conn.Config.ValidateAnyCertificate = true;
                conn.Connect();

                conn.Disconnect();

            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        public List<string> ListaDirectorios(string directorio)
        {
            return GetFtpDirectoryContents(new Uri(directorio), new NetworkCredential(user, pass));
        }

        private List<string> GetFtpDirectoryContents(Uri requestUri, NetworkCredential networkCredential)
        {
            var directoryContents = new List<string>(); //Create empty list to fill it later.
                                                        //Create ftpWebRequest object with given options to get the Directory Contents. 
            var ftpWebRequest = GetFtpWebRequest(requestUri, networkCredential, WebRequestMethods.Ftp.ListDirectory);
            try
            {
                using (var ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse()) //Excute the ftpWebRequest and Get It's Response.
                using (var streamReader = new StreamReader(ftpWebResponse.GetResponseStream())) //Get list of the Directory Contentss as Stream.
                {
                    var line = string.Empty; //Initial default value for line.
                    do
                    {
                        line = streamReader.ReadLine(); //Read current line of Stream.
                        directoryContents.Add(line); //Add current line to Directory Contentss List.
                    } while (!string.IsNullOrEmpty(line)); //Keep reading while the line has value.
                }
            }
            catch (Exception) { } //Do nothing incase of Exception occurred.
            return directoryContents; //Return all list of Directory Contentss: Files/Sub Directories.
        }

        private FtpWebRequest GetFtpWebRequest(Uri requestUri, NetworkCredential networkCredential, string method = null)
        {
            var ftpWebRequest = (FtpWebRequest)WebRequest.Create(requestUri); //Create FtpWebRequest with given Request Uri.
            ftpWebRequest.Credentials = networkCredential; //Set the Credentials of current FtpWebRequest.

            if (!string.IsNullOrEmpty(method))
                ftpWebRequest.Method = method; //Set the Method of FtpWebRequest incase it has a value.
            return ftpWebRequest; //Return the configured FtpWebRequest.
        }
    }
}
