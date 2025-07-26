using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Contexts;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace EndesaBusiness.sharePoint
{
    public class Utilidades
    {

        public string siteURL { get; set; }
        public string fileURL { get; set; }
        public string localFile { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
        public string destination { get; set; }

        logs.Log ficheroLog;

        public Utilidades()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "Sharepoint");
        }



        public void DownloadFile()
        {

        }

        public void Download()
        {

            Stream filestream = GetFile();
            string fileName = System.IO.Path.GetFileName(fileURL);
            string filePath = destination;

            FileStream fileStream = System.IO.File.Create(filePath, (int)filestream.Length);

            byte[] bytesInStream = new byte[filestream.Length];

            filestream.Read(bytesInStream, 0, bytesInStream.Length);

            fileStream.Write(bytesInStream, 0, bytesInStream.Length);

            fileStream.Close();

        }

        public void Upload_()
        {

            using (ClientContext clientContext = GetContextObject())
            {
                Web web = clientContext.Web;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                clientContext.Load(web, website => website.ServerRelativeUrl);
                clientContext.ExecuteQuery();
                Regex regex = new Regex(siteURL, RegexOptions.IgnoreCase);
                string strSiteRelativeURL = regex.Replace(fileURL, string.Empty);
                string strServerRelativeURL = CombineURL(web.ServerRelativeUrl, strSiteRelativeURL);

                clientContext.ExecuteQuery();
            }



        }

        public void Upload()
        {
            
        }

        private string CombineURL(string path1, string path2)
        {
            return path1.TrimEnd('/') + '/' + path2.TrimStart('/');
        }

        private Stream GetFile()
        {

            using (ClientContext clientContext = GetContextObject())
            {
                Web web = clientContext.Web;

                clientContext.Load(web, website => website.ServerRelativeUrl);

                clientContext.ExecuteQuery();

                Regex regex = new Regex(siteURL, RegexOptions.IgnoreCase);

                string strSiteRelativeURL = regex.Replace(fileURL, string.Empty);

                string strServerRelativeURL = CombineURL(web.ServerRelativeUrl, strSiteRelativeURL);


                Microsoft.SharePoint.Client.File oFile = web.GetFileByServerRelativeUrl(strServerRelativeURL);

                clientContext.Load(oFile);

                ClientResult<Stream> stream = oFile.OpenBinaryStream();

                clientContext.ExecuteQuery();

                return ReadFully(stream.Value);

            }



        }


        private Stream ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }

                return new MemoryStream(ms.ToArray());

            }
        }

        private ClientContext GetContextObject()
        {
            ClientContext context = new ClientContext(siteURL);
            context.Credentials = new SharePointOnlineCredentials(userName, GetPasswordFromString(password));
            return context;
        }

        private SecureString GetPasswordFromString(string password)
        {
            SecureString securePassword = new SecureString();
            char[] arrPassword = password.ToCharArray();
            foreach (char c in arrPassword)
                securePassword.AppendChar(c);

            return securePassword;
        }

        private void Prueba()

        { 
            //ClientContext context = new ClientContext(siteURL);
            //context.Credentials = new SharePointOnlineCredentials();
        }
    }
}
