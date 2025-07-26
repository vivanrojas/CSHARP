using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;

namespace EndesaBusiness.utilidades
{
    public class ZIP
    {
        public ZIP()
        {

        }

        public void Descomprimir(string ficheroOrigen, string destino, string password)
        {
            //ZipFile zf = null;

            //double percent = 0;
            //int i = 0;
            ////forms.FrmProgressBar pb = new forms.FrmProgressBar();

            //try
            //{
            //    FileStream fs = File.OpenRead(ficheroOrigen);
            //    zf = new ZipFile(fs);
            //    if (!String.IsNullOrEmpty(password))
            //    {
            //        zf.Password = password;     // AES encrypted entries are handled automatically
            //    }

            //    //pb.Text = "Descomprimiendo " + ficheroOrigen;
            //    //pb.Show();
            //    //pb.progressBar.Step = 1;
            //    //pb.progressBar.Maximum = Convert.ToInt32(zf.Count);


            //    foreach (ZipEntry zipEntry in zf)
            //    {

            //        if (!zipEntry.IsFile)
            //        {
            //            continue;           // Ignore directories
            //        }
            //        String entryFileName = zipEntry.Name;

            //        i++;
            //        percent = (i / Convert.ToDouble(zf.Count)) * 100;
            //        //pb.progressBar.Increment(1);
            //        //pb.txtDescripcion.Text = "Extrayendo " + i.ToString("#.###") + " / " + zf.Count.ToString("#.###") + " archivos --> " + zipEntry.Name;
            //        //pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
            //        //pb.Refresh();


            //        // to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
            //        // Optionally match entrynames against a selection list here to skip as desired.
            //        // The unpacked length is available in the zipEntry.Size property.

            //        byte[] buffer = new byte[4096];     // 4K is optimum
            //        Stream zipStream = zf.GetInputStream(zipEntry);

            //        // Manipulate the output filename here as desired.
            //        String fullZipToPath = Path.Combine(destino, entryFileName);
            //        string directoryName = Path.GetDirectoryName(fullZipToPath);
            //        if (directoryName.Length > 0)
            //            Directory.CreateDirectory(directoryName);

            //        // Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
            //        // of the file, but does not waste memory.
            //        // The "using" will close the stream even if an exception occurs.
            //        using (FileStream streamWriter = File.Create(fullZipToPath))
            //        {
            //            StreamUtils.Copy(zipStream, streamWriter, buffer);
            //        }
            //    }
            //}
            //finally
            //{
            //    if (zf != null)
            //    {
            //        zf.IsStreamOwner = true; // Makes close also shut the underlying stream
            //        zf.Close(); // Ensure we release resources
            //        //pb.Close();
            //    }
            //}
        }

        public void Comprmir(string ficheroOrigen, string destino)
        {
            FileInfo fichero = new FileInfo(ficheroOrigen);

            //FastZip fz = new FastZip();
            //fz.CreateZip(destino, fichero.DirectoryName, false, fichero.Name);


        }

        public void ComprimirVarios(string directorio_origen, string patron, string destino)
        {
            //FastZip fz = new FastZip();
            // fz.CreateZip(zipFile, prjDir, true, ".*\\.(xls|doc|xml)$", "");
            //fz.CreateZip(destino, directorio_origen, false, patron,"");
        }

        public void DescomprimirArchivoZip(string rutaArchivoZip, string rutaDestino)
        {
            
            if (!Directory.Exists(rutaDestino))
            {
                Directory.CreateDirectory(rutaDestino);
            }

            
            ZipFile.ExtractToDirectory(rutaArchivoZip, rutaDestino);
        }
    }
}
