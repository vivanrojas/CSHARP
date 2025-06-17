//using ICSharpCode.SharpZipLib.Core;
//using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class ZipUnZip
    {
        public void ComprimirArchivo(String ficheroOrigen, String ficheroZip)
        {

            try
            {
                string sourceName = ficheroOrigen;
                string targetName = ficheroZip;
                Console.WriteLine("Comprimiendo " + ficheroOrigen + " --> " + ficheroZip);
                ProcessStartInfo p = new ProcessStartInfo();
                p.FileName = AppDomain.CurrentDomain.BaseDirectory + @"bin\7z.exe";
                p.Arguments = "a -tzip \"" + targetName + "\" \"" + sourceName + "\" -y";
                p.WindowStyle = ProcessWindowStyle.Hidden;
                Process x = Process.Start(p);
                x.WaitForExit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
        }

        public void ComprimirVarios(string patron, string ficheroZip)
        {
            try
            {
                string sourceName = patron;
                string targetName = ficheroZip;
                Console.WriteLine("Comprimiendo  --> " + ficheroZip);
                ProcessStartInfo p = new ProcessStartInfo();
                p.FileName = AppDomain.CurrentDomain.BaseDirectory + @"bin\7z.exe";
                p.Arguments = "a -tzip \"" + targetName + "\" \"" + sourceName + "\" -y";
                p.WindowStyle = ProcessWindowStyle.Hidden;
                Process x = Process.Start(p);
                x.WaitForExit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
        }
        public void DescomprimirArchivo(String ficheroOrigen)
        {
            string sourceName = ficheroOrigen;
            string targetName = sourceName.Substring(0, sourceName.LastIndexOf("\\", sourceName.Length));
            ProcessStartInfo p = new ProcessStartInfo();
            p.FileName = AppDomain.CurrentDomain.BaseDirectory + @"bin\7z.exe";
            p.Arguments = "e \"" + sourceName + " \"" + " -o" + "\"" + targetName + "\" -y";
            p.WindowStyle = ProcessWindowStyle.Hidden;
            Console.WriteLine("Descomprimiendo " + ficheroOrigen);
            Process x = Process.Start(p);
            x.WaitForExit();
        }

        public void ComprimirArchivo_Split(String ficheroOrigen, String ficheroZip, int megas)
        {

            try
            {
                string sourceName = ficheroOrigen;
                string targetName = ficheroZip;
                Console.WriteLine("Comprimiendo " + ficheroOrigen + " --> " + ficheroZip);
                ProcessStartInfo p = new ProcessStartInfo();
                p.FileName = AppDomain.CurrentDomain.BaseDirectory + @"bin\7z.exe";
                p.Arguments = "a -tzip -v" + megas + "m \"" + targetName + "\" \"" + sourceName + "\" -y";
                p.WindowStyle = ProcessWindowStyle.Hidden;
                Process x = Process.Start(p);
                x.WaitForExit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
        }

        public void Descomprimir(string ficheroOrigen, string destino)
        {
            string sourceName = ficheroOrigen;
            string targetName = destino;
            ProcessStartInfo p = new ProcessStartInfo();
            p.FileName = AppDomain.CurrentDomain.BaseDirectory + @"bin\7z.exe";
            p.Arguments = "e \"" + sourceName + " \"" + " -o" + "\"" + targetName + "\" -y";
            p.WindowStyle = ProcessWindowStyle.Hidden;
            Console.WriteLine("Descomprimiendo " + ficheroOrigen);
            Process x = Process.Start(p);
            x.WaitForExit();
        }

        //public void DescomprimirNuevo(string ficheroOrigen, string destino, string password, bool utilizar_barra_progreso)
        //{
        //    ZipFile zf = null;

        //    double percent = 0;
        //    int i = 0;

        //    forms.FrmProgressBar pb = new forms.FrmProgressBar();



        //    try
        //    {
        //        FileStream fs = File.OpenRead(ficheroOrigen);
        //        zf = new ZipFile(fs);
        //        if (!String.IsNullOrEmpty(password))
        //        {
        //            zf.Password = password;     // AES encrypted entries are handled automatically
        //        }

        //        if (utilizar_barra_progreso)
        //        {
        //            pb.Text = "Descomprimiendo " + ficheroOrigen;
        //            pb.Show();
        //            pb.progressBar.Step = 1;
        //            pb.progressBar.Maximum = Convert.ToInt32(zf.Count);
        //        }



        //        foreach (ZipEntry zipEntry in zf)
        //        {

        //            if (!zipEntry.IsFile)
        //            {
        //                continue;           // Ignore directories
        //            }
        //            String entryFileName = zipEntry.Name;

        //            if (utilizar_barra_progreso)
        //            {
        //                i++;
        //                percent = (i / Convert.ToDouble(zf.Count)) * 100;
        //                pb.progressBar.Increment(1);
        //                pb.txtDescripcion.Text = "Extrayendo " + i.ToString("#.###") + " / " + zf.Count.ToString("#.###") + " archivos --> " + zipEntry.Name;
        //                pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
        //                pb.Refresh();
        //            }



        //            // to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
        //            // Optionally match entrynames against a selection list here to skip as desired.
        //            // The unpacked length is available in the zipEntry.Size property.

        //            byte[] buffer = new byte[4096];     // 4K is optimum
        //            Stream zipStream = zf.GetInputStream(zipEntry);

        //            // Manipulate the output filename here as desired.
        //            String fullZipToPath = Path.Combine(destino, entryFileName);
        //            string directoryName = Path.GetDirectoryName(fullZipToPath);
        //            if (directoryName.Length > 0)
        //                Directory.CreateDirectory(directoryName);

        //            // Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
        //            // of the file, but does not waste memory.
        //            // The "using" will close the stream even if an exception occurs.
        //            using (FileStream streamWriter = File.Create(fullZipToPath))
        //            {
        //                StreamUtils.Copy(zipStream, streamWriter, buffer);
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        if (zf != null)
        //        {
        //            zf.IsStreamOwner = true; // Makes close also shut the underlying stream
        //            zf.Close(); // Ensure we release resources

        //            if(utilizar_barra_progreso)
        //                pb.Close();
        //        }
        //    }
        //}
    }
}
