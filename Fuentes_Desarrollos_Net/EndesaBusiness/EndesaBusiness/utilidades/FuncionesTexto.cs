using iTextSharp.xmp.impl;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class FuncionesTexto
    {


        public static string TextToHtml(string text)
        {
            text = HttpUtility.HtmlEncode(text);
            text = text.Replace("\r\n", "\r");
            text = text.Replace("\n", "\r");
            text = text.Replace("\r", "<br>\r\n");
            text = text.Replace("  ", " &nbsp;");
            return text;
        }

        public static string TextToHtml_href(string text)
        {
            text = HttpUtility.HtmlEncode(text);
            text = text.Replace("\r\n", "\r");
            text = text.Replace("\n", "\r");
            text = text.Replace("\r", "<br>\r\n");
            text = text.Replace("  ", " &nbsp;");
            text = text.Replace("&lt;", "<");
            text = text.Replace("&gt;", ">");
            
            return text;
        }

        public static string CN(String t)
        {
            t = t.Trim();
            if (t == "")
            {
                return "null";
            }
            else
            {
                t = t.Replace(" ", "");
                t = t.Replace("+", string.Empty);
                t = t.Replace("----------", string.Empty);
                t = t.Replace(".", string.Empty);
                t = t.Replace(",", ".");

                if (t == "")
                {
                    t = "null";
                }
            }

            return t;
        }

        public static string SinTildes(string texto)
        {

            Regex replace_a_Accents = new Regex("[á|à|ä|â]", RegexOptions.Compiled);
            Regex replace_e_Accents = new Regex("[é|è|ë|ê]", RegexOptions.Compiled);
            Regex replace_i_Accents = new Regex("[í|ì|ï|î]", RegexOptions.Compiled);
            Regex replace_o_Accents = new Regex("[ó|ò|ö|ô]", RegexOptions.Compiled);
            Regex replace_u_Accents = new Regex("[ú|ù|ü|û]", RegexOptions.Compiled);
            texto = replace_a_Accents.Replace(texto, "a");
            texto = replace_e_Accents.Replace(texto, "e");
            texto = replace_i_Accents.Replace(texto, "i");
            texto = replace_o_Accents.Replace(texto, "o");
            texto = replace_u_Accents.Replace(texto, "u");
            return texto;
        }

        

        public static string CS(String t)
        {
            String salida = "";

            if (t.Trim() == "")
            {
                salida = "null";
            }
            else
            {
                salida = "'" + t.Trim() + "'";
            }

            return salida;
        }

        public static string CD(String t)
        {
            String salida = "";

            if (t.Trim() == "00000000" || t.Trim() == "")
            {
                salida = "null";
            }
            else
            {
                salida = "'" + t.Substring(0, 4) + "-" + t.Substring(4, 2) + "-" + t.Substring(6, 2) + "'";
            }

            return salida;
        }

        public static string ArreglaAcentos(String t)
        {
            String salida;

            salida = t.Replace("\"", string.Empty);
            salida = salida.Replace("\\", string.Empty);
            salida = salida.Replace("\t", string.Empty);
            salida = salida.Replace("'", "´");
            salida = salida.Replace("Âª", "");
            salida = salida.Replace("_", "-");
            salida = salida.Trim();
            return salida;
        }

        public static string RT(String t)
        {
            String salida;

            salida = t.Replace("\"", string.Empty);
            salida = salida.Replace("\\", string.Empty);
            salida = salida.Replace("\t", string.Empty);
            salida = salida.Replace("'", "´");
            salida = salida.Replace("Âª", "");
            salida = salida.Trim();
            return salida;
        }

        public static string Decrypt(string cipherString, bool useHashing)
        {
            byte[] keyArray;
            //get the byte code of the string

            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            ////System.Configuration.AppSettingsReader settingsReader =
            ////                                    new AppSettingsReader();
            //Get your key from config file to open the lock!
            ////string key = (string)settingsReader.GetValue("SecurityKey",
            ////                                             typeof(String));

            string key = "operaciones1234";

            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(
                                 toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }

        public static string Encrypt(string toEncrypt, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            ////System.Configuration.AppSettingsReader settingsReader =
            ////                                    new AppSettingsReader();
            ////// Get the key from config file

            ////string key = (string)settingsReader.GetValue("SecurityKey",
            ////                                                 typeof(String));

            string key = "operaciones1234";
            //System.Windows.Forms.MessageBox.Show(key);
            //If hashing use get hashcode regards to your key
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //Always release the resources and flush data
                // of the Cryptographic service provide. Best Practice

                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes.
            //We choose ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)

            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            //transform the specified region of bytes array to resultArray
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor
            tdes.Clear();
            //Return the encrypted data into unreadable string format
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
    }
}
