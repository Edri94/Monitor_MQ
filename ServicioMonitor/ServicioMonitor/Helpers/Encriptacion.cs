using System;
using System.Security.Cryptography;
using System.Text;


namespace ServicioMonitor.Helpers
{
    public class Encriptacion
    {
        private static TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
        private static MD5CryptoServiceProvider MD5 = new MD5CryptoServiceProvider();

        public static byte[] MD5Hash(string value)
        {
            return MD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(value));
        }

        public string Encrypt(string stringToEncrypt, string key = "BBVA BANCOMER")
        {
            // TODO: On Error Resume Next Warning!!!: The statement is not translatable
            // Warning!!! Optional parameters not supported
            if ((stringToEncrypt == ""))
            {
                stringToEncrypt = "BBVA BANCOMER";
            }

            DES.Key = MD5Hash(key);
            DES.Mode = CipherMode.ECB;
            byte[] Buffer = ASCIIEncoding.ASCII.GetBytes(stringToEncrypt);
            return Convert.ToBase64String(DES.CreateEncryptor().TransformFinalBlock(Buffer, 0, Buffer.Length));
        }

        public string Decrypt(string encryptedString, string key = "BBVA BANCOMER")
        {
            try
            {
                if ((encryptedString == ""))
                {
                    return "";
                    // Warning!!! Optional parameters not supported
                    // TODO: Exit Function: Warning!!! Need to return the value                    
                }

                DES.Key = MD5Hash(key);
                DES.Mode = CipherMode.ECB;
                byte[] Buffer = Convert.FromBase64String(encryptedString);
                return ASCIIEncoding.ASCII.GetString(DES.CreateDecryptor().TransformFinalBlock(Buffer, 0, Buffer.Length));
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
