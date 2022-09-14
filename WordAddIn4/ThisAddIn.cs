using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;
using System.Security.Cryptography;
using System.IO;
using System.Windows.Forms;


namespace WordAddIn4
{
    public partial class ThisAddIn
    {

        // Initialize value for AES protocol
        byte[] Iv = { 54, 23, 72, 254, 1, 36, 193, 153, 27, 83, 13, 154, 64, 3, 201, 85 };
        byte[] Key = { 12, 64, 42, 46, 235, 222, 125, 5, 34, 164, 42, 2, 200, 64, 13, 96, 65,
            153, 176, 2, 63, 7, 24, 199, 59, 14, 106, 34, 132, 55, 222, 70 };



        static void Check_plainText(string plainText)
        {
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
        }
        static void Check_cipherText(byte[] cipherText)
        {
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
        }


        static String AES_Encryption(string plainText, byte[] Key, byte[] IV)
        {
            byte[] encrypted;
            String Base64String;
            // Check arguments.
            Check_plainText(plainText);
            // Create an Aes object
            // with the specified key and IV.
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = Key;
                aesAlg.IV = IV;

                // Create an encryptor to perform the stream transform.
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                // Create the streams used for encryption.
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                        Base64String = Convert.ToBase64String(encrypted);
                    }
                }
            }
            // Return the encrypted string in base64 type
            return Base64String;
        }




        static string AES_Decryption(byte[] cipherText, byte[] Key, byte[] IV)
        {
            // Check arguments.
            Check_cipherText(cipherText);

            // Declare the string used to hold
            // the decrypted text.
            string plaintext = null;

            // Create an Aes object
            // with the specified key and IV.
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = Key;
                aesAlg.IV = IV;
                // Create a decryptor to perform the stream transform.
                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);

                // Create the streams used for decryption.
                using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                        {
                            // Read the decrypted bytes from the decrypting stream
                            // and place them in a string.
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }
            return plaintext;
        }

        //Decrypt function for ActiveDocument content text. (from Base64context to readable text using AES)
        //1- Get Active document
        //2- Get Base64string from Content.Text
        //3- Convert Base64String to bytearray
        //4- Start decryption by AES
        //5- Save document
        public void Decrypt(Word.Document Doc)
        {       
            MessageBox.Show("Press Ok to decrypt context...");
            String context;
            byte[] context_bytes;
            String Base64context = Doc.Content.Text;
            context_bytes = Convert.FromBase64String(Base64context);
            context = AES_Decryption(context_bytes, Key, Iv);
            Doc.Content.Text = context;
            Doc.Save();

        }

        //Encrypt function for ActiveDocument content text. (from readabletext to Base64context ,using AES)
        //1- Get Active document
        //2- Get readabletext from Active document
        //3- Start encryption by AES
        //5- Save document



        private void Encrypt(Word.Document Doc, ref bool cancel )
        {
            String Base64context;
            Doc = Application.ActiveDocument;
            String context = Doc.Content.Text;
            Base64context = AES_Encryption(context, Key, Iv);
            Doc.Content.Text = Base64context;
            Doc.Save();
        }



        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Encrypt document before closing
            this.Application.DocumentBeforeClose +=
            new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Encrypt);

        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
