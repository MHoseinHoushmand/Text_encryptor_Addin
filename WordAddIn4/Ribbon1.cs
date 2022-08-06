using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn4
{
    public partial class Ribbon1
    {

        public bool IsBase64Encoded(String str)
        {
            try
            {
                // If no exception is caught, then it is possibly a base64 encoded string
                byte[] data = Convert.FromBase64String(str);
                // The part that checks if the string was properly padded to the
                // correct length was borrowed from d@anish's solution
                return (str.Replace(" ", "").Length % 4 == 0);
            }
            catch
            {
                // If exception is caught, then it is not a base64 encoded string
                return false;
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
             Word.Document Doc  = Globals.ThisAddIn.Application.ActiveDocument;
             Globals.ThisAddIn.Decrypt(Doc);
             this.button1.Enabled = false;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
