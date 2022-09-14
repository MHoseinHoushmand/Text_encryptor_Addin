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

        // Decript document when addin and ribbon is loaded
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
             Word.Document Doc  = Globals.ThisAddIn.Application.ActiveDocument;
             Globals.ThisAddIn.Decrypt(Doc);

        }

    }
}
