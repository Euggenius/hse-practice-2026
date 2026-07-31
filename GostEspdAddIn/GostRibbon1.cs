using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GostEspdAddIn
{
    public partial class GostRibbon1
    {
        private void GostRibbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var service = new GostEspdAddIn.Services.WordInteropService();
            service.FormatSelectedList(wordApp);
        }
    }
}
