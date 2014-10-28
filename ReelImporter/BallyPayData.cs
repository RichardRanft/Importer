using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public abstract class BallyPayData
    {
        public abstract BallyPayType Type
        {
            get;
        }

        public abstract void Parse(StreamReader inStream, PayParserState parseState);

        protected abstract void exportPays(String sheetName, Excel.Workbook targetBook);
    }
}
