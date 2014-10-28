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
    public class BallyFreeLinePayData : BallyPayData
    {
        private List<PaylineDescription> m_freeLinePays;
        private Utils m_util;
        private BallyPayType m_type;

        public override BallyPayType Type
        {
            get
            {
                return m_type;
            }
        }

        public BallyFreeLinePayData()
        {
            m_freeLinePays = new List<PaylineDescription>();
            m_util = new Utils();
            m_type = BallyPayType.FREEGAME_LINEPAY;
        }

        public override void Parse(StreamReader inStream, String line, PayParserState parseState)
        {
            PaylineDescription payline;
            // ----
            payline = new PaylineDescription();
            //payline.Add(line, m_util);
            if (parseState.CurrentPayType == BallyPayType.FREEGAME_LINEPAY)
            {
                m_freeLinePays.Add(payline);
                parseState.ResetFreeLinePay();
            }
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
