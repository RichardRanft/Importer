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
    public class BallyFreeScatterPayData : BallyPayData
    {
        private List<PaylineDescription> m_freeScatterPays;
        private Utils m_util;
        private BallyPayType m_type;

        public override BallyPayType Type
        {
            get
            {
                return m_type;
            }
        }

        public BallyFreeScatterPayData()
        {
            m_freeScatterPays = new List<PaylineDescription>();
        }

        public override void Parse(StreamReader inStream, PayParserState parseState)
        {
            PaylineDescription payline;
            // ----
            payline = new PaylineDescription();
            payline.Add(line, m_util);
            if (parseState.CurrentPayType == BallyPayType.FREEGAME_SCATTER_PAY)
            {
                m_freeScatterPays.Add(payline);
                parseState.ResetFreeScatterPay();
            }
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
