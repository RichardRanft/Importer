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
    public class BallyScatterPayData : BallyPayData
    {
        private List<PaylineDescription> m_scatterPays;
        private Utils m_util;
        private BallyPayType m_type;

        public override BallyPayType Type
        {
            get
            {
                return m_type;
            }
        }

        public BallyScatterPayData()
        {
            m_scatterPays = new List<PaylineDescription>();
            m_util = new Utils();
            m_type = BallyPayType.SCATTER_PAY;
        }

        public override void Parse(StreamReader inStream, String line, PayParserState parseState)
        {
            PaylineDescription payline;
            // ----
            payline = new PaylineDescription();
            //payline.Add(line, m_util);
            if (parseState.CurrentPayType == BallyPayType.SCATTER_PAY)
            {
                m_scatterPays.Add(payline);
                parseState.ResetScatterPay();
            }
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
