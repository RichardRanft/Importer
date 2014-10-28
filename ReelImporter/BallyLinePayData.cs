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
    public class BallyLinePayData : BallyPayData
    {
        private List<PaylineDescription> m_linePays;
        private Utils m_util;
        private BallyPayType m_type;

        public override BallyPayType Type
        {
            get
            {
                return m_type;
            }
        }

        public BallyLinePayData()
        {
            m_util = new Utils();
            m_linePays = new List<PaylineDescription>();
            m_type = BallyPayType.LINEPAY;
        }

        public override void Parse(StreamReader inStream, String line, PayParserState parseState)
        {
            bool lineHasOpenBrace = false;
            bool lineHasCloseBrace = false;

            PaylineDescription payline;

            using (inStream)
            {
                while ((line = inStream.ReadLine()) != null)
                {
                    // strip comments
                    if (line.Contains("/"))
                    {
                        int pos = line.IndexOf("/");
                        line = line.Remove(pos);
                    }

                    line = line.Trim();

                    if (line.Length == 0 || line == "")
                        continue;

                    // check for braces
                    if (line == m_util.openBrace)
                    {
                        parseState.EnterArrayLevel();
                        continue;
                    }

                    lineHasOpenBrace = line.Contains(m_util.openBrace);
                    lineHasCloseBrace = line.Contains(m_util.closeBrace);

                    if (!lineHasOpenBrace && !lineHasCloseBrace)
                        continue;

                    if (lineHasCloseBrace)
                    {
                        if (line == m_util.closeBrace)
                        {
                            parseState.LeaveArrayLevel();
                            if (parseState.LinePayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.LINEPAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveLinePay();
                            }
                            else if (parseState.FreeLinePayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_LINEPAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeLinePay();
                            }
                            else if (parseState.ScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveScatterPay();
                            }
                            else if (parseState.FreeScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeScatterPay();
                            }

                            continue;
                        }

                        // could be end of a reelstop definition, or moving up a level
                        if (line == m_util.arrayEnd)
                        {
                            parseState.LeaveArrayLevel();
                            if (parseState.LinePayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.LINEPAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveLinePay();
                            }
                            else if (parseState.FreeLinePayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_LINEPAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeLinePay();
                            }
                            else if (parseState.ScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveScatterPay();
                            }
                            else if (parseState.FreeScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeScatterPay();
                            }

                            continue;
                        }
                    }

                    payline = new PaylineDescription();
                    payline.Add(line, m_util);
                    if (parseState.CurrentPayType == BallyPayType.LINEPAY)
                    {
                        m_linePays.Add(payline);
                        parseState.ResetLinePay();
                    }
                }
            }
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
