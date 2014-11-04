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

        public override BallyPayType PayType
        {
            get
            {
                return m_type;
            }
        }

        public List<PaylineDescription> LinePays
        {
            get
            {
                return m_freeLinePays;
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
                            if (parseState.FreeLinePayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_LINEPAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeLinePay();
                            }
                            break;
                        }

                        // could be end of a reelstop definition, or moving up a level
                        if (line == m_util.arrayEnd)
                        {
                            parseState.LeaveArrayLevel();
                            if (parseState.FreeLinePayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_LINEPAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeLinePay();
                            }
                            break;
                        }
                    }

                    payline = new PaylineDescription();
                    payline.Add(line, m_util);
                    if (parseState.FreeLinePayStart)
                    {
                        m_freeLinePays.Add(payline);
                    }
                    else
                    {
                        parseState.ResetFreeLinePay();
                        break;
                    }
                }
            }

            IComparer<PaylineDescription> payComparer = new PaylineSorter();
            m_freeLinePays.Sort(payComparer);
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
