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
                return m_freeScatterPays;
            }
        }

        public BallyFreeScatterPayData()
        {
            m_freeScatterPays = new List<PaylineDescription>();
            m_util = new Utils();
            m_type = BallyPayType.FREEGAME_SCATTER_PAY;
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
                            if (parseState.FreeScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeScatterPay();
                            }

                            break;
                        }

                        // could be end of a reelstop definition, or moving up a level
                        if (line == m_util.arrayEnd)
                        {
                            parseState.LeaveArrayLevel();
                            if (parseState.FreeScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveFreeScatterPay();
                            }

                            break;
                        }
                    }

                    payline = new PaylineDescription();
                    payline.Add(line, m_util);
                    if (parseState.FreeScatterPayStart)
                    {
                        m_freeScatterPays.Add(payline);
                    }
                    else
                    {
                        parseState.ResetFreeScatterPay();
                        break;
                    }
                }
            }
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
