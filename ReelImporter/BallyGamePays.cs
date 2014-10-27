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
    public class BallyGamePays
    {
        private PayParserState m_parseState;
        private Utils m_util;

        public BallyGamePays()
        {
            m_parseState = new PayParserState();
            m_util = new Utils();
        }

        private List<BallyPayData> getSubSets(BallyPayData set)
        {
            List<BallyPayData> temp = new List<BallyPayData>();

            return temp;
        }

        public void Parse(StreamReader inStream)
        {
            if (m_parseState == null)
                m_parseState = new PayParserState();
            bool lineHasOpenBrace = false;
            bool lineHasCloseBrace = false;

            
            String line = "";

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

                    // look for reels
                    if (line == "linepays")
                    {
                        m_parseState.EnterLinePay();
                        continue;
                    }
                    if (line == "freegame_linepays")
                    {
                        m_parseState.EnterFreeLinePay();
                        continue;
                    }
                    if (line == "scatterpays")
                    {
                        m_parseState.EnterScatterPay();
                        continue;
                    }
                    if (line == "freegame_scatterpays")
                    {
                        m_parseState.EnterFreeScatterPay();
                        continue;
                    }
                    if (line == "featuredefs")
                    {
                        break;
                    }

                    if (m_parseState.LinePayStart || m_parseState.FreeLinePayStart || m_parseState.ScatterPayStart || m_parseState.FreeScatterPayStart)
                    {
                        // check for braces
                        if (line == m_util.openBrace)
                        {
                            m_parseState.EnterArrayLevel();
                            continue;
                        }

                        lineHasOpenBrace = line.Contains(m_util.openBrace);
                        lineHasCloseBrace = line.Contains(m_util.closeBrace);

                        if (!lineHasOpenBrace && !lineHasCloseBrace)
                            continue;

                        //if ((line.Length > 1) && lineHasOpenBrace)
                        //{
                        //    if (!m_parseState.ReelStart && !m_parseState.FreeStart && !m_parseState.ModifierStart)
                        //    {
                        //        if (m_parseState.ReelSetStart && !m_parseState.LINEPAYSTART)
                        //            m_parseState.EnterBaseReel();
                        //        if (m_parseState.FreeSetStart && !m_parseState.LINEPAYSTART)
                        //            m_parseState.EnterFreeReel();
                        //        if (m_parseState.LINEPAYSTART)
                        //            m_parseState.EnterModifierReel();
                        //    }
                        //}

                        if (lineHasCloseBrace)
                        {
                            if (line == m_util.closeBrace)
                            {
                                m_parseState.LeaveArrayLevel();
                                if (m_parseState.LinePayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.LINEPAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveLinePay();
                                }
                                else if (m_parseState.FreeLinePayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_LINEPAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveFreeLinePay();
                                }
                                else if (m_parseState.ScatterPayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveScatterPay();
                                }
                                else if (m_parseState.FreeScatterPayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveFreeScatterPay();
                                }

                                continue;
                            }

                            // could be end of a reelstop definition, or moving up a level
                            if (line == m_util.arrayEnd)
                            {
                                m_parseState.LeaveArrayLevel();
                                if (m_parseState.LinePayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.LINEPAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveLinePay();
                                }
                                else if (m_parseState.FreeLinePayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_LINEPAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveFreeLinePay();
                                }
                                else if (m_parseState.ScatterPayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveScatterPay();
                                }
                                else if (m_parseState.FreeScatterPayStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveFreeScatterPay();
                                }

                                continue;
                            }
                        }

                        tmpReel = new BallyReel();
                        tmpReel.Parse(inStream, line, m_parseState);
                        if (m_parseState.CurrentSetType == ReelSetType.BASEMODREEL)
                        {
                            m_baseModReelset.AddReel(tmpReel);
                            m_parseState.ResetModifierReel();
                        }

                        if (m_parseState.CurrentSetType == ReelSetType.BASEREEL)
                        {
                            m_baseReelset.AddReel(tmpReel);
                            m_parseState.ResetBaseReel();
                        }

                        if (m_parseState.CurrentSetType == ReelSetType.FREEMODREEL)
                        {
                            m_freeModReelset.AddReel(tmpReel);
                            m_parseState.ResetModifierReel();
                        }

                        if (m_parseState.CurrentSetType == ReelSetType.FREEREEL)
                        {
                            m_freeReelset.AddReel(tmpReel);
                            m_parseState.ResetFreeReel();
                        }
                    }
                }
            }
            m_isValid = checkValid();
            m_reelWidth = m_baseReelset.Count;
        }

        public void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
