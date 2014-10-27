using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ReelImporter
{
    public class BallyLinePayData : BallyPayData
    {
        private Utils m_util;
        private List<PaylineDescription> m_stopSets;

        public BallyLinePayData()
        {
            m_util = new Utils();
            m_stopSets = new List<PaylineDescription>();
        }

        public void Parse(StreamReader inStream, PayParserState parseState)
        {
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

                    //if ((line.Length > 1) && lineHasOpenBrace)
                    //{
                    //    if (!parseState.ReelStart && !parseState.FreeStart && !parseState.ModifierStart)
                    //    {
                    //        if (parseState.ReelSetStart && !parseState.LINEPAYSTART)
                    //            parseState.EnterBaseReel();
                    //        if (parseState.FreeSetStart && !parseState.LINEPAYSTART)
                    //            parseState.EnterFreeReel();
                    //        if (parseState.LINEPAYSTART)
                    //            parseState.EnterModifierReel();
                    //    }
                    //}

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

                    tmpReel = new BallyReel();
                    tmpReel.Parse(inStream, line, parseState);
                    if (parseState.CurrentSetType == ReelSetType.BASEMODREEL)
                    {
                        m_baseModReelset.AddReel(tmpReel);
                        parseState.ResetModifierReel();
                    }

                    if (parseState.CurrentSetType == ReelSetType.BASEREEL)
                    {
                        m_baseReelset.AddReel(tmpReel);
                        parseState.ResetBaseReel();
                    }

                    if (parseState.CurrentSetType == ReelSetType.FREEMODREEL)
                    {
                        m_freeModReelset.AddReel(tmpReel);
                        parseState.ResetModifierReel();
                    }

                    if (parseState.CurrentSetType == ReelSetType.FREEREEL)
                    {
                        m_freeReelset.AddReel(tmpReel);
                        parseState.ResetFreeReel();
                    }
                }
            }
        }
    }
}
