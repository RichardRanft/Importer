using System;
using System.Collections;
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
    public class PaylineSorter : IComparer<PaylineDescription>
    {
        public int Compare(PaylineDescription first, PaylineDescription second)
        {
            if (first == null)
                return -1;
            if (second == null)
                return 1;

            int value = 0; // a and b are equal
            int freeSet = 0;
            int hasWild = 0;
            int alphaRank = 0;
            int noHitRank = 0;
            int winRank = 0;

            //// first, reel type
            //if (first.IsFreegameSet)
            //    freeSet = -1;
            //if (second.IsFreegameSet)
            //    freeSet = 1;

            // next, wild sets
            int firstIndex = first.StopValues.Count - 1;
            int secondIndex = second.StopValues.Count - 1;
            //if (first.StopValues[firstIndex].HasWild())
            //    hasWild = -1;
            //if (second.StopValues[secondIndex].HasWild())
            //    hasWild = 1;

            // next, alphabetical
            String currFirst = "";
            String currSecond = "";
            ReelDescription firstReels = first.StopValues[firstIndex];
            ReelDescription secondReels = second.StopValues[secondIndex];
            currFirst = firstReels.ToString();
            currSecond = secondReels.ToString(); ;
            alphaRank = String.Compare(currFirst, currSecond);

            // next pay value
            //if (first.Win > second.Win)
            //    winRank = 1;
            //if (first.Win < second.Win)
            //    winRank = -1;

            // next, count no hit ("XX" or "-") entries
            //int noHitA = 0;
            //int noHitB = 0;

            //foreach( String entry in first.StopValues[firstIndex].Values )
            //{
            //    if (entry.Contains("XX") || entry.Contains("-"))
            //        noHitA++;
            //}

            //foreach (String entry in second.StopValues[secondIndex].Values)
            //{
            //    if (entry.Contains("XX") || entry.Contains("-"))
            //        noHitB++;
            //}

            //noHitRank = noHitA - noHitB;

            value = freeSet + hasWild + alphaRank + noHitRank + winRank;

            // x < 0 < y
            return value;
        }
    }

    public class BallyLinePayData : BallyPayData
    {
        private List<PaylineDescription> m_linePays;
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
                return m_linePays;
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
                            break;
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
                            break;
                        }
                    }

                    payline = new PaylineDescription();
                    payline.Add(line, m_util);
                    if (parseState.LinePayStart)
                    {
                        m_linePays.Add(payline);
                    }
                    else
                    {
                        parseState.ResetLinePay();
                        break;
                    }
                }
            }

            IComparer<PaylineDescription> payComparer = new PaylineSorter();
            m_linePays.Sort(payComparer);
        }

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
