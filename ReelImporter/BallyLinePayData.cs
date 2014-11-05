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
        private int m_rowCount;

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
            m_rowCount = 0;
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
                                if (parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == parseState.ArrayDepth)
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
                                if (parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == parseState.ArrayDepth)
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

        public void SendToWorksheet(Excel.Workbook targetBook, Excel.Worksheet targetSheet)
        {
            Globals.Program.Application.ScreenUpdating = false;
            targetSheet.Activate();
            // reel entries A5~E5, pay at M5
            String col = "A";
            int row = 5;
            String cell = col + row.ToString();
            String payCell = "N" + row.ToString();
            int stopSet = m_linePays[0].StopValues.Count - 1;
            bool trim = true;
            String val = "";
            foreach (PaylineDescription line in m_linePays)
            {
                if (line.Win <= 0)
                    continue;
                if (line.IsFreegameSet || line.IsModifierSet || line.HasWild)
                    //continue;
                    if (line.IsFreegameSet || line.IsModifierSet)
                        trim = true;

                col = "A";
                foreach (String stop in line.StopValues[stopSet].Values)
                {
                    cell = col + row.ToString();
                    if (line.StopValues[stopSet].Values.Count > 5)
                        break;
                    if (trim)
                        val = line.StopValues[stopSet].TrimValue(stop);
                    else
                        val = stop;
                    outputCell(targetSheet, cell, val);

                    col = incrementColumn(col);
                }
                payCell = "N" + row.ToString();
                outputCell(targetSheet, payCell, line.Win.ToString());
                row++;
            }
            row = 5;
            foreach (PaylineDescription line in m_linePays)
            {
                if (line.Win <= 0)
                    continue;
                if (line.IsFreegameSet || line.IsModifierSet || line.HasWild)
                    //continue;
                    if (line.IsFreegameSet || line.IsModifierSet)
                        trim = true;

                col = "H";
                foreach (String stop in line.StopValues[stopSet].Values)
                {
                    cell = col + row.ToString();
                    if (line.StopValues[stopSet].Values.Count > 5)
                        break;
                    if (trim)
                        val = line.StopValues[stopSet].TrimValue(stop);
                    else
                        val = stop;
                    outputCell(targetSheet, cell, val);

                    col = incrementColumn(col);
                }
                row++;
            }
            m_rowCount = row;

            Globals.Program.Application.ScreenUpdating = true;

            updatePayLinks(targetBook);
        }

        private void updatePayLinks(Excel.Workbook book)
        {
            // need to update all links to point to the new target worksheet.
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            Globals.Program.Application.ScreenUpdating = false;

            String paySheetName = "Pays";
            String col = "A";
            String cell = "";
            String payCell = "";
            String targetCell = "L6";
            String payTargetCell = "X6";
            String targetCol = "L";
            int row = 5;
            String equation = "='Wins Combination'!";
            String val = "='Wins Combination'!A5";
            int stopSet = m_linePays[0].StopValues.Count - 1;
            Excel.Worksheet paySheet = null;
            // find the parsed reel worksheet
            int sheetIndex = getSheetIndex(book, paySheetName);
            if (sheetIndex > 0)
            {
                paySheet = book.Worksheets[sheetIndex];
                if (paySheet != null)
                    paySheet.Select();
                // copy the parsed reels to the pays sheet
                foreach (PaylineDescription line in m_linePays)
                {
                    if (line.Win <= 0)
                        continue;

                    col = "A";
                    targetCol = "L";
                    foreach (String stop in line.StopValues[stopSet].Values)
                    {
                        cell = col + row.ToString();
                        targetCell = targetCol + (row + 1).ToString();
                        val = equation + cell;
                        setCellFormula(paySheet, targetCell, val);

                        col = incrementColumn(col);
                        targetCol = incrementColumn(targetCol);
                    }
                    payCell = "$N$" + row.ToString();
                    payTargetCell = "X" + (row + 1).ToString();
                    setCellFormula(paySheet, payTargetCell, equation + payCell);
                    row++;
                }
            }
            Globals.Program.Application.ScreenUpdating = true;
        }
    }
}
