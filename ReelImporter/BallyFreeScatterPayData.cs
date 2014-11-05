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

            IComparer<PaylineDescription> payComparer = new PaylineSorter();
            m_freeScatterPays.Sort(payComparer);
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
            int stopCount = 0;
            String cell = col + row.ToString();
            String payCell = "N" + row.ToString();
            int stopSet = m_freeScatterPays[0].StopValues.Count - 1;
            bool trim = true;
            String val = "";
            foreach (PaylineDescription line in m_freeScatterPays)
            {
                if (line.Win <= 0)
                    continue;
                if (line.IsFreegameSet || line.IsModifierSet || line.HasWild)
                    //continue;
                    if (line.IsFreegameSet || line.IsModifierSet)
                        trim = true;

                col = "A";
                stopCount = 0;
                foreach (String stop in line.StopValues[stopSet].Values)
                {
                    cell = col + row.ToString();
                    if (stopCount == 5)
                        break;
                    stopCount++;
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
            foreach (PaylineDescription line in m_freeScatterPays)
            {
                if (line.Win <= 0)
                    continue;
                if (line.IsFreegameSet || line.IsModifierSet || line.HasWild)
                    //continue;
                    if (line.IsFreegameSet || line.IsModifierSet)
                        trim = true;

                col = "H";
                stopCount = 0;
                foreach (String stop in line.StopValues[stopSet].Values)
                {
                    cell = col + row.ToString();
                    if (stopCount == 5)
                        break;
                    stopCount++;
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
            int stopCount = 0;
            String equation = "='Wins Combination'!";
            String val = "='Wins Combination'!A5";
            int stopSet = m_freeScatterPays[0].StopValues.Count - 1;
            Excel.Worksheet paySheet = null;
            // find the parsed reel worksheet
            int sheetIndex = getSheetIndex(book, paySheetName);
            if (sheetIndex > 0)
            {
                paySheet = book.Worksheets[sheetIndex];
                if (paySheet != null)
                    paySheet.Select();
                // copy the parsed reels to the pays sheet
                foreach (PaylineDescription line in m_freeScatterPays)
                {
                    if (line.Win <= 0)
                        continue;

                    col = "A";
                    targetCol = "L";
                    stopCount = 0;
                    foreach (String stop in line.StopValues[stopSet].Values)
                    {
                        if (stopCount == 5)
                            break;
                        stopCount++;
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
