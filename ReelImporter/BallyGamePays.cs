﻿using System;
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
        private BallyLinePayData m_linePays;
        private BallyScatterPayData m_scatterPays;
        private BallyFreeLinePayData m_freeLinePays;
        private BallyFreeScatterPayData m_freeScatterPays;
        private List<String> m_symbolList;
        
        private PayParserState m_parseState;
        private Utils m_util;

        public BallyGamePays()
        {
            m_linePays = new BallyLinePayData();
            m_scatterPays = new BallyScatterPayData();
            m_freeLinePays = new BallyFreeLinePayData();
            m_freeScatterPays = new BallyFreeScatterPayData();
            m_symbolList = new List<String>();

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
            
            String line = "";

            using (inStream)
            {
                while (line != null)
                {
                    try
                    {
                        line = inStream.ReadLine();
                    }
                    catch(ObjectDisposedException ex)
                    {
                        break;
                    }
                    // strip comments
                    if (line.Contains("/"))
                    {
                        int pos = line.IndexOf("/");
                        line = line.Remove(pos);
                    }

                    line = line.Trim();

                    if (line.Length == 0 || line == "")
                        continue;

                    // look for symbols
                    if (line == "symbols")
                    {
                        m_parseState.EnterSymbols();
                        parseSymbols(inStream, line);
                    }
                    
                    // look for pays
                    if (line == "linepays")
                    {
                        m_parseState.EnterLinePay();
                        m_linePays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "freegame_linepays")
                    {
                        m_parseState.EnterFreeLinePay();
                        m_freeLinePays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "scatterpays")
                    {
                        m_parseState.EnterScatterPay();
                        m_scatterPays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "freegame_scatterpays")
                    {
                        m_parseState.EnterFreeScatterPay();
                        m_freeScatterPays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "featuredefs")
                    {
                        break;
                    }
                }
            }
        }

        private void parseSymbols(StreamReader inStream, String line)
        {
            bool lineHasOpenBrace = false;
            bool lineHasCloseBrace = false;
            String[] parts;
            String symbol = "";

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
                        m_parseState.EnterArrayLevel();
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
                            m_parseState.LeaveArrayLevel();
                            if (m_parseState.SymbolStart)
                            {
                                if (m_parseState.StateEnteredLevel[(int)PayReadState.SYMBOLSTART] == m_parseState.ArrayDepth)
                                    m_parseState.LeaveSymbols();
                            }
                            break;
                        }

                        // could be end of a reelstop definition, or moving up a level
                        if (line == m_util.arrayEnd)
                        {
                            m_parseState.LeaveArrayLevel();
                            if (m_parseState.SymbolStart)
                            {
                                if (m_parseState.StateEnteredLevel[(int)PayReadState.SYMBOLSTART] == m_parseState.ArrayDepth)
                                    m_parseState.LeaveSymbols();
                            }
                            break;
                        }
                    }

                    if (m_parseState.CurrentPayType == BallyPayType.NONE)
                    {
                        symbol = line.Replace(m_util.openBrace, "");
                        symbol = symbol.Replace(m_util.closeBrace, "");
                        parts = symbol.Split(m_util.comma);
                        m_symbolList.Add(parts[0]);
                        m_parseState.ResetSymbols();
                    }
                }
            }
        }

        public void ExportPays(String sheetName, Excel.Workbook targetBook)
        {
            // stop screen updates - reduces run time by nearly a factor of 10
            m_symbolList.Sort();

            Globals.Program.Application.ScreenUpdating = false;
            Excel.Window excelWin = Globals.Program.Application.ActiveWindow;

            if (targetBook == null)
            {
                targetBook = excelWin.Application.ActiveWorkbook;
                if (targetBook == null)
                {
                    System.Windows.Forms.MessageBox.Show("No source Workbook or Worksheet available.", "Error - No source.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }
            }

            int sheetIndex = getSheetIndex("Wins Combination", targetBook);
            Excel.Worksheet targetSheet = targetBook.Worksheets[sheetIndex];
            
            String payStartCell = "R6";
            String cell = payStartCell;
            String col = parseCol(payStartCell);
            int row = 0;
            for (int i = 0; i < 4; i++)
            {
                row = parseRow(payStartCell);
                foreach (String symbol in m_symbolList)
                {
                    OutputCell(targetSheet, cell, symbol);
                    row++;
                    cell = col + row.ToString();
                }
                row = parseRow(payStartCell);
                col = incrementColumn(cell);
                cell = col + row.ToString();
            }
        }

        public bool OutputCell(Excel.Worksheet targetSheet, String cell, String value)
        {
            bool result = true;
            int row = parseRow(cell);
            if (row < 1)
                row = 1;
            try
            {
                // try, because it might fail
                targetSheet.Cells.Range[cell, Type.Missing].Value2 = value;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + e.Message, "File Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                result = false;
            }
            return result;
        }

        private String incrementColumn(String current)
        {
            String nextColumn = parseCol(current);
            if (nextColumn.Length < 2)
            {
                char temp = nextColumn[0];
                temp++;
                if (temp <= 'Z')
                    nextColumn = temp.ToString();
                else
                {
                    nextColumn = "AA";
                }
            }
            else
            {
                char temp = nextColumn[1];
                temp++;
                nextColumn = "A" + temp.ToString();
            }
            return nextColumn;
        }

        public String parseCol(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[\d]");
            return digits.Replace(data, "");
        }

        public int parseRow(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[^\d]");
            int value = 0;
            try
            {
                value = System.Convert.ToInt32(digits.Replace(data, ""));
            }
            catch (Exception e)
            {
                value = 0;
            }
            return value;
        }

        private int getSheetIndex(String sheetName, Excel.Workbook target)
        {
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == sheetName)
                    return i;
            }
            return 0;
        }
    }
}
