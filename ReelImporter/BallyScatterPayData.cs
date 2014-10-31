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
                return m_scatterPays;
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
                            if (parseState.ScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveScatterPay();
                            }
                            break;
                        }

                        // could be end of a reelstop definition, or moving up a level
                        if (line == m_util.arrayEnd)
                        {
                            parseState.LeaveArrayLevel();
                            if (parseState.ScatterPayStart)
                            {
                                if (parseState.StateEnteredLevel[(int)PayReadState.SCATTER_PAYSTART] == parseState.ArrayDepth)
                                    parseState.LeaveScatterPay();
                            }
                            break;
                        }
                    }

                    payline = new PaylineDescription();
                    payline.Add(line, m_util);
                    if (parseState.ScatterPayStart)
                    {
                        m_scatterPays.Add(payline);
                    }
                    else
                    {
                        parseState.ResetScatterPay();
                        break;
                    }
                }
            }
        }

        public void SendToWorksheet(Excel.Workbook targetBook, Excel.Worksheet targetSheet)
        {
            Globals.Program.Application.ScreenUpdating = false;
            // reel entries A5~E5, pay at M5
            String col = "A";
            int row = 5;
            String cell = col + row.ToString();
            String payCell = "N" + row.ToString();
            int stopSet = m_scatterPays[0].StopValues.Count - 1;
            bool trim = false;
            String val = "";
            foreach(PaylineDescription line in m_scatterPays)
            {
                if (line.IsFreegameSet || line.HasWild)
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
            Globals.Program.Application.ScreenUpdating = true;
        }

        private bool outputCell(Excel.Worksheet targetSheet, String cell, String value)
        {
            bool result = true;
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

        protected override void exportPays(String sheetName, Excel.Workbook targetBook)
        {

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

        private String parseCol(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[\d]");
            return digits.Replace(data, "");
        }

        private int parseRow(String data)
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
    }
}
