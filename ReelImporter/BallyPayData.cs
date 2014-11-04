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
    public abstract partial class BallyPayData
    {
        public abstract BallyPayType PayType
        {
            get;
        }

        public abstract void Parse(StreamReader inStream, String line, PayParserState parseState);

        protected abstract void exportPays(String sheetName, Excel.Workbook targetBook);

        protected bool outputCell(Excel.Worksheet targetSheet, String cell, String value)
        {
            bool result = true;
            try
            {
                // try, because it might fail
                targetSheet.Cells.Range[cell, Type.Missing].Value2 = value;
                int row = parseRow(cell);
                if ((row % 2) != 0)
                    targetSheet.Cells.Range[cell, Type.Missing].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray;
                else
                    targetSheet.Cells.Range[cell, Type.Missing].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhiteSmoke;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + e.Message, "File Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                result = false;
            }
            return result;
        }

        protected bool setCellFormula(Excel.Worksheet targetSheet, String cell, String value)
        {
            bool result = true;
            try
            {
                // try, because it might fail
                Excel.Range targetCell = targetSheet.get_Range(cell);
                targetCell.Value2 = "";
                targetCell.Formula = value;
                int row = parseRow(cell);
                if ((row % 2) != 0)
                    if ((row % 2) != 0)
                        targetSheet.Cells.Range[cell, Type.Missing].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray;
                    else
                        targetSheet.Cells.Range[cell, Type.Missing].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhiteSmoke;

                targetCell.Calculate();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + e.Message, "File Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                result = false;
            }
            return result;
        }

        protected int getSheetIndex(Excel.Workbook target, String sheetName)
        {
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == sheetName)
                    return i;
            }
            return 0;
        }

        protected String incrementColumn(String current)
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

        protected String parseCol(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[\d]");
            return digits.Replace(data, "");
        }

        protected int parseRow(String data)
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
