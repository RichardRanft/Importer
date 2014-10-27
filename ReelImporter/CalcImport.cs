using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public class CalcImport
    {
        private Excel.Window excelWin;
        private Excel.Range startCell;
        private Excel.Workbook source;
        private Excel.Worksheet sourceSheet;
        private Excel.Workbook target;
        private Excel.Worksheet targetSheet;
        private ImporterRibbon ribbon;

        public CalcImport()
        {
            excelWin = Globals.Program.Application.ActiveWindow;
        }

        public CalcImport(ImporterRibbon parent)
        {
            excelWin = Globals.Program.Application.ActiveWindow;
            ribbon = parent;
        }

        public void setTargetWorkbook(Excel.Workbook book)
        {
            target = book;
        }

        public void setTargetWorksheet(Excel.Worksheet sheet)
        {
            targetSheet = sheet;
        }

        public void setStartCell(Excel.Range cell)
        {
            startCell = cell;
        }

        public void openWorkbook(String name)
        {
            source = Globals.Program.Application.Workbooks.Open(name);
        }

        public void importReels()
        {
            // CYA buddy - when the add-ins load there is apparently no active window, so
            // catch that here.
            // stop screen updates - reduces run time by nearly a factor of 10
            Globals.Program.Application.ScreenUpdating = false;
            if (excelWin == null)
                excelWin = Globals.Program.Application.ActiveWindow;
            if (source.ActiveSheet == null)
            {
                System.Windows.Forms.MessageBox.Show("No source Workbook or Worksheet available.", "Error - No source.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }
            sourceSheet = source.ActiveSheet;

            // Copy the parsed reels to the match sheet
            String start = startCell.Address.ToString().Replace("$", "");
            if (startCell.Column > 5 || !checkValue(source, source.ActiveSheet, start))
            {
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("Starting cell appears to be invalid.  Continue?", "Check Starting Cell", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Information);
                if (result == System.Windows.Forms.DialogResult.No)
                {
                    Globals.Program.Application.ScreenUpdating = true;
                    return;
                }
            }
            int startRow = 1;
            try
            {
                startRow = System.Convert.ToInt32(start.Substring(1));
            }
            catch (Exception convEx)
            {
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("You appear to have more than one selected starting cell.  Please select a single starting cell and try again.", "Check Starting Cell", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Globals.Program.Application.ScreenUpdating = true;
                return;
            }
            String end = parseCol(start) + (startRow + 300).ToString();

            // We know this is one reel, so copy it
            copyRange(sourceSheet, targetSheet, start, end, "B8");
            // Now we start looking for more reels.  Frequently document reels are separated by
            // blank columns so we can't just march on down the row.
            start = incrementColumn(start);
            while (!checkValue(source, sourceSheet, start))
            {
                start = incrementColumn(start);
            }
            end = parseCol(start) + (startRow + 300).ToString();
            copyRange(sourceSheet, targetSheet, start, end, "C8");

            start = incrementColumn(start);
            while (!checkValue(source, sourceSheet, start))
            {
                start = incrementColumn(start);
            }
            end = parseCol(start) + (startRow + 300).ToString();
            copyRange(sourceSheet, targetSheet, start, end, "D8");

            start = incrementColumn(start);
            while (!checkValue(source, sourceSheet, start))
            {
                start = incrementColumn(start);
            }
            end = parseCol(start) + (startRow + 300).ToString();
            copyRange(sourceSheet, targetSheet, start, end, "E8");

            start = incrementColumn(start);
            while (!checkValue(source, sourceSheet, start))
            {
                start = incrementColumn(start);
            }
            end = parseCol(start) + (startRow + 300).ToString();
            copyRange(sourceSheet, targetSheet, start, end, "F8");

            Globals.Program.Application.ScreenUpdating = true;

            this.source.Close();
            this.ribbon.EnableCalcImport(false);
        }

        private bool checkValue(Excel.Workbook book, Excel.Worksheet sheet, String cell)
        {
            book.Activate();
            if (book.ActiveSheet != sheet)
                sheet.Select();
            Excel.Range test = sheet.get_Range(cell);
            try
            {
                sheet.Select(test);
            }
            catch (Exception e)
            {
                // apparently, an empty cell is not valid for selection.
                return true;
            }
            try
            {
                if (test.Value2 == "" || test.Value2 == null)
                    return false;
                else
                    return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private void copyRange(Excel.Worksheet origin, Excel.Worksheet dest, String startCell, String endCell, String startDest)
        {
            // get the length of the range - this might need to change later, figured it would be
            // easier if this just adapted by itself.
            int end = System.Convert.ToInt32(endCell.Substring(1)) + System.Convert.ToInt32(startDest.Substring(1));
            String col = startDest.Substring(0, 1);
            String row = end.ToString();
            // select the source worksheet
            this.source.Activate();
            origin.Select();
            // get the cell range
            String blank = "\0";
            Excel.Range firstBlank = origin.Cells.Find(blank, origin.get_Range(startCell), Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByColumns);
            if (firstBlank != null)
            {
                String temp = firstBlank.Cells.get_Address().ToString().Replace("$", "");
                int endRow = System.Convert.ToInt32(temp.Substring(1));
                row = endRow.ToString();
            }
            int destEndRow = (8 + end) - parseRow(startCell);
            String destEnd = String.Concat(col, destEndRow);
            String sourceEnd = String.Concat(startCell.Substring(0, 1), row);
            Excel.Range reel = origin.get_Range(startCell, sourceEnd);
            // copy our data
            reel.Copy();
            // select the destination worksheet
            this.target.Activate();
            this.targetSheet.Select();
            // pick our destination cell
            Excel.Range targetCell = dest.get_Range(startDest, destEnd);
            try
            {
                targetCell.Select();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + e.Message.ToString(), "Error - target cells not empty", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }
            // paste our data
            try
            {
                targetCell.PasteSpecial(Excel.XlPasteType.xlPasteValues);
            }
            catch (Exception pasteEx)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + pasteEx.Message.ToString(), "Error - Can't paste cells", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }
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
            String row = current.Substring(nextColumn.Length, (current.Length - nextColumn.Length));
            String nextCell = nextColumn + row;
            return nextCell;
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
