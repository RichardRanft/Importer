using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public class Utils
    {
        public String openBrace = "{";
        public String closeBrace = "}";
        public char[] cBrace = { '}' };
        public String arrayEnd = "},";
        public char[] arrayStop = { '}', ',' };
        public char[] doubleBackSlash = { '\\', '\\' };
        public char underscore = '_';
        public String endOfReel = "END_OF_REEL";

        public void exportReels(BallyReelSet set, String sheetName, Excel.Workbook targetBook)
        {
            int tableIndex = parseInteger(sheetName);
            String tableName = "";

            tableName = "Paytable" + tableIndex.ToString();

            // copy the match sheet template to a new worksheet
            copyMatchSheet(tableName, targetBook);
            // copy the pay sheet template to a new worksheet
            copyPaySheet(tableName, targetBook);

            Globals.Program.Application.ScreenUpdating = false;

            tableIndex++;
            Excel.Worksheet newSheet = createSheet(tableName, targetBook);
            set.SendToWorksheet(newSheet, "A1");

            // copy the reel data to the corresponding match and pay sheets
            //updateMatchLinks(newSheet, targetBook, tableName);
            //updatePayLinks(newSheet, targetBook, tableName);

            // get this baby out from under foot - move it to the end of the workbook
            moveSheetToEnd(newSheet, targetBook);

            // let the user see that we're working
            Globals.Program.Application.ScreenUpdating = true;
        }

        public Excel.Worksheet createSheet(String name, Excel.Workbook target)
        {
            // creates a new worksheet and passes it back
            // first, make sure we're not duplicating worksheets
            int sheetCount = Globals.Program.Application.ActiveWorkbook.Sheets.Count;
            for (int i = 1; i <= sheetCount; i++)
            {
                // if the sheet exists, pass it out and let the parser update
                // the contents from the file.
                if (target.Worksheets[i].Name == name)
                    return target.Worksheets[i];
            }

            // the sheet does not exist yet, so make a new one and pass it back
            Excel.Worksheet newWorksheet;
            newWorksheet = target.Worksheets.Add();
            newWorksheet.Name = name;
            return newWorksheet;
        }

        public void autofitColumn(String columnName, Excel.Worksheet sheet)
        {
            // hackish - there has to be a better way to select the whole column....
            String colName = columnName + "1";
            String colEnd = columnName + "1000";
            try
            {
                // Get the column and tell it to autofit
                Excel.Range columns = sheet.get_Range(colName, colEnd);
                columns.Columns.AutoFit();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Column AutoFit Failed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void moveSheetToEnd(Excel.Worksheet sheet, Excel.Workbook target)
        {
            // moves the sheet to the end of the workbook
            int sheetCount = target.Sheets.Count;
            try
            {
                sheet.Move(Type.Missing, target.Sheets[sheetCount]);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Worksheet Move Failed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void updateMatchLinks(BallyReelSet set, Excel.Worksheet sheet, Excel.Workbook target, String name)
        {
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String matchSheetName = trimName(name) + " Match";
            Excel.Worksheet matchSheet = null;
            int index = getSheetIndex("Game Info", target);
            Excel.Worksheet info = target.Worksheets[index];
            String link = "='" + matchSheetName + "'!$G$4";
            String nameCell = "B" + set.ToString();
            String linkCell = "C" + set.ToString();
            info.Range[nameCell].Value = matchSheetName;
            info.Range[linkCell].Value = link;

            // find the parsed reel worksheet
            int sheetIndex = getSheetIndex(matchSheetName, target);
            if (sheetIndex > 0)
            {
                matchSheet = target.Worksheets[getSheetIndex(matchSheetName, target)];

                // copy the parsed reels to the match sheet
                copyRange(sheet, matchSheet, "A1", "A300", "Q8");
                copyRange(sheet, matchSheet, "B1", "B300", "R8");
                copyRange(sheet, matchSheet, "C1", "C300", "S8");
                copyRange(sheet, matchSheet, "D1", "D300", "T8");
                copyRange(sheet, matchSheet, "E1", "E300", "U8");
            }
            else
                return;
        }

        public void updatePayLinks(Excel.Worksheet sheet, Excel.Workbook target, String name)
        {
            // need to update all links to point to the new target worksheet.
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String paySheetName = trimName(name) + " Pays";
            Excel.Worksheet paySheet = null;
            // find the parsed reel worksheet
            int sheetIndex = getSheetIndex(paySheetName, target);
            if (sheetIndex > 0)
            {
                paySheet = target.Worksheets[sheetIndex];
                // copy the parsed reels to the match sheet
                copyRange(sheet, paySheet, "A1", "A300", "A6");
                copyRange(sheet, paySheet, "B1", "B300", "C6");
                copyRange(sheet, paySheet, "C1", "C300", "E6");
                copyRange(sheet, paySheet, "D1", "D300", "G6");
                copyRange(sheet, paySheet, "E1", "E300", "I6");
            }
        }

        public void copyMatchSheet(String fileName, Excel.Workbook target)
        {
            String name = trimName(fileName) + " Match";
            if (findSheet(name, target))
                return;
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                // if the sheet exists, pass it out and let the parser update
                // the contents from the file.
                if (target.Worksheets[i].Name == "Match")
                {
                    //Excel.Worksheet newWorksheet = active.Worksheets.Add();
                    target.Worksheets[i].Copy(target.Worksheets[target.Worksheets.Count]);
                    Excel.Worksheet newWorksheet = target.Worksheets[target.Worksheets.Count - 1];
                    newWorksheet.Name = name;
                    moveSheetToEnd(newWorksheet, target);
                    return;
                }
            }
        }

        public void copyPaySheet(String fileName, Excel.Workbook target)
        {
            String name = trimName(fileName) + " Pays";
            if (findSheet(name, target))
                return;
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                // if the sheet exists, pass it out and let the parser update
                // the contents from the file.
                if (target.Worksheets[i].Name == "Pays")
                {
                    //Excel.Worksheet newWorksheet = active.Worksheets.Add();
                    target.Worksheets[i].Copy(target.Worksheets[target.Worksheets.Count]);
                    Excel.Worksheet newWorksheet = target.Worksheets[target.Worksheets.Count - 1];
                    newWorksheet.Name = name;
                    moveSheetToEnd(newWorksheet, target);
                    return;
                }
            }
        }

        public String trimName(String name)
        {
            name = stripFileName(name);
            if (name.Length > 24)
            {
                String[] nameParts = name.Split(underscore);
                String shortName = "";
                bool firstWord = true;
                int firstv = -1;
                int firstcr = -1;
                int countDown = 0;
                foreach (String part in nameParts)
                {
                    firstv = part.IndexOf("v");
                    firstcr = part.IndexOf("cr");
                    if (firstv >= 0 || firstcr >= 0 || countDown > 0)
                    {
                        if (countDown > 0 && (firstv < shortName.Length && firstcr < shortName.Length))
                        {
                            shortName += ("_" + part);
                            countDown--;
                        }
                        else
                        {
                            shortName = shortName + part;
                        }
                        if (firstWord)
                        {
                            shortName = shortName + "_";
                            firstWord = false;
                        }
                        if (part == "vf")
                        {
                            countDown = 2;
                        }
                    }
                }
                name = shortName;
            }
            return name;
        }

        public void copyRange(Excel.Worksheet source, Excel.Worksheet dest, String startCell, String endCell, String startDest)
        {
            // get the length of the range - this might need to change later, figured it would be
            // easier if this just adapted by itself.
            int end = System.Convert.ToInt32(endCell.Substring(1)) + System.Convert.ToInt32(startDest.Substring(1));
            String col = startDest.Substring(0, 1);
            String row = end.ToString();
            String destEnd = col + row;
            // select the source worksheet
            source.Select();
            // get the cell range
            Excel.Range reel = source.get_Range(startCell, endCell);
            // copy our data
            reel.Copy();
            // select the destination worksheet
            dest.Select();
            // pick our destination cell
            Excel.Range targetCell = dest.get_Range(startDest, destEnd);
            try
            {
                targetCell.Select();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message.ToString(), "Error - target cells not empty", System.Windows.Forms.MessageBoxButtons.OK);
                return;
            }
            // paste our data
            dest.Paste(targetCell);
        }

        public String stripFileName(String name)
        {
            int slashIndex = 0;
            int dotIndex = 0;
            for (int i = 0; i < name.Length; i++)
            {
                if (name[i].ToString() == "\\")
                    slashIndex = i + 1;
            }
            for (int j = 0; j < name.Length; j++)
            {
                if (name[j].ToString() == ".")
                    dotIndex = j;
            }
            return name.Substring(slashIndex, dotIndex - slashIndex);
        }

        public bool findSheet(String sheetName, Excel.Workbook target)
        {
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == sheetName)
                    return true;
            }
            return false;
        }

        public int getSheetIndex(String sheetName, Excel.Workbook target)
        {
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == sheetName)
                    return i;
            }
            return 0;
        }

        public int parseInteger(String data)
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
