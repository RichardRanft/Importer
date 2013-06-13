using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public class HeaderImport
    {
        private String currentFolder;
        private Excel.Worksheet newSheet;
        private Array fileList;
        private Excel.Window excelWin;
        private String currentCell;
        private String openBrace = "{";
        private String closeBrace = "}";
        private char[] doubleBackSlash = {'\\','\\'};
        private String endOfReel = "END_OF_REEL";
        private String line;
        private String tempLine;
        private String[] parsedRow;
        private bool moveSheet;
        private Excel.Workbook target;
        private int currentReelSet;

        public HeaderImport(Excel.Window window)
        {
            if (excelWin != null)
                excelWin = window;
            else
                excelWin = Globals.Program.Application.ActiveWindow;
        }

        public String getFolder()
        {
            return currentFolder;
        }

        public void importFolder(String folder)
        {
            // CYA buddy - when the add-ins load there is apparently no active window, so
            // catch that here.
            if (excelWin == null)
                excelWin = Globals.Program.Application.ActiveWindow;
            target = Globals.Program.Application.ActiveWorkbook;
            // get a list of all header files in the selected directory
            currentFolder = folder;
            currentReelSet = 7;
            fileList = Directory.GetFiles(currentFolder, "*.h");
            if (fileList.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("There are no header files in the selected folder.", "No Header Files Found.");
                return;
            }

            // run down the list and import each file into an Excel sheet
            for (int index = 0; index < fileList.Length; index++)
            {
                Globals.Program.Application.ScreenUpdating = false;
                copyMatchSheet(fileList.GetValue(index).ToString());
                copyPaySheet(fileList.GetValue(index).ToString());
                Excel.Worksheet temp = importFile(fileList.GetValue(index).ToString(), (index + 1).ToString());
                updateMatchLinks(temp, fileList.GetValue(index).ToString());
                updatePayLinks(temp, fileList.GetValue(index).ToString());
                Globals.Program.Application.ScreenUpdating = true;
                currentReelSet++;
            }
        }

        private void updateMatchLinks(Excel.Worksheet sheet, String name)
        {
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String matchSheetName = stripFileName(name) + " Match";
            Excel.Worksheet matchSheet = null;
            Excel.Worksheet info = target.Worksheets[1];
            if (info.Name == "Game Info")
            {
                String link = "='" + matchSheetName + "'!$G$4";
                String nameCell = "B" + currentReelSet.ToString();
                String linkCell = "C" + currentReelSet.ToString();
                info.Range[nameCell].Value = matchSheetName;
                info.Range[linkCell].Value = link;
            }

            // find the parsed reel worksheet
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == matchSheetName)
                {
                    matchSheet = target.Worksheets[i];
                }
            }
            // copy the parsed reels to the match sheet
            copyRange(sheet, matchSheet, "A1", "A300", "Q8");
            copyRange(sheet, matchSheet, "B1", "B300", "R8");
            copyRange(sheet, matchSheet, "C1", "C300", "S8");
            copyRange(sheet, matchSheet, "D1", "D300", "T8");
            copyRange(sheet, matchSheet, "E1", "E300", "U8");
        }

        private void copyRange(Excel.Worksheet source, Excel.Worksheet dest, String startCell, String endCell, String startDest)
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
            dest.Select(targetCell);
            // paste our data
            dest.Paste(targetCell);
        }

        private void updatePayLinks(Excel.Worksheet sheet, String name)
        {
            // need to update all links to point to the new target worksheet.
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String paySheetName = stripFileName(name) + " Pays";
            Excel.Worksheet paySheet = null;
            // find the parsed reel worksheet
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == paySheetName)
                {
                    paySheet = target.Worksheets[i];
                }
            }
            // copy the parsed reels to the match sheet
            copyRange(sheet, paySheet, "A1", "A300", "A6");
            copyRange(sheet, paySheet, "B1", "B300", "C6");
            copyRange(sheet, paySheet, "C1", "C300", "E6");
            copyRange(sheet, paySheet, "D1", "D300", "G6");
            copyRange(sheet, paySheet, "E1", "E300", "I6");
        }

        private void getCalcReels(String filename)
        {
            // No idea at the moment how I'll manage to find these in other files
            // At least I should be able to look for "REEL 1", "REEL 2" and so on.
        }

        private Excel.Worksheet importFile(String fileName, String index)
        {
            // clean up the file name to use as the worksheet name
            String[] trimmedName = fileName.Split(doubleBackSlash);
            int end = trimmedName.Length;
            String sheetName = trimmedName[end - 1];

            // create a new worksheet for our reel data
            newSheet = createSheet(sheetName);
            // read in the reel data
            using (StreamReader inputFile = new StreamReader(fileName))
            {
                int row = 1;
                String column = "A";
                String cellValue = "";
                bool addCell = false;
                while ((line = inputFile.ReadLine()) != null)
                {
                    // clean up and parse the line
                    tempLine = line.Trim();
                    parsedRow = line.Split(',');

                    for (int i = 0; i < parsedRow.Length; i++)
                    {
                        cellValue = parsedRow[i].Trim();
                        // check for start of a reel
                        if (cellValue == openBrace || tempLine == openBrace)
                        {
                            addCell = true;
                            continue;
                        }
                        // check for the end of a reel
                        if(cellValue == closeBrace)
                        {
                            addCell = false;
                            continue;
                        }
                        // check for last entry in a reel
                        if (cellValue == endOfReel)
                        {
                            // we're finished adding reel entries, tidy up our worksheet
                            // as we go along
                            autofitColumn(column);

                            // advance to the next reel in the data set
                            if (column == "A")
                                column = "B";
                            else if (column == "B")
                                column = "C";
                            else if (column == "C")
                                column = "D";
                            else if (column == "D")
                                column = "E";
                            else if (column == "E")
                                column = "F";
                            else if (column == "F")
                                column = "G";
                            else if (column == "G")
                                column = "H";
                            else if (column == "H")
                                column = "I";

                            row = 1;
                            continue;
                        }
                        // we're ignoring blanks
                        if (cellValue == "")
                        {
                            continue;
                        }
                        // Ok, we've got a valid reel entry - add it to the current column
                        if (addCell)
                        {
                            // dummy-proofing
                            if (row < 1)
                                row = 1;
                            // build the cell name.  post-increment our row counter
                            currentCell = column + (row++).ToString();
                            try
                            {
                                // try, because it might fail
                                newSheet.Cells.Range[currentCell, Type.Missing].Value2 = cellValue;
                            }
                            catch (Exception e)
                            {
                                System.Windows.Forms.MessageBox.Show(e.Message, "File Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
            // get this baby out from under foot - move it to the end of the workbook
            if (moveSheet)
                moveSheetToEnd(newSheet);
            return newSheet;
        }

        private Excel.Worksheet createSheet(String name)
        {
            // creates a new worksheet and passes it back
            // first, make sure we're not duplicating worksheets
            int sheetCount = Globals.Program.Application.ActiveWorkbook.Sheets.Count;
            for (int i = 1; i <= sheetCount; i++)
            {
                // if the sheet exists, pass it out and let the parser update
                // the contents from the file.
                if (target.Worksheets[i].Name == name)
                {
                    moveSheet = false;
                    return target.Worksheets[i];
                }
            }
            // the sheet does not exist yet, so make a new one and pass it back
            Excel.Worksheet newWorksheet;
            newWorksheet = target.Worksheets.Add();
            newWorksheet.Name = name;
            return newWorksheet;
        }

        private void autofitColumn(String columnName)
        {
            // hackish - there has to be a better way to select the whole column....
            String colName = columnName + "1";
            String colEnd = columnName + "1000";
            try
            {
                // Get the column and tell it to autofit
                Excel.Range columns = newSheet.get_Range(colName, colEnd);
                columns.Columns.AutoFit();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Column AutoFit Failed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void moveSheetToEnd(Excel.Worksheet sheet)
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

        private void copyMatchSheet(String fileName)
        {
            String name = stripFileName(fileName);
            name += " Match";
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                // if the sheet exists, pass it out and let the parser update
                // the contents from the file.
                if (target.Worksheets[i].Name == "Match" && !findSheet(name))
                {
                    //Excel.Worksheet newWorksheet = active.Worksheets.Add();
                    target.Worksheets[i].Copy(target.Worksheets[target.Worksheets.Count]);
                    Excel.Worksheet newWorksheet = target.Worksheets[target.Worksheets.Count - 1];
                    newWorksheet.Name = name;
                    moveSheetToEnd(newWorksheet);
                    return;
                }
            }
        }

        private void copyPaySheet(String fileName)
        {
            String name = stripFileName(fileName);
            name += " Pays";
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                // if the sheet exists, pass it out and let the parser update
                // the contents from the file.
                if (target.Worksheets[i].Name == "Pays" && !findSheet(name))
                {
                    //Excel.Worksheet newWorksheet = active.Worksheets.Add();
                    target.Worksheets[i].Copy(target.Worksheets[target.Worksheets.Count]);
                    Excel.Worksheet newWorksheet = target.Worksheets[target.Worksheets.Count - 1];
                    newWorksheet.Name = name;
                    moveSheetToEnd(newWorksheet);
                    return;
                }
            }
        }

        private bool findSheet(String sheetName)
        {
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == sheetName)
                    return true;
            }
            return false;
        }

        private String stripFileName(String name)
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
    }
}
