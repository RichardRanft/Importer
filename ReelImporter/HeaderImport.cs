using System;
using System.Collections;
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
    public class KeyValue
    {
        private String StringItem;
        private Int32 ItemCount;

        public String Key
        {
            get { return StringItem; }
            set { StringItem = value; }
        }

        public Int32 Value
        {
            get { return ItemCount; }
            set { ItemCount = value; }
        }
    }

    public class fileSorter : IComparer
    {
        int IComparer.Compare(Object a, Object b)
        {
            int value = 0; // a and b are equal
            // split the file name into parts and compare for numerical ordering
            String first;
            String second;
            try
            {
                first = Convert.ToString(a);
            }
            catch(Exception e)
            {
                return 0;
            }
            try
            {
                second = Convert.ToString(b);
            }
            catch (Exception e)
            {
                return 0;
            }
            String[] firstParts = first.Split('_');
            String[] secondParts = second.Split('_');
            for (int i = 1; i < firstParts.Count(); i++)
            {
                int val1, val2;
                String temp1 = "";
                String temp2 = "";
                bool convertFailed = false;
                try
                {
                    val1 = Convert.ToInt32(firstParts.GetValue(i));
                }
                catch (Exception e)
                {
                    val1 = 0;
                    temp1 = Convert.ToString(firstParts.GetValue(i));
                    convertFailed = true;
                }
                try
                {
                    val2= Convert.ToInt32(secondParts.GetValue(i));
                }
                catch (Exception e)
                {
                    val2 = 0;
                    temp2 = Convert.ToString(secondParts.GetValue(i));
                    convertFailed = true;
                }
                if (convertFailed)
                {
                    value = (new CaseInsensitiveComparer()).Compare(temp1, temp2);
                }
                else
                {
                    if (val1 > val2)
                    {
                        value = 1;
                        break;
                    }
                    else if (val2 > val1)
                    {
                        value = -1;
                        break;
                    }
                    else
                        value = 0;
                }
                if (value != 0)
                    break;
            }
            return value;
        }
    }

    public class HeaderImport
    {
        private String currentFolder;
        private Excel.Worksheet newSheet;
        private Array fileList;
        private Excel.Window excelWin;
        private String currentCell;
        private String openBrace = "{";
        private String closeBrace = "}";
        private char[] cBrace = { '}' };
        private String arrayEnd = "},";
        private char[] arrayStop = { '}', ',' };
        private char[] doubleBackSlash = {'\\', '\\'};
        private char underscore = '_';
        private String endOfReel = "END_OF_REEL";
        private String line;
        private String tempLine;
        private String[] parsedRow;
        private Excel.Workbook target;
        private int currentReelSet;
        private ParserState m_parseState;
        private BallyReelGame m_gameSet;
        private BallyGamePays m_gamePays;

        private bool moveSheet;

        public HeaderImport(Excel.Window window)
        {
            if (excelWin != null)
                excelWin = window;
            else
                excelWin = Globals.Program.Application.ActiveWindow;
            m_parseState = new ParserState();
        }

        public String getFolder()
        {
            return currentFolder;
        }

        public void importFolder(String folder, ReelDataType type = ReelDataType.SHFL)
        {
            // CYA buddy - when the add-ins load there is apparently no active window, so
            // catch that here.
            if (excelWin == null)
                excelWin = Globals.Program.Application.ActiveWindow;
            target = Globals.Program.Application.ActiveWorkbook;
            // get a list of all header files in the selected directory
            currentFolder = folder;
            // This is the starting row for the Game Info worksheet reel match summary.
            currentReelSet = 7;
            IComparer comp = new fileSorter();

            if ( type == ReelDataType.SHFL ) // Equinox/SLV reel definition header files
                fileList = Directory.GetFiles(currentFolder, "*.h");
            if ( type == ReelDataType.BALLY ) // Alpha II paytable.cfg file
                fileList = Directory.GetFiles(currentFolder, "*.cfg");

            Array.Sort(fileList, comp);
            if (fileList.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("There are no header files in the selected folder.", "No Header Files Found.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            if (type == ReelDataType.SHFL)
            {
                // run down the list and import each file into an Excel sheet
                for (int index = 0; index < fileList.Length; index++)
                {
                    String name = trimName(fileList.GetValue(index).ToString()) + " Match";
                    if (findSheet(name))
                    {
                        System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("These headers have already been imported.  Would you like to import them again?", "Re-import headers?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Information);
                        if (result == System.Windows.Forms.DialogResult.No)
                            return;
                    }

                    // copy the match sheet template to a new worksheet
                    copyMatchSheet(fileList.GetValue(index).ToString());
                    // copy the pay sheet template to a new worksheet
                    copyPaySheet(fileList.GetValue(index).ToString());
                    // stop screen updates - reduces run time by nearly a factor of 10
                    Globals.Program.Application.ScreenUpdating = false;

                    // import the data
                    Excel.Worksheet temp = importFile(fileList.GetValue(index).ToString(), (index + 1).ToString(), type);

                    // copy the reel data to the corresponding match and pay sheets
                    updateMatchLinks(temp, fileList.GetValue(index).ToString());
                    updatePayLinks(temp, fileList.GetValue(index).ToString());
                    // let the user see that we're working
                    Globals.Program.Application.ScreenUpdating = true;
                    currentReelSet++;
                }
            }
            if (type == ReelDataType.BALLY)
            {
                m_gamePays = new BallyGamePays();
                m_gameSet = new BallyReelGame();
                importBallyFile(fileList.GetValue(0).ToString(), "1");
            }
        }

        private void updateMatchLinks(Excel.Worksheet sheet, String name)
        {
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String matchSheetName = trimName(name) + " Match";
            Excel.Worksheet matchSheet = null;
            int index = getSheetIndex("Game Info");
            Excel.Worksheet info = target.Worksheets[index];
            String link = "='" + matchSheetName + "'!$G$4";
            String nameCell = "B" + currentReelSet.ToString();
            String linkCell = "C" + currentReelSet.ToString();
            info.Range[nameCell].Value = matchSheetName;
            info.Range[linkCell].Value = link;
            
            // find the parsed reel worksheet
            int sheetIndex = getSheetIndex(matchSheetName);
            if (sheetIndex > 0)
            {
                matchSheet = target.Worksheets[getSheetIndex(matchSheetName)];

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

        private void updatePayLinks(Excel.Worksheet sheet, String name)
        {
            // need to update all links to point to the new target worksheet.
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String paySheetName = trimName(name) + " Pays";
            Excel.Worksheet paySheet = null;
            // find the parsed reel worksheet
            int sheetIndex = getSheetIndex(paySheetName);
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

        private void getCalcReels(String filename)
        {
            // No idea at the moment how I'll manage to find these in other files
            // At least I should be able to look for "REEL 1", "REEL 2" and so on.
        }

        private Excel.Worksheet importFile(String fileName, String index, ReelDataType type = ReelDataType.SHFL)
        {
            String prefix = "";
            if (type == ReelDataType.SHFL)
            {
                prefix = preParseFile(fileName);
                if (prefix != "")
                    prefix += underscore;
            }
            // clean up the file name to use as the worksheet name
            String[] trimmedName = fileName.Split(doubleBackSlash);
            int end = trimmedName.Length;
            String sheetName = trimmedName[end - 1];

            // create a new worksheet for our reel data
            newSheet = createSheet(trimName(sheetName));
            StreamReader inputFile = new StreamReader(fileName);

            // read reels from file
            if ( type == ReelDataType.SHFL )
                newSheet = readEquinoxHeader(prefix, inputFile, newSheet);

            // get this baby out from under foot - move it to the end of the workbook
            moveSheetToEnd(newSheet);
            return newSheet;
        }

        private Excel.Worksheet readEquinoxHeader(String prefix, StreamReader inStream, Excel.Worksheet sheet)
        {
            using (inStream)
            {
                int row = 1;
                String column = "A";
                String cellValue = "";
                bool addCell = false;
                while ((line = inStream.ReadLine()) != null)
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
                        if (cellValue == closeBrace)
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
                                if (prefix != "")
                                    sheet.Cells.Range[currentCell, Type.Missing].Value2 = cellValue.Replace(prefix, "");
                                else
                                    sheet.Cells.Range[currentCell, Type.Missing].Value2 = cellValue;
                            }
                            catch (Exception e)
                            {
                                System.Windows.Forms.MessageBox.Show("Error code:\n" + e.Message, "File Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }

            return sheet;
        }

        private void importBallyFile(String fileName, String index)
        {
            // clean up the file name to use as the worksheet name
            String[] trimmedName = fileName.Split(doubleBackSlash);
            int end = trimmedName.Length;
            String sheetName = trimmedName[end - 1];

            StreamReader inputFile = new StreamReader(fileName);
            m_gamePays.Parse(inputFile);

            inputFile.Close();

            inputFile = new StreamReader(fileName);
            m_gameSet.Parse(inputFile);
            if (m_gameSet.IsValid)
                m_gameSet.SendToWorksheet(sheetName, target);
        }

        private void cleanStringArray(String[] stringList)
        {
            for( int i = 0; i < stringList.Length; i++)
            {
                stringList[i] = stringList[i].Trim();
            }
        }

        private String preParseFile(String fileName)
        {
            // clean up the file name to use as the worksheet name
            System.Collections.Generic.List<String> cells = new List<String>();
            Array tempVal;
            // read in the reel data
            using (StreamReader inputFile = new StreamReader(fileName))
            {
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
                        if (cellValue == endOfReel)
                        {
                            addCell = false;
                            continue;
                        }
                        if (cellValue == closeBrace)
                        {
                            addCell = false;
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
                            cells.Add(cellValue);
                        }
                    }
                }
            }

            String first = "";
            System.Collections.Generic.List<KeyValue> foundList = new List<KeyValue>();
            bool found = false;
            int foundCount = foundList.Count;
            
            for (int j = 0; j < cells.Count; j++)
            {
                tempVal = cells[j].Split('_');
                first = tempVal.GetValue(0).ToString();
                found = false;
                foundCount = foundList.Count;
                for (int k = 0; k < foundCount; k++)
                {
                    if (foundList[k].Key == first)
                    {
                        found = true;
                        foundList[k].Value++;
                    }
                }
                if (!found)
                {
                    KeyValue item = new KeyValue();
                    item.Key = first;
                    item.Value = 1;
                    foundList.Add(item);
                }
            }
            int highest = 0;
            int highCount = 0;
            for (int a = 0; a < foundList.Count; a++)
            {
                if (foundList[a].Value > highCount)
                {
                    highCount = foundList[a].Value;
                    highest = a;
                }
            }
            if (highCount > 0)
            {
                double ratio = System.Convert.ToDouble(highCount) / System.Convert.ToDouble(cells.Count);
                if (ratio > 0.9)
                    return foundList[highest].Key;
            }
            return "";
        }

        private String trimName(String name)
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
            String name = trimName(fileName) + " Match";
            if (findSheet(name))
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
                    moveSheetToEnd(newWorksheet);
                    return;
                }
            }
        }

        private void copyPaySheet(String fileName)
        {
            String name = trimName(fileName) + " Pays";
            if (findSheet(name))
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

        private int getSheetIndex(String sheetName)
        {
            for (int i = 1; i <= target.Sheets.Count; i++)
            {
                if (target.Worksheets[i].Name == sheetName)
                    return i;
            }
            return 0;
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

        private int parseInteger(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[^\d]");
            int value = -1;
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

        private int getFirstIntegerPosition(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[^\d]");
            int index = 0;
            int value = 0;
            String temp = data.Substring(index, 1);
            for (int i = 0; i < data.Length; i++)
            {
                try
                {
                    value = System.Convert.ToInt32(digits.Replace(temp, ""));
                }
                catch (Exception e)
                {
                    value = 0;
                }
                if(value > 0)
                {
                    index = i;
                    break;
                }
                temp = data.Substring(i, 1);
            }
            return index;
        }
    }
}
