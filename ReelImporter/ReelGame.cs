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
    public abstract partial class ReelGame
    {
        protected ReelSet m_baseReelset;
        protected ReelSet m_freeReelset;
        protected ReelSet m_baseModReelset;
        protected ReelSet m_freeModReelset;
        protected ReelSet m_currentSet;

        protected int m_setIndex;
        protected int m_reelWidth;
        protected bool m_isValid;
        protected bool m_hasModifierReels;
        protected bool m_hasFreeReels;
        protected bool m_hasFreeModReels;

        private char underscore = '_';

        public virtual ReelSet BaseReels
        {
            get
            {
                return m_baseReelset;
            }
            set
            {
                m_baseReelset = value;
                m_reelWidth = m_baseReelset.Count;
                m_isValid = checkValid();
            }
        }

        public virtual ReelSet BaseModifierReels
        {
            get
            {
                return m_baseModReelset;
            }
            set
            {
                m_baseModReelset = value;
                m_hasModifierReels = true;
                m_isValid = checkValid();
            }
        }

        public virtual ReelSet FreeReels
        {
            get
            {
                return m_freeReelset;
            }
            set
            {
                m_freeReelset = value;
                m_hasFreeReels = true;
                m_isValid = checkValid();
            }
        }

        public virtual ReelSet FreeModifierReels
        {
            get
            {
                return m_freeModReelset;
            }
            set
            {
                m_freeModReelset = value;
                m_hasFreeModReels = true;
                m_isValid = checkValid();
            }
        }

        public virtual bool IsValid
        {
            get
            {
                return m_isValid;
            }
        }

        public virtual bool HasModifierReels
        {
            get
            {
                return m_hasModifierReels;
            }
        }

        public virtual bool HasFreegameReels
        {
            get
            {
                return m_hasFreeReels;
            }
        }

        public virtual bool HasFreegameModifierReels
        {
            get
            {
                return m_hasFreeModReels;
            }
        }

        public virtual void SendToWorksheet(String sheetName, Excel.Workbook targetBook)
        {
            int setIndex = 1;
            m_setIndex = 7;
            m_currentSet = m_baseReelset;
            m_currentSet.Clean();
            exportReels(sheetName + "base" + setIndex++.ToString(), targetBook);

            List<ReelSet> tempSets = null;
            if (m_baseModReelset.Count > 0 && m_baseModReelset.Count == m_reelWidth)
            {
                m_currentSet = m_baseModReelset;
                m_currentSet.Clean();
                exportReels(sheetName + "base_mod" + setIndex++.ToString(), targetBook);
            }
            else if (m_baseModReelset.Count > 0 && m_baseModReelset.Count > m_reelWidth)
            {
                tempSets = getSubSets(m_baseModReelset);
                if (tempSets != null)
                {
                    foreach (ReelSet set in tempSets)
                    {
                        m_currentSet = set;
                        m_currentSet.Clean();
                        exportReels(sheetName + "base_mod" + setIndex++.ToString(), targetBook);
                    }
                }
            }

            if (m_freeReelset.Count > 0 && m_freeReelset.Count == m_reelWidth)
            {
                m_currentSet = m_freeReelset;
                m_currentSet.Clean();
                exportReels(sheetName + "free" + setIndex++.ToString(), targetBook);
            }
            else if (m_freeReelset.Count > 0 && m_freeReelset.Count > m_reelWidth)
            {
                tempSets = getSubSets(m_freeReelset);
                if (tempSets != null)
                {
                    foreach (ReelSet set in tempSets)
                    {
                        m_currentSet = set;
                        m_currentSet.Clean();
                        exportReels(sheetName + "free" + setIndex++.ToString(), targetBook);
                    }
                }
            }

            if (m_freeModReelset.Count > 0 && m_freeModReelset.Count == m_reelWidth)
            {
                m_currentSet = m_freeModReelset;
                m_currentSet.Clean();
                exportReels(sheetName + "free_mod" + setIndex++.ToString(), targetBook);
            }
            else if (m_freeModReelset.Count > 0 && m_freeModReelset.Count > m_reelWidth)
            {
                tempSets = getSubSets(m_freeModReelset);
                if (tempSets != null)
                {
                    foreach (ReelSet set in tempSets)
                    {
                        m_currentSet = set;
                        m_currentSet.Clean();
                        exportReels(sheetName + "free_mod" + setIndex++.ToString(), targetBook);
                    }
                }
            }
        }

        protected abstract bool checkValid();

        protected abstract List<ReelSet> getSubSets(ReelSet set);

        public abstract void Parse(StreamReader inStream);

        protected abstract void exportReels(String sheetName, Excel.Workbook targetBook);

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

        public void updateMatchLinks(Excel.Worksheet sheet, Excel.Workbook target, String name, int setIndex)
        {
            // Notes:
            // Reel columns start at Q8
            // This also needs to update all links to point to the new target worksheet.
            String matchSheetName = trimName(name) + " Match";
            Excel.Worksheet matchSheet = null;
            int index = getSheetIndex("Game Info", target);
            Excel.Worksheet info = target.Worksheets[index];
            String link = "='" + matchSheetName + "'!$G$4";
            String nameCell = "B" + setIndex.ToString();
            String linkCell = "C" + setIndex.ToString();
            m_setIndex++;
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

        public int getFirstIntegerPosition(String data)
        {
            int index = 0;
            String temp;
            for (int i = 0; i < data.Length; i++)
            {
                temp = data.Substring(i, 1);
                if (Char.IsNumber(temp.ToCharArray()[0]))
                {
                    index = i + 1;
                    break;
                }
            }
            return index;
        }

        public bool getEndsWithInteger(String data)
        {
            return (Char.IsNumber(data.ToCharArray()[data.Length - 1]));
        }
    }
}
