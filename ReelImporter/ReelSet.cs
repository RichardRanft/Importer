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
    public abstract class ReelSet
    {
        private Excel.Window excelWin;
        private Excel.Workbook targetBook;

        private List<Reel> m_reels;
        protected String m_setName;

        public virtual int Count
        {
            get
            {
                return m_reels.Count;
            }
        }

        public virtual List<Reel> Reels
        {
            get
            {
                return m_reels;
            }
            set
            {
                m_reels = value;
            }
        }

        public String Name
        {
            get
            {
                return m_setName;
            }
            set
            {
                m_setName = value;
            }
        }

        public ReelSet()
        {
            m_reels = new List<Reel>();
        }

        public ReelSet(List<Reel> reelSet)
        {
            m_reels = reelSet;
        }

        public virtual void AddReel(Reel reel)
        {
            m_reels.Add(reel);
        }

        public virtual void Clear()
        {
            m_reels.Clear();
        }

        public virtual Reel Get(int index)
        {
            if (index < m_reels.Count)
            {
                return m_reels[index];
            }
            return null;
        }

        public virtual void Clean()
        {
            foreach (Reel reel in m_reels)
            {
                reel.Clean();
            }
        }

        public virtual void SendToWorksheet(Excel.Worksheet targetSheet, String StartCell, bool skipColumns = false, bool fullOutput = false)
        {
            // Fill in a set of cells with this reelset's values
            // The skipColumns parameter will skip columns between reels, starting in the target cell.
            // The fullOutput parameter will include the weight and nudge values in separate columns in the output.  If 
            // skipColumns is specified a column will be skipped between reels.

            // stop screen updates - reduces run time by nearly a factor of 10
            Globals.Program.Application.ScreenUpdating = false;
            if (excelWin == null)
                excelWin = Globals.Program.Application.ActiveWindow;

            targetBook = excelWin.Application.ActiveWorkbook;
            if (targetBook == null)
            {
                System.Windows.Forms.MessageBox.Show("No source Workbook or Worksheet available.", "Error - No source.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            String cell = StartCell;
            String col = parseCol(StartCell);
            int row = parseRow(StartCell);
            foreach( Reel reel in m_reels )
            {
                reel.OutputColumn(targetSheet, cell, false);
                col = incrementColumn(col);
                cell = col + row.ToString();
                if (skipColumns)
                    cell = incrementColumn(col) + row.ToString();
            }
        }

        public virtual void SendToWorksheet(Excel.Worksheet targetSheet, String StartCell, int startReel, int reelWidth, bool skipColumns = false, bool fullOutput = false)
        {
            // Fill in a set of cells with this reelset's values
            // The skipColumns parameter will skip columns between reels, starting in the target cell.
            // The fullOutput parameter will include the weight and nudge values in separate columns in the output.  If 
            // skipColumns is specified a column will be skipped between reels.

            if (startReel < 0 || (startReel + reelWidth) >= m_reels.Count)
            {
                MessageBox.Show("SendToWorksheet() - invalid range specified");
                return;
            }

            // stop screen updates - reduces run time by nearly a factor of 10
            Globals.Program.Application.ScreenUpdating = false;
            if (excelWin == null)
                excelWin = Globals.Program.Application.ActiveWindow;

            targetBook = excelWin.Application.ActiveWorkbook;
            if (targetBook == null)
            {
                System.Windows.Forms.MessageBox.Show("No source Workbook or Worksheet available.", "Error - No source.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            String col = parseCol(StartCell);
            int row = parseRow(StartCell);
            if (row < 1)
                row = 1;
            String cell = col + row.ToString();
            for (int i = startReel; i < (startReel + reelWidth); i++)
            {
                m_reels[i].OutputColumn(targetSheet, cell, false);
                cell = incrementColumn(cell);
                if (skipColumns)
                    cell = incrementColumn(cell);
            }
            Globals.Program.Application.ScreenUpdating = true;
        }

        public virtual String incrementColumn(String current)
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
                nextColumn = nextColumn[0] + temp.ToString();
            }
            return nextColumn;
        }

        public virtual String parseCol(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[\d]");
            return digits.Replace(data, "");
        }

        public virtual int parseRow(String data)
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
