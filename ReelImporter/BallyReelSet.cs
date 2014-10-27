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
    public class BallyReelSet : ReelSet
    {
        private Excel.Window excelWin;
        private Excel.Workbook targetBook;

        private List<BallyReel> m_reels;

        new public int Count
        {
            get
            {
                return m_reels.Count;
            }
        }

        new public List<BallyReel> Reels
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

        public BallyReelSet()
        {
            m_reels = new List<BallyReel>();
            m_setName = "";
        }

        public BallyReelSet(List<BallyReel> reelSet)
        {
            m_reels = reelSet;
            m_setName = "";
        }

        public void AddReel(BallyReel reel)
        {
            m_reels.Add(reel);
        }

        new public void Clear()
        {
            m_reels.Clear();
        }

        public override void Clean()
        {
            foreach (BallyReel reel in m_reels)
            {
                reel.Clean();
            }
        }

        new public BallyReel Get(int index)
        {
            if (index < m_reels.Count)
            {
                return m_reels[index];
            }
            return null;
        }

        new public void SendToWorksheet(Excel.Worksheet targetSheet, String StartCell, bool skipColumns = false, bool fullOutput = false)
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
            foreach( BallyReel reel in m_reels )
            {
                reel.OutputColumn(targetSheet, cell, false);
                col = incrementColumn(col);
                cell = col + row.ToString();
                if (skipColumns)
                    cell = incrementColumn(col) + row.ToString();
            }
        }

        new public void SendToWorksheet(Excel.Worksheet targetSheet, String StartCell, int startReel, int reelWidth, bool skipColumns = false, bool fullOutput = false)
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
    }
}
