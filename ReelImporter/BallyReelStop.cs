using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


namespace ReelImporter
{
    public class BallyReelStop : ReelStop
    {
        private Tuple<String, int, int> m_stopValue;
        private bool m_isValid;

        public BallyReelStop()
        {
            m_stopValue = new Tuple<string, int, int>("", 0, 0);
            m_isValid = true;
        }

        public BallyReelStop(String Name, int Weight, int Nudge)
        {
            m_stopValue = new Tuple<string, int, int>(Name, Weight, Nudge);
            if (m_stopValue != null)
                m_isValid = true;
        }

        public override bool IsValid
        {
            get
            {
                return m_isValid;
            }
        }

        new public Tuple<String, int, int> Value
        {
            get
            {
                return m_stopValue;
            }
            set
            {
                m_stopValue = value;
                if (m_stopValue != null)
                    m_isValid = true;
            }
        }

        public override void Clean()
        {
            String temp = "";
            String data = m_stopValue.Item1;
            if (data.Length > 2)
            {
                if (getEndsWithInteger(data))
                    temp = data.Substring(data.Length - 3, 3);
                else
                    temp = data.Substring(data.Length - 2, 2);
            }
            m_stopValue = new Tuple<string, int, int>(temp, m_stopValue.Item2, m_stopValue.Item3);
        }

        public override void Parse(String reelEntry)
        {
            reelEntry = reelEntry.Replace("{", "");
            reelEntry = reelEntry.Replace("}", "");

            String[] reelStop = reelEntry.Split(' ');
            String name = reelStop[0];
            int weight = 0;
            int nudge = 0;
            try
            {
                weight = Convert.ToInt32(reelStop[1]);
            }
            catch (System.FormatException e)
            {
                m_isValid = false;
                return;
            }
            try
            {
                nudge = Convert.ToInt32(reelStop[2]);
            }
            catch (System.FormatException e)
            {
                m_isValid = false;
                return;
            }
            m_stopValue = new Tuple<string, int, int>(name, weight, nudge);
            if (m_stopValue != null)
                m_isValid = true;
        }

        new public bool OutputCell(Excel.Worksheet targetSheet, String cell, bool fullOutput = false)
        {
            bool result = true;
            int row = parseRow(cell);
            if (row < 1)
                row = 1;
            try
            {
                // try, because it might fail
                if (fullOutput)
                    targetSheet.Cells.Range[cell, Type.Missing].Value2 = m_stopValue.Item1 + " " + m_stopValue.Item2 + " " + m_stopValue.Item3;
                else
                    targetSheet.Cells.Range[cell, Type.Missing].Value2 = m_stopValue.Item1;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + e.Message, "File Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                result = false;
            }
            return result;
        }
    }
}
