using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public abstract class ReelStop
    {
        private Tuple<String, int, int> m_stopValue;

        public ReelStop()
        {
            m_stopValue = new Tuple<string, int, int>("", 0, 0);
        }

        public ReelStop(String Name, int Weight, int Nudge)
        {
            m_stopValue = new Tuple<string, int, int>(Name, Weight, Nudge);
        }

        public virtual Tuple<String, int, int> Value
        {
            get
            {
                return m_stopValue;
            }
            set
            {
                m_stopValue = value;
            }
        }

        public abstract bool IsValid
        {
            get;
        }

        public virtual void Clean()
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

        public abstract void Parse(String reelEntry);

        public virtual bool OutputCell(Excel.Worksheet targetSheet, String cell, bool fullOutput = false)
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

        public String parseCol(String data)
        {
            System.Text.RegularExpressions.Regex digits = new System.Text.RegularExpressions.Regex(@"[\d]");
            return digits.Replace(data, "");
        }

        public int parseRow(String data)
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

        public bool getEndsWithInteger(String data)
        {
            return (Char.IsNumber(data.ToCharArray()[data.Length - 1]));
        }
    }
}
