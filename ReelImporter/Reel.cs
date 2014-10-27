using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public abstract class Reel
    {
        private List<ReelStop> m_reelStops;

        private Utils m_utils;

        public virtual List<ReelStop> ReelStops
        {
            get
            {
                return m_reelStops;
            }
            set
            {
                m_reelStops = value;
            }
        }

        public Reel()
        {
            m_reelStops = new List<ReelStop>();
        }

        public  Reel(List<ReelStop> reelStops)
        {
            m_reelStops = reelStops;
        }

        public virtual void Add(ReelStop stop)
        {
            m_reelStops.Add(stop);
        }

        public virtual ReelStop Get(int index)
        {
            if (index < m_reelStops.Count)
            {
                return m_reelStops[index];
            }
            return null;
        }

        public virtual void Clean()
        {
            foreach (ReelStop stop in m_reelStops)
            {
                String data;
                data = stop.Value.Item1;
                if (data.Length > 2)
                    stop.Clean();
            }
        }

        public abstract void Parse(StreamReader inStream, String line, ParserState parseState);

        public virtual void OutputColumn(Excel.Worksheet targetSheet, String startCell, bool fullOutput = false)
        {
            String col = parseCol(startCell);
            int row = parseRow(startCell);
            foreach( ReelStop stop in m_reelStops )
            {
                stop.OutputCell(targetSheet, col + row.ToString(), fullOutput);
                row++;
            }
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

        protected void cleanStringArray(String[] stringList)
        {
            for (int i = 0; i < stringList.Length; i++)
            {
                stringList[i] = stringList[i].Trim();
            }
        }
    }
}
