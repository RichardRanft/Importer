using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


namespace ReelImporter
{
    public class BallyReel : Reel
    {
        private List<BallyReelStop> m_reelStops;
        private Utils m_utils;

        new public List<BallyReelStop> ReelStops
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

        public BallyReel()
        {
            m_utils = new Utils();
            m_reelStops = new List<BallyReelStop>();
        }

        public BallyReel(List<BallyReelStop> reelStops)
        {
            m_utils = new Utils();
            m_reelStops = reelStops;
        }

        public void Add(BallyReelStop stop)
        {
            m_reelStops.Add(stop);
        }

        public override void Parse(System.IO.StreamReader inStream, string line, ParserState parseState)
        {
            bool reelEndFound = false;
            String[] parsedRow;
            do
            {
                line = line.Trim();

                if (line.Length == 0)
                    continue;

                if (line.Contains("/"))
                {
                    int pos = line.IndexOf("/");
                    line = line.Remove(pos);
                }

                if (line == m_utils.closeBrace || line == m_utils.arrayEnd)
                    reelEndFound = true;

                if (reelEndFound)
                {
                    parseState.LeaveArrayLevel();
                    if (parseState.ReelStart)
                        parseState.LeaveBaseReel();
                    if (parseState.FreeStart)
                        parseState.LeaveFreeReel();
                    if (parseState.ModifierStart)
                        parseState.LeaveModifierReel();

                    return;
                }

                if (!line.StartsWith(m_utils.openBrace))
                    continue;

                // parse the line
                line = line.Replace("{", ""); // kill open braces
                line = line.Replace(" ", ""); // kill extra spaces
                line = line.Replace(",", " "); // kill commas
                parsedRow = line.Split(m_utils.cBrace, StringSplitOptions.RemoveEmptyEntries); // now we can split into space-separated tuples
                cleanStringArray(parsedRow);

                for (int i = 0; i < parsedRow.Length; i++)
                {
                    if (parsedRow[i] == "")
                        continue;
                    String[] parsedCell = parsedRow[i].Split(' ');
                    if (parsedCell.Length != 3)
                        continue;
                    BallyReelStop stop = new BallyReelStop();
                    stop.Parse(parsedRow[i]);
                    if (stop.IsValid)
                        Add(stop);
                }
            } while ((line = inStream.ReadLine()) != null);

            if (parseState.ReelStart)
                parseState.LeaveBaseReel();
            if (parseState.FreeStart)
                parseState.LeaveFreeReel();
            if (parseState.ModifierStart)
                parseState.LeaveModifierReel();
        }

        public override void Clean()
        {
            foreach (BallyReelStop stop in m_reelStops)
            {
                String data;
                data = stop.Value.Item1;
                if (data.Length > 2)
                    stop.Clean();
            }
        }

        new public BallyReelStop Get(int index)
        {
            if (index < m_reelStops.Count)
            {
                return m_reelStops[index];
            }
            return null;
        }

        new public void OutputColumn(Excel.Worksheet targetSheet, String startCell, bool fullOutput = false)
        {
            String col = parseCol(startCell);
            int row = parseRow(startCell);
            foreach( BallyReelStop stop in m_reelStops )
            {
                stop.OutputCell(targetSheet, col + row.ToString(), fullOutput);
                row++;
            }
        }
    }
}
