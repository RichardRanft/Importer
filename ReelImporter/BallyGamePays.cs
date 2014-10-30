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
    public class BallyGamePays
    {
        private BallyLinePayData m_linePays;
        private BallyScatterPayData m_scatterPays;
        private BallyFreeLinePayData m_freeLinePays;
        private BallyFreeScatterPayData m_freeScatterPays;
        private String m_symbolList;
        
        private PayParserState m_parseState;
        private Utils m_util;

        public BallyGamePays()
        {
            m_linePays = new BallyLinePayData();
            m_scatterPays = new BallyScatterPayData();
            m_freeLinePays = new BallyFreeLinePayData();
            m_freeScatterPays = new BallyFreeScatterPayData();

            m_parseState = new PayParserState();
            m_util = new Utils();
        }

        private List<BallyPayData> getSubSets(BallyPayData set)
        {
            List<BallyPayData> temp = new List<BallyPayData>();

            return temp;
        }

        public void Parse(StreamReader inStream)
        {
            if (m_parseState == null)
                m_parseState = new PayParserState();
            
            String line = "";

            using (inStream)
            {
                while (line != null)
                {
                    try
                    {
                        line = inStream.ReadLine();
                    }
                    catch(ObjectDisposedException ex)
                    {
                        break;
                    }
                    // strip comments
                    if (line.Contains("/"))
                    {
                        int pos = line.IndexOf("/");
                        line = line.Remove(pos);
                    }

                    line = line.Trim();

                    if (line.Length == 0 || line == "")
                        continue;

                    // look for symbols
                    if (line == "symbols")
                    {
                        m_parseState.EnterSymbols();
                        m_linePays.Parse(inStream, line, m_parseState);
                    }
                    
                    // look for pays
                    if (line == "linepays")
                    {
                        m_parseState.EnterLinePay();
                        m_linePays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "freegame_linepays")
                    {
                        m_parseState.EnterFreeLinePay();
                        m_freeLinePays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "scatterpays")
                    {
                        m_parseState.EnterScatterPay();
                        m_scatterPays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "freegame_scatterpays")
                    {
                        m_parseState.EnterFreeScatterPay();
                        m_freeScatterPays.Parse(inStream, line, m_parseState);
                    }
                    if (line == "featuredefs")
                    {
                        break;
                    }
                }
            }
        }

        public void exportPays(String sheetName, Excel.Workbook targetBook)
        {

        }
    }
}
