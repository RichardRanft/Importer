using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReelImporter
{
    class PaylineDescription
    {
        private List<String> m_stopValues;
        private int m_betMultiplier;
        private int m_group;
        private int m_id;
        private int m_win;
        private List<String> m_flags;

        public List<String> StopValues
        {
            get
            {
                return m_stopValues;
            }
        }

        public int Multiplier
        {
            get
            {
                return m_betMultiplier;
            }
        }

        public int Group
        {
            get
            {
                return m_group;
            }
        }

        public int ID
        {
            get
            {
                return m_id;
            }
        }

        public int Win
        {
            get
            {
                return m_win;
            }
        }

        public List<String> Flags
        {
            get
            {
                return m_flags;
            }
        }

        // {{CWC,CWC,CWC,CWC,CWC},	flags={ PRECOG_SET },xbet=1,		group = 100,	id = 10,	win = 0   }
        // {{ XX,XX,XX,XX,XX }, { MCAA,MCAA,MCAA, XX, XX },  xbet=0, group = 100, id = 730, win = 50,    flags = {WAY_WIN,ADJ_SCATTER}  },
        public void Add(String payline, Utils util)
        {
            // parse the line and store the data.
            payline.Trim();
            List<int> openBraceLoc = new List<int>();
            List<int> closeBraceLoc = new List<int>();
            int position = 0;
            foreach (Char ch in payline)
            {
                if (ch == util.openBrace[0])
                    openBraceLoc.Add(position);
                if (ch == util.closeBrace[0])
                    closeBraceLoc.Add(position);
            }
            int start, length = 0;
            String temp;
            String[] parts;
            for(int i = 0; i < (openBraceLoc.Count - 1); i++)
            {
                start = openBraceLoc[i + 1];
                length = closeBraceLoc[i] - start;
                temp = payline.Substring(start, length);
                payline = payline.Replace(temp, "");
                payline.Trim();
                parts = temp.Split(util.comma);
            }
        }

        public void Clear()
        {
            m_stopValues.Clear();
        }
    }
}
