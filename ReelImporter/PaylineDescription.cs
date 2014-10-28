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
        private int m_winLevel;
        private int m_minBet;
        private int m_maxBet;
        private int m_minLines;
        private int m_maxLines;
        private int m_validPayLines;
        private List<String> m_flags;

        private bool m_isValid;

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

        public bool IsValid
        {
            get
            {
                return m_isValid;
            }
        }

        public List<String> Flags
        {
            get
            {
                return m_flags;
            }
        }

        public PaylineDescription()
        {
            m_stopValues = new List<String>();
            m_betMultiplier = 0;
            m_group = 0;
            m_id = 0;
            m_win = 0;
            m_winLevel = 0;
            m_minBet = 0;
            m_maxBet = 0;
            m_minLines = 0;
            m_maxLines = 0;
            m_validPayLines = 0;
            m_flags = new List<String>();
            m_isValid = false;
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
