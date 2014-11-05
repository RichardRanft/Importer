using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReelImporter
{
    public class PaylineDescription
    {
        private List<ReelDescription> m_stopValues;
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
        private int m_feature;
        private List<String> m_flags;

        private bool m_isValid;

        public List<ReelDescription> StopValues
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

        public bool IsModifierSet
        {
            get
            {
                int set = m_stopValues.Count - 1;
                return m_stopValues[set].IsModifierReel();
            }
        }

        public bool HasWild
        {
            get
            {
                int set = m_stopValues.Count - 1;
                return m_stopValues[set].HasWild();
            }
        }

        public bool IsFreegameSet
        {
            get
            {
                int set = m_stopValues.Count - 1;
                return m_stopValues[set].IsFreegameReel();
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
            m_stopValues = new List<ReelDescription>();
            m_betMultiplier = -1;
            m_group = -1;
            m_id = -1;
            m_win = -1;
            m_winLevel = -1;
            m_minBet = -1;
            m_maxBet = -1;
            m_minLines = -1;
            m_maxLines = -1;
            m_validPayLines = -1;
            m_feature = -1;
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

            // find our open and close braces
            foreach (Char ch in payline)
            {
                if (ch == '=')
                    break; // we're past the reel description, move on
                if (ch == util.openBrace[0])
                    openBraceLoc.Add(position);
                if (ch == util.closeBrace[0])
                    closeBraceLoc.Add(position);
                position++;
            }
            int startSub = 0, start = 0, length = 0;
            String temp = "";
            ReelDescription desc;
            // grab brace-enclosed sets - these are reel descriptions
            for(int i = 0; i < (openBraceLoc.Count - 1); i++)
            {
                startSub = temp.Length;
                start = openBraceLoc[i + 1] - startSub;
                length = closeBraceLoc[i] - start - startSub;
                temp = payline.Substring(start, length);
                payline = payline.Replace(temp, "");
                payline.Trim();
                desc = new ReelDescription();
                desc.Parse(temp, util);
                m_stopValues.Add(desc);
            }
            // now grab the payline data
            // first, comma-separated chunks
            payline = payline.Replace("\t", "");
            String[] entries = payline.Split(util.comma);
            String[] parts;
            foreach(String entry in entries)
            {
                parts = entry.Split('=');
                if (parts.Length < 2)
                    continue;
                addDataEntry(parts[0], parts[1], util);
                if (!m_isValid)
                    break;
            }
        }

        private void addDataEntry(String field, String value, Utils util)
        {
            field = field.Replace('{', ' ');
            field = field.Replace('}', ' ');
            field = field.Trim();

            value = value.Replace('{', ' ');
            value = value.Replace('}', ' ');
            value = value.Trim();
            switch (field)
            {
                case "xbet":
                    {
                        try
                        {
                            m_betMultiplier = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_betMultiplier = -1;
                        }
                        break;
                    }
                case "feature":
                    {
                        try
                        {
                            m_feature = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_feature = -1;
                        }
                        break;
                    }
                case "group":
                    {
                        try
                        {
                            m_group = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_group = -1;
                        }
                        break;
                    }
                case "flags":
                    {
                        value = value.Replace(util.openBrace, "");
                        value = value.Replace(util.closeBrace, "");
                        String[] flags = value.Split(util.comma);

                        foreach (String flag in flags)
                        {
                            m_flags.Add(flag);
                        }
                        
                        m_isValid = true;
                        break;
                    }
                case "validpaylines":
                    {
                        try
                        {
                            m_validPayLines = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_validPayLines = -1;
                        }
                        break;
                    }
                case "win":
                    {
                        try
                        {
                            m_win = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_win = -1;
                        }
                        break;
                    }
                case "maxpaylines":
                    {
                        try
                        {
                            m_maxLines = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_maxLines = -1;
                        }
                        break;
                    }
                case "minpaylines":
                    {
                        try
                        {
                            m_minLines = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_minLines = -1;
                        }
                        break;
                    }
                case "maxbet":
                    {
                        try
                        {
                            m_maxBet = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_maxBet = -1;
                        }
                        break;
                    }
                case "minbet":
                    {
                        try
                        {
                            m_minBet = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_minBet = -1;
                        }
                        break;
                    }
                case "winlevel":
                    {
                        try
                        {
                            m_winLevel = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_winLevel = -1;
                        }
                        break;
                    }
                case "id":
                    {
                        try
                        {
                            m_id = Convert.ToInt32(value);
                            m_isValid = true;
                        }
                        catch
                        {
                            m_isValid = false;
                            m_id = -1;
                        }
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
        }

        public void Clear()
        {
            m_stopValues.Clear();
            m_betMultiplier = -1;
            m_group = -1;
            m_id = -1;
            m_win = -1;
            m_winLevel = -1;
            m_minBet = -1;
            m_maxBet = -1;
            m_minLines = -1;
            m_maxLines = -1;
            m_validPayLines = -1;
            m_flags.Clear();
            m_isValid = false;
        }

        public static bool operator <(PaylineDescription pd1, PaylineDescription pd2)
        {
            bool reelComp = (pd1.StopValues[pd1.StopValues.Count - 1] == pd2.StopValues[pd2.StopValues.Count - 1]);
            if (reelComp)
                return false;
            reelComp = (pd1.StopValues[pd1.StopValues.Count - 1] > pd2.StopValues[pd2.StopValues.Count - 1]);
            if (reelComp)
                return false;
            return false;
        }

        public static bool operator >(PaylineDescription pd1, PaylineDescription pd2)
        {
            return false;
        }

        public static bool operator ==(PaylineDescription pd1, PaylineDescription pd2)
        {
            return false;
        }

        public static bool operator !=(PaylineDescription pd1, PaylineDescription pd2)
        {
            return false;
        }

        public override bool Equals(Object o)
        {
            try
            {
                return (bool)(this == (PaylineDescription)o);
            }
            catch
            {
                return false;
            }
        }

        public override int GetHashCode()
        {
            int numVal = 0;
            String temp = ToString();
            foreach (Char c in temp.ToCharArray())
            {
                numVal += c.GetHashCode();
            }
            return numVal;
        }

        public override string ToString()
        {
            string temp = "";
            foreach(ReelDescription desc in m_stopValues)
            {
                temp += desc.ToString() + "\t";
            }

            temp += m_betMultiplier.ToString() + "\t";
            temp += m_group.ToString() + "\t";
            temp += m_id.ToString() + "\t";
            temp += m_win.ToString() + "\t";
            temp += m_winLevel.ToString() + "\t";
            temp += m_minBet.ToString() + "\t";
            temp += m_maxBet.ToString() + "\t";
            temp += m_minLines.ToString() + "\t";
            temp += m_maxLines.ToString() + "\t";
            temp += m_validPayLines.ToString() + "\t";
            temp += m_flags.ToString() + "\t";
            temp += m_isValid.ToString();

            return temp;
        }
    }
}
