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
            m_win = 0;
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

        private void findBraces(String data, Utils util, out List<int> openBraceList, out List<int> closeBraceList)
        {
            openBraceList = new List<int>();
            closeBraceList = new List<int>();
            int position = 0;

            // find our open and close braces
            foreach (Char ch in data)
            {
                if (ch == '=')
                    break; // we're past the reel description, move on
                if (ch == util.openBrace[0])
                    openBraceList.Add(position);
                if (ch == util.closeBrace[0])
                    closeBraceList.Add(position);
                position++;
            }
        }

        private String getReelSets(String data, Utils util)
        {
            List<int> openBraceLoc;
            List<int> closeBraceLoc;

            findBraces(data, util, out openBraceLoc, out closeBraceLoc);

            int startSub = 0, start = 0, length = 0;
            String temp = "";
            ReelDescription desc;
            // grab brace-enclosed sets - these are reel descriptions
            for (int i = 0; i < (openBraceLoc.Count - 1); i++)
            {
                startSub = temp.Length;
                start = openBraceLoc[i + 1] - startSub;
                length = closeBraceLoc[i] - start - startSub;
                temp = data.Substring(start, length);
                data = data.Remove(start, length);
                data.Trim();
                desc = new ReelDescription();
                desc.Parse(temp, util);
                m_stopValues.Add(desc);
            }
            data = data.Replace("{},", "");
            data = data.Replace("}}", "");
            return data;
        }

        // {{CWC,CWC,CWC,CWC,CWC},	flags={ PRECOG_SET },xbet=1,		group = 100,	id = 10,	win = 0   }
        // {{ XX,XX,XX,XX,XX }, { MCAA,MCAA,MCAA, XX, XX },  xbet=0, group = 100, id = 730, win = 50,    flags = {WAY_WIN,ADJ_SCATTER}  },
        public void Add(String payline, Utils util)
        {
            // parse the line and store the data.
            payline.Trim();
            payline = payline.Replace("\t", "");
            payline = payline.Replace(" ", "");

            payline = getReelSets(payline, util);

            m_flags = findFlags(payline, util);

            // now grab the payline data
            // comma-separated chunks
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

        private void parseFlags(String value, Utils util)
        {
            value = value.Replace(util.openBrace, "");
            value = value.Replace(util.closeBrace, "");
            String[] flags = value.Split(util.comma);

            foreach (String flag in flags)
            {
                if(!m_flags.Contains(flag))
                    m_flags.Add(flag);
            }
        }

        private List<String> findFlags(String data, Utils util)
        {
            List<String> flagList = new List<String>();
            String[] parts = data.Split('=');
            String[] bits;
            String temp;
            bool flagFound = false;
            foreach (String part in parts)
            {
                if(flagFound)
                {
                    bits = part.Split(',');
                    foreach (String bit in bits)
                    {
                        temp = bit.Replace(util.openBrace, "");
                        temp = temp.Replace(util.closeBrace, "");
                        if (checkNextField(temp))
                            break;
                        if (!flagList.Contains(temp) && temp != "")
                            flagList.Add(temp);
                    }
                    flagFound = false;
                }
                if (part.Contains("flags"))
                {
                    flagFound = true;
                }
            }

            return flagList;
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

        private bool checkNextField(String field)
        {
            field = field.Replace('{', ' ');
            field = field.Replace('}', ' ');
            field = field.Trim();
            bool isField = false;
            switch (field)
            {
                case "xbet":
                    {
                        isField = true;
                        break;
                    }
                case "feature":
                    {
                        isField = true;
                        break;
                    }
                case "group":
                    {
                        isField = true;
                        break;
                    }
                case "validpaylines":
                    {
                        isField = true;
                        break;
                    }
                case "win":
                    {
                        isField = true;
                        break;
                    }
                case "maxpaylines":
                    {
                        isField = true;
                        break;
                    }
                case "minpaylines":
                    {
                        isField = true;
                        break;
                    }
                case "maxbet":
                    {
                        isField = true;
                        break;
                    }
                case "minbet":
                    {
                        isField = true;
                        break;
                    }
                case "winlevel":
                    {
                        isField = true;
                        break;
                    }
                case "id":
                    {
                        isField = true;
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
            return isField;
        }

        public void Clear()
        {
            m_stopValues.Clear();
            m_betMultiplier = -1;
            m_group = -1;
            m_id = -1;
            m_win = 0;
            m_winLevel = -1;
            m_minBet = -1;
            m_maxBet = -1;
            m_minLines = -1;
            m_maxLines = -1;
            m_validPayLines = -1;
            m_feature = -1;
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

            temp += "xbet=" + m_betMultiplier.ToString() + "\t";
            temp += "group=" + m_group.ToString() + "\t";
            temp += "id=" + m_id.ToString() + "\t";
            temp += "win=" + m_win.ToString() + "\t";
            temp += "winlevel=" + m_winLevel.ToString() + "\t";
            temp += "minbet=" + m_minBet.ToString() + "\t";
            temp += "maxbet=" + m_maxBet.ToString() + "\t";
            temp += "minpaylines=" + m_minLines.ToString() + "\t";
            temp += "maxpaylines=" + m_maxLines.ToString() + "\t";
            temp += "validpaylines" + m_validPayLines.ToString() + "\t";
            temp += "feature" + m_feature.ToString() + "\t";
            temp += "flags=";
            foreach (String s in m_flags)
            {
                temp += s + ",";
            }
            temp += "\tSet is valid: " + m_isValid.ToString();

            return temp;
        }
    }
}
