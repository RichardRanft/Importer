using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReelImporter
{
    public class ReelDescription
    {
        private List<String> m_reelValues;

        public List<String> Values
        {
            get
            {
                return m_reelValues;
            }
        }

        public int Hits
        {
            get
            {
                return getHits();
            }
        }

        public ReelDescription()
        {
            m_reelValues = new List<String>();
        }

        public void Parse(String input, Utils util)
        {
            input = input.Replace('{', ' ');
            input = input.Replace('}', ' ');
            input = input.Trim();

            String[] parts = input.Split(util.comma, 15, StringSplitOptions.RemoveEmptyEntries);
            foreach ( String value in parts )
            {
                m_reelValues.Add(value);
            }
        }

        public bool IsModifierReel()
        {
            bool isMod = false;
            foreach (String stop in m_reelValues)
            {
                if(stop.Contains("MC"))
                {
                    isMod = true;
                    break;
                }
            }

            return isMod;
        }

        public bool HasWild()
        {
            bool isMod = false;
            foreach (String stop in m_reelValues)
            {
                if (stop.Contains("WC"))
                {
                    isMod = true;
                    break;
                }
            }

            return isMod;
        }

        public bool IsFreegameReel()
        {
            bool isFG = false;
            foreach (String stop in m_reelValues)
            {
                if (stop.Contains("FG"))
                {
                    isFG = true;
                    break;
                }
            }

            return isFG;
        }

        public String TrimValue(String value)
        {
            String temp = value;
            if (value.Length > 2)
            {
                if (getEndsWithInteger(value))
                    temp = value.Substring(value.Length - 3, 3);
                else
                    temp = value.Substring(value.Length - 2, 2);
            }
            if (temp == "XX")
                temp = "-";
            return temp;
        }

        public bool getEndsWithInteger(String data)
        {
            return (Char.IsNumber(data.ToCharArray()[data.Length - 1]));
        }

        private int getHits()
        {
            int hits = 5;
            foreach (String entry in m_reelValues)
            {
                if (entry.Contains("XX") || entry.Contains("-"))
                    hits--;
            }
            return hits;
        }

        public static bool operator <(ReelDescription rd1, ReelDescription rd2)
        {
            int alphaRank = 0;
            String currFirst = "";
            String currSecond = "";
            for (int i = 0; i < rd1.Values.Count; i++)
            {
                currFirst = rd1.TrimValue(rd1.Values[i]);
                if (currFirst == "-" || currFirst == "XX")
                    currFirst = "~";
                currSecond = rd2.TrimValue(rd2.Values[i]);
                if (currSecond == "-" || currSecond == "XX")
                    currSecond = "~";
                alphaRank = String.Compare(currFirst, currSecond);
                if (alphaRank != 0)
                    break;
            }
            return (alphaRank < 0);
        }

        public static bool operator >(ReelDescription rd1, ReelDescription rd2)
        {
            int alphaRank = 0;
            String currFirst = "";
            String currSecond = "";
            for (int i = 0; i < rd1.Values.Count; i++)
            {
                currFirst = rd1.TrimValue(rd1.Values[i]);
                if (currFirst == "-" || currFirst == "XX")
                    currFirst = "~";
                currSecond = rd2.TrimValue(rd2.Values[i]);
                if (currSecond == "-" || currSecond == "XX")
                    currSecond = "~";
                alphaRank = String.Compare(currFirst, currSecond);
                if (alphaRank != 0)
                    break;
            }
            return (alphaRank > 0);
        }

        public static bool operator ==(ReelDescription rd1, ReelDescription rd2)
        {
            int alphaRank = 0;
            String currFirst = "";
            String currSecond = "";
            for (int i = 0; i < rd1.Values.Count; i++)
            {
                currFirst = rd1.TrimValue(rd1.Values[i]);
                if (currFirst == "-" || currFirst == "XX")
                    currFirst = "~";
                currSecond = rd2.TrimValue(rd2.Values[i]);
                if (currSecond == "-" || currSecond == "XX")
                    currSecond = "~";
                alphaRank = String.Compare(currFirst, currSecond);
                if (alphaRank != 0)
                    break;
            }
            return (alphaRank == 0);
        }

        public static bool operator !=(ReelDescription rd1, ReelDescription rd2)
        {
            int alphaRank = 0;
            String currFirst = "";
            String currSecond = "";
            for (int i = 0; i < rd1.Values.Count; i++)
            {
                currFirst = rd1.TrimValue(rd1.Values[i]);
                if (currFirst == "-" || currFirst == "XX")
                    currFirst = "~";
                currSecond = rd2.TrimValue(rd2.Values[i]);
                if (currSecond == "-" || currSecond == "XX")
                    currSecond = "~";
                alphaRank = String.Compare(currFirst, currSecond);
                if (alphaRank != 0)
                    break;
            }
            return (alphaRank != 0);
        }

        public override bool Equals(Object o)
        {
            try
            {
                return (bool)(this == (ReelDescription)o);
            }
            catch
            {
                return false;
            }
        }

        public override string ToString()
        {
            string temp = "";
            for ( int i = 0; i < m_reelValues.Count; i++ )
            {
                if (i == m_reelValues.Count - 1)
                    temp += m_reelValues[i];
                else
                    temp += m_reelValues[i] + "\t";
            }
            return temp;
        }
    }
}
