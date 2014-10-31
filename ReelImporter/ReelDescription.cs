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
            String temp = "";
            foreach ( String value in parts )
            {
                temp = value;
                if (temp == "XX")
                    temp = "-";
                m_reelValues.Add(temp);
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
    }
}
