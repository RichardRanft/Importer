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
            foreach ( String value in parts )
            {
                m_reelValues.Add(value);
            }
        }
    }
}
