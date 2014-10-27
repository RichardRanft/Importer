using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

namespace ReelImporter
{
    public class LineParseState
    {
        private BitArray m_currentReadState;
        private int m_arrayDepth;
        private int[] m_enteredState;

        public int[] StateEnteredLevel
        {
            get
            {
                return m_enteredState;
            }
        }

        public BitArray State
        {
            get
            {
                return m_currentReadState;
            }
        }

        public bool None
        {
            get
            {
                return m_currentReadState[(int)LineReadState.NONE];
            }
        }

        public bool DescriptionStart
        {
            get
            {
                return m_currentReadState[(int)LineReadState.DESC_START];
            }
        }

        public bool DescriptionEnd
        {
            get
            {
                return m_currentReadState[(int)LineReadState.DESC_END];
            }
        }

        public bool DataStart
        {
            get
            {
                return m_currentReadState[(int)LineReadState.DATA_START];
            }
        }

        public bool DataEnd
        {
            get
            {
                return m_currentReadState[(int)LineReadState.DATA_END];
            }
        }

        public int ArrayDepth
        {
            get
            {
                return m_arrayDepth;
            }
            set
            {
                m_arrayDepth = value;
            }
        }

        public LineParseState()
        {
            m_currentReadState = new BitArray(13);
            m_enteredState = new int[5];
            m_enteredState[(int)LineReadState.NONE] = 0;
            m_enteredState[(int)LineReadState.DESC_START] = 0;
            m_enteredState[(int)LineReadState.DESC_END] = 0;
            m_enteredState[(int)LineReadState.DATA_START] = 0;
            m_enteredState[(int)LineReadState.DATA_END] = 0;
            m_arrayDepth = 0;
        }

        public void EnterReelDesc()
        {
            m_enteredState[(int)LineReadState.DESC_START] = m_arrayDepth;

            m_currentReadState[(int)LineReadState.DESC_START] = true;
            m_currentReadState[(int)LineReadState.DESC_END] = false;
        }

        public void LeaveReelDesc()
        {
            m_enteredState[(int)LineReadState.DESC_END] = m_arrayDepth;

            m_currentReadState[(int)LineReadState.DESC_START] = false;
            m_currentReadState[(int)LineReadState.DESC_END] = true;
        }

        public void ResetReelDesc()
        {
            m_currentReadState[(int)LineReadState.DESC_START] = false;
            m_currentReadState[(int)LineReadState.DESC_END] = false;
        }

        public void EnterDataSection()
        {
            m_enteredState[(int)LineReadState.DATA_START] = m_arrayDepth;

            m_currentReadState[(int)LineReadState.DATA_START] = true;
            m_currentReadState[(int)LineReadState.DATA_END] = false;
        }

        public void LeaveDataSection()
        {
            m_enteredState[(int)LineReadState.DATA_END] = m_arrayDepth;

            m_currentReadState[(int)LineReadState.DATA_START] = false;
            m_currentReadState[(int)LineReadState.DATA_END] = true;
        }

        public void ResetDataSection()
        {
            m_currentReadState[(int)LineReadState.DATA_START] = false;
            m_currentReadState[(int)LineReadState.DATA_END] = false;
        }

        public void EnterArrayLevel()
        {
            m_arrayDepth++;
        }

        public void LeaveArrayLevel()
        {
            m_arrayDepth--;
        }
    }

    public enum LineReadState : int
    {
        NONE = 0,
        DESC_START,
        DESC_END,
        DATA_START,
        DATA_END,
    };
}
