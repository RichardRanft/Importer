using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

namespace ReelImporter
{
    public class ParserState
    {
        private ReelType m_currentReelType;
        private ReelType m_previousReelType;

        private ReelSetType m_currentSetType;
        private ReelSetType m_previousSetType;
        
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

        public ReelSetType CurrentSetType
        {
            get
            {
                return m_currentSetType;
            }
        }

        public ReelSetType PreviousSetType
        {
            get
            {
                return m_previousSetType;
            }
        }

        public ReelType CurrentReelType
        {
            get
            {
                return m_currentReelType;
            }
        }

        public ReelType PreviousReelType
        {
            get
            {
                return m_previousReelType;
            }
        }

        public bool None
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.NONE];
            }
        }

        public bool ReelStart
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.REELSTART];
            }
        }

        public bool ReelEnd
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.REELEND];
            }
        }

        public bool ReelSetStart
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.REELSETSTART];
            }
        }

        public bool ReelSetEnd
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.REELSETEND];
            }
        }

        public bool ModifierStart
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.MODIFIERSTART];
            }
        }

        public bool ModifierEnd
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.MODIFIEREND];
            }
        }

        public bool ModifierSetStart
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.MODIFIERSETSTART];
            }
        }

        public bool ModifierSetEnd
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.MODIFIERSETEND];
            }
        }

        public bool FreeStart
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.FREEREELSTART];
            }
        }

        public bool FreeEnd
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.FREEREELEND];
            }
        }

        public bool FreeSetStart
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.FREEREELSETSTART];
            }
        }

        public bool FreeSetEnd
        {
            get
            {
                return m_currentReadState[(int)ReelReadState.FREEREELSETEND];
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

        public ParserState()
        {
            m_currentReelType = ReelType.NONE;
            m_previousReelType = ReelType.NONE;
            m_currentSetType = ReelSetType.NONE;
            m_previousSetType = ReelSetType.NONE;
            m_currentReadState = new BitArray(13);
            m_enteredState = new int[13];
            m_enteredState[(int)ReelReadState.NONE] = 0;
            m_enteredState[(int)ReelReadState.REELSTART] = 0;
            m_enteredState[(int)ReelReadState.REELEND] = 0;
            m_enteredState[(int)ReelReadState.MODIFIERSTART] = 0;
            m_enteredState[(int)ReelReadState.MODIFIEREND] = 0;
            m_enteredState[(int)ReelReadState.FREEREELSTART] = 0;
            m_enteredState[(int)ReelReadState.FREEREELEND] = 0;
            m_enteredState[(int)ReelReadState.REELSETSTART] = 0;
            m_enteredState[(int)ReelReadState.REELSETEND] = 0;
            m_enteredState[(int)ReelReadState.MODIFIERSETSTART] = 0;
            m_enteredState[(int)ReelReadState.MODIFIERSETEND] = 0;
            m_enteredState[(int)ReelReadState.FREEREELSETSTART] = 0;
            m_enteredState[(int)ReelReadState.FREEREELSETEND] = 0;
            m_arrayDepth = 0;
        }

        public void EnterBaseReelSet()
        {
            m_previousSetType = m_currentSetType;
            m_currentSetType = ReelSetType.BASEREEL;

            m_enteredState[(int)ReelReadState.REELSETSTART] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.REELSETSTART] = true;
            m_currentReadState[(int)ReelReadState.REELSETEND] = false;
        }

        public void LeaveBaseReelSet()
        {
            m_previousSetType = m_currentSetType;
            m_currentSetType = ReelSetType.NONE;

            m_enteredState[(int)ReelReadState.REELSETEND] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.REELSETSTART] = false;
            m_currentReadState[(int)ReelReadState.REELSETEND] = true;
        }

        public void EnterBaseReel()
        {
            m_previousReelType = m_currentReelType;
            m_currentReelType = ReelType.BASEREEL;

            m_enteredState[(int)ReelReadState.REELSTART] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.REELSTART] = true;
            m_currentReadState[(int)ReelReadState.REELEND] = false;
        }

        public void LeaveBaseReel()
        {
            m_previousReelType = m_currentReelType;
            m_currentReelType = ReelType.NONE;

            m_enteredState[(int)ReelReadState.REELEND] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.REELSTART] = false;
            m_currentReadState[(int)ReelReadState.REELEND] = true;
        }

        public void ResetBaseReel()
        {
            m_previousReelType = m_currentReelType;
            m_currentReelType = ReelType.NONE;

            m_currentReadState[(int)ReelReadState.REELSTART] = false;
            m_currentReadState[(int)ReelReadState.REELEND] = false;
        }

        public void EnterFreeReelSet()
        {
            m_previousSetType = m_currentSetType;
            m_currentSetType = ReelSetType.FREEREEL;

            m_enteredState[(int)ReelReadState.FREEREELSETSTART] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.FREEREELSETSTART] = true;
            m_currentReadState[(int)ReelReadState.FREEREELSETEND] = false;
        }

        public void LeaveFreeReelSet()
        {
            m_previousSetType = m_currentSetType;
            m_currentSetType = ReelSetType.NONE;

            m_enteredState[(int)ReelReadState.FREEREELSETEND] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.FREEREELSETSTART] = false;
            m_currentReadState[(int)ReelReadState.FREEREELSETEND] = true;
        }

        public void EnterFreeReel()
        {
            m_previousReelType = m_currentReelType;
            m_currentReelType = ReelType.FREEREEL;

            m_enteredState[(int)ReelReadState.FREEREELSTART] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.FREEREELSTART] = true;
            m_currentReadState[(int)ReelReadState.FREEREELEND] = false;
        }

        public void LeaveFreeReel()
        {
            m_previousReelType = m_currentReelType;
            m_currentReelType = ReelType.NONE;

            m_enteredState[(int)ReelReadState.FREEREELEND] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.FREEREELSTART] = false;
            m_currentReadState[(int)ReelReadState.FREEREELEND] = true;
        }

        public void ResetFreeReel()
        {
            ReelType tempType = m_previousReelType;
            m_previousReelType = m_currentReelType;
            m_currentReelType = tempType;

            m_currentReadState[(int)ReelReadState.FREEREELSTART] = false;
            m_currentReadState[(int)ReelReadState.FREEREELEND] = false;
        }

        public void EnterModifierReelSet()
        {
            m_previousSetType = m_currentSetType;

            if (m_currentSetType == ReelSetType.BASEREEL)
                m_currentSetType = ReelSetType.BASEMODREEL;
            else
                m_currentSetType = ReelSetType.FREEMODREEL;
            
            m_enteredState[(int)ReelReadState.MODIFIERSETSTART] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.MODIFIERSETSTART] = true;
            m_currentReadState[(int)ReelReadState.MODIFIERSETEND] = false;
        }

        public void LeaveModifierReelSet()
        {
            ReelSetType tempSetType = m_previousSetType;
            m_previousSetType = m_currentSetType;
            m_currentSetType = tempSetType;

            m_enteredState[(int)ReelReadState.MODIFIERSETEND] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.MODIFIERSETSTART] = false;
            m_currentReadState[(int)ReelReadState.MODIFIERSETEND] = true;
        }

        public void EnterModifierReel()
        {
            m_previousReelType = m_currentReelType;
            if (m_currentSetType == ReelSetType.FREEMODREEL)
                m_currentReelType = ReelType.FREEMODREEL;
            else
                m_currentReelType = ReelType.BASEMODREEL;

            m_enteredState[(int)ReelReadState.MODIFIERSTART] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.MODIFIERSTART] = true;
            m_currentReadState[(int)ReelReadState.MODIFIEREND] = false;
        }

        public void LeaveModifierReel()
        {
            ReelType tempType = m_previousReelType;
            m_previousReelType = m_currentReelType;
            m_currentReelType = tempType;

            m_enteredState[(int)ReelReadState.MODIFIEREND] = m_arrayDepth;

            m_currentReadState[(int)ReelReadState.MODIFIERSTART] = false;
            m_currentReadState[(int)ReelReadState.MODIFIEREND] = true;
        }

        public void ResetModifierReel()
        {
            ReelType tempType = m_previousReelType;
            m_previousReelType = m_currentReelType;
            m_currentReelType = tempType;

            m_currentReadState[(int)ReelReadState.MODIFIERSTART] = false;
            m_currentReadState[(int)ReelReadState.MODIFIEREND] = false;
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

    public enum ReelDataType : int
    {
        SHFL = 0,
        BALLY
    };

    public enum ReelType : int
    {
        NONE = 0,
        BASEREEL,
        BASEMODREEL,
        FREEREEL,
        FREEMODREEL
    };

    public enum ReelSetType : int
    {
        NONE = 0,
        BASEREEL,
        BASEMODREEL,
        FREEREEL,
        FREEMODREEL
    };

    public enum ReelReadState : int
    {
        NONE = 0,
        REELSTART,
        REELEND,
        MODIFIERSTART,
        MODIFIEREND,
        FREEREELSTART,
        FREEREELEND,
        REELSETSTART,
        REELSETEND,
        MODIFIERSETSTART,
        MODIFIERSETEND,
        FREEREELSETSTART,
        FREEREELSETEND
    };
}
