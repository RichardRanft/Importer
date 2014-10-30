using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

namespace ReelImporter
{
    public class PayParserState
    {
        private BallyPayType m_currentPayType;
        private BallyPayType m_previousPayType;

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

        public BallyPayType CurrentPayType
        {
            get
            {
                return m_currentPayType;
            }
        }

        public BallyPayType PreviousPayType
        {
            get
            {
                return m_previousPayType;
            }
        }

        public bool None
        {
            get
            {
                return m_currentReadState[(int)PayReadState.NONE];
            }
        }

        public bool SymbolStart
        {
            get
            {
                return m_currentReadState[(int)PayReadState.SYMBOLSTART];
            }
        }

        public bool SymbolEnd
        {
            get
            {
                return m_currentReadState[(int)PayReadState.SYMBOLEND];
            }
        }

        public bool LinePayStart
        {
            get
            {
                return m_currentReadState[(int)PayReadState.LINEPAYSTART];
            }
        }

        public bool LinePayEnd
        {
            get
            {
                return m_currentReadState[(int)PayReadState.LINEPAYEND];
            }
        }

        public bool FreeLinePayStart
        {
            get
            {
                return m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYSTART];
            }
        }

        public bool FreeLinePayEnd
        {
            get
            {
                return m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYEND];
            }
        }

        public bool ScatterPayStart
        {
            get
            {
                return m_currentReadState[(int)PayReadState.SCATTER_PAYSTART];
            }
        }

        public bool ScatterPayEnd
        {
            get
            {
                return m_currentReadState[(int)PayReadState.SCATTER_PAYEND];
            }
        }

        public bool FreeScatterPayStart
        {
            get
            {
                return m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYSTART];
            }
        }

        public bool FreeScatterPayEnd
        {
            get
            {
                return m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYEND];
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

        public PayParserState()
        {
            m_currentPayType = BallyPayType.NONE;
            m_previousPayType = BallyPayType.NONE;
            m_currentReadState = new BitArray(15);
            m_enteredState = new int[11];
            m_enteredState[(int)PayReadState.NONE] = 0;
            m_enteredState[(int)PayReadState.SYMBOLSTART] = 0;
            m_enteredState[(int)PayReadState.SYMBOLEND] = 0;
            m_enteredState[(int)PayReadState.LINEPAYSTART] = 0;
            m_enteredState[(int)PayReadState.LINEPAYEND] = 0;
            m_enteredState[(int)PayReadState.FREEGAME_LINEPAYSTART] = 0;
            m_enteredState[(int)PayReadState.FREEGAME_LINEPAYEND] = 0;
            m_enteredState[(int)PayReadState.SCATTER_PAYSTART] = 0;
            m_enteredState[(int)PayReadState.SCATTER_PAYEND] = 0;
            m_enteredState[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] = 0;
            m_enteredState[(int)PayReadState.FREEGAME_SCATTER_PAYEND] = 0;
            m_arrayDepth = 0;
        }

        public void EnterSymbols()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_enteredState[(int)PayReadState.SYMBOLSTART] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.SYMBOLSTART] = true;
            m_currentReadState[(int)PayReadState.SYMBOLEND] = false;
        }

        public void LeaveSymbols()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_enteredState[(int)PayReadState.SYMBOLEND] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.SYMBOLSTART] = false;
            m_currentReadState[(int)PayReadState.SYMBOLEND] = true;
        }

        public void ResetSymbols()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_currentReadState[(int)PayReadState.SYMBOLSTART] = false;
            m_currentReadState[(int)PayReadState.SYMBOLEND] = false;
        }

        public void EnterLinePay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.LINEPAY;

            m_enteredState[(int)PayReadState.LINEPAYSTART] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.LINEPAYSTART] = true;
            m_currentReadState[(int)PayReadState.LINEPAYEND] = false;
        }

        public void LeaveLinePay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_enteredState[(int)PayReadState.LINEPAYEND] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.LINEPAYSTART] = false;
            m_currentReadState[(int)PayReadState.LINEPAYEND] = true;
        }

        public void ResetLinePay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_currentReadState[(int)PayReadState.LINEPAYSTART] = false;
            m_currentReadState[(int)PayReadState.LINEPAYEND] = false;
        }

        public void EnterFreeLinePay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.FREEGAME_LINEPAY;

            m_enteredState[(int)PayReadState.FREEGAME_LINEPAYSTART] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYSTART] = true;
            m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYEND] = false;
        }

        public void LeaveFreeLinePay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_enteredState[(int)PayReadState.FREEGAME_LINEPAYEND] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYSTART] = false;
            m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYEND] = true;
        }

        public void ResetFreeLinePay()
        {
            BallyPayType tempType = m_previousPayType;
            m_previousPayType = m_currentPayType;
            m_currentPayType = tempType;

            m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYSTART] = false;
            m_currentReadState[(int)PayReadState.FREEGAME_LINEPAYEND] = false;
        }

        public void EnterScatterPay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.SCATTER_PAY;

            m_enteredState[(int)PayReadState.SCATTER_PAYSTART] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.SCATTER_PAYSTART] = true;
            m_currentReadState[(int)PayReadState.SCATTER_PAYEND] = false;
        }

        public void LeaveScatterPay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_enteredState[(int)PayReadState.SCATTER_PAYEND] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.SCATTER_PAYSTART] = false;
            m_currentReadState[(int)PayReadState.SCATTER_PAYEND] = true;
        }

        public void ResetScatterPay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_currentReadState[(int)PayReadState.SCATTER_PAYSTART] = false;
            m_currentReadState[(int)PayReadState.SCATTER_PAYEND] = false;
        }

        public void EnterFreeScatterPay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.FREEGAME_SCATTER_PAY;

            m_enteredState[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] = true;
            m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYEND] = false;
        }

        public void LeaveFreeScatterPay()
        {
            m_previousPayType = m_currentPayType;
            m_currentPayType = BallyPayType.NONE;

            m_enteredState[(int)PayReadState.FREEGAME_SCATTER_PAYEND] = m_arrayDepth;

            m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] = false;
            m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYEND] = true;
        }

        public void ResetFreeScatterPay()
        {
            BallyPayType tempType = m_previousPayType;
            m_previousPayType = m_currentPayType;
            m_currentPayType = tempType;

            m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYSTART] = false;
            m_currentReadState[(int)PayReadState.FREEGAME_SCATTER_PAYEND] = false;
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

    public enum PayReadState : int
    {
        NONE = 0,
        SYMBOLSTART,
        SYMBOLEND,
        LINEPAYSTART,
        LINEPAYEND,
        FREEGAME_LINEPAYSTART,
        FREEGAME_LINEPAYEND,
        SCATTER_PAYSTART,
        SCATTER_PAYEND,
        FREEGAME_SCATTER_PAYSTART,
        FREEGAME_SCATTER_PAYEND
    };

    public enum BallyPayType : int
    {
        NONE = 0,
        LINEPAY,
        FREEGAME_LINEPAY,
        SCATTER_PAY,
        FREEGAME_SCATTER_PAY
    };
}
