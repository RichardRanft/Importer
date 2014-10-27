using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReelImporter
{
    public class BallyReelGame : ReelGame
    {
        private BallyReelSet m_baseReelset;
        private BallyReelSet m_freeReelset;
        private BallyReelSet m_baseModReelset;
        private BallyReelSet m_freeModReelset;
        private BallyReelSet m_currentSet;

        private ParserState m_parseState;

        private int m_setIndex;
        private int m_reelWidth;
        private bool m_isValid;
        private bool m_hasModifierReels;
        private bool m_hasFreeReels;
        private bool m_hasFreeModReels;

        private Utils m_util;

        public BallyReelGame()
        {
            m_baseReelset = new BallyReelSet();
            m_freeReelset = new BallyReelSet();
            m_baseModReelset = new BallyReelSet();
            m_freeModReelset = new BallyReelSet();
            m_currentSet = null;

            m_parseState = new ParserState();

            m_util = new Utils();

            m_setIndex = 7;
            m_reelWidth = m_baseReelset.Count;
            m_isValid = false;
            m_hasModifierReels = false;
            m_hasFreeReels = false;
            m_hasFreeModReels = false;
        }

        new public BallyReelSet BaseReels
        {
            get
            {
                return m_baseReelset;
            }
            set
            {
                m_baseReelset = value;
                m_reelWidth = m_baseReelset.Count;
                m_isValid = checkValid();
            }
        }

        new public BallyReelSet BaseModifierReels
        {
            get
            {
                return m_baseModReelset;
            }
            set
            {
                m_baseModReelset = value;
                m_hasModifierReels = true;
                m_isValid = checkValid();
            }
        }

        new public BallyReelSet FreeReels
        {
            get
            {
                return m_freeReelset;
            }
            set
            {
                m_freeReelset = value;
                m_hasFreeReels = true;
                m_isValid = checkValid();
            }
        }

        new public BallyReelSet FreeModifierReels
        {
            get
            {
                return m_freeModReelset;
            }
            set
            {
                m_freeModReelset = value;
                m_hasFreeModReels = true;
                m_isValid = checkValid();
            }
        }

        new public bool IsValid
        {
            get
            {
                return m_isValid;
            }
        }

        new public bool HasModifierReels
        {
            get
            {
                return m_hasModifierReels;
            }
        }

        new public bool HasFreegameReels
        {
            get
            {
                return m_hasFreeReels;
            }
        }

        new public bool HasFreegameModifierReels
        {
            get
            {
                return m_hasFreeModReels;
            }
        }

        new public void SendToWorksheet(String sheetName, Excel.Workbook targetBook)
        {
            int setIndex = 1;
            m_setIndex = 7;
            m_currentSet = m_baseReelset;
            m_currentSet.Clean();
            exportReels(sheetName + "base" + setIndex++.ToString(), targetBook);

            List<BallyReelSet> tempSets = null;
            if (m_baseModReelset.Count > 0 && m_baseModReelset.Count == m_reelWidth)
            {
                m_currentSet = m_baseModReelset;
                m_currentSet.Clean();
                exportReels(sheetName + "base_mod" + setIndex++.ToString(), targetBook);
            }
            else if (m_baseModReelset.Count > 0 && m_baseModReelset.Count > m_reelWidth)
            {
                tempSets = getSubSets(m_baseModReelset);
                if (tempSets != null)
                {
                    foreach (BallyReelSet set in tempSets)
                    {
                        m_currentSet = set;
                        m_currentSet.Clean();
                        exportReels(sheetName + "base_mod" + setIndex++.ToString(), targetBook);
                    }
                }
            }

            if (m_freeReelset.Count > 0 && m_freeReelset.Count == m_reelWidth)
            {
                m_currentSet = m_freeReelset;
                m_currentSet.Clean();
                exportReels(sheetName + "free" + setIndex++.ToString(), targetBook);
            }
            else if (m_freeReelset.Count > 0 && m_freeReelset.Count > m_reelWidth)
            {
                tempSets = getSubSets(m_freeReelset);
                if (tempSets != null)
                {
                    foreach (BallyReelSet set in tempSets)
                    {
                        m_currentSet = set;
                        m_currentSet.Clean();
                        exportReels(sheetName + "free" + setIndex++.ToString(), targetBook);
                    }
                }
            }

            if (m_freeModReelset.Count > 0 && m_freeModReelset.Count == m_reelWidth)
            {
                m_currentSet = m_freeModReelset;
                m_currentSet.Clean();
                exportReels(sheetName + "free_mod" + setIndex++.ToString(), targetBook);
            }
            else if (m_freeModReelset.Count > 0 && m_freeModReelset.Count > m_reelWidth)
            {
                tempSets = getSubSets(m_freeModReelset);
                if (tempSets != null)
                {
                    foreach (BallyReelSet set in tempSets)
                    {
                        m_currentSet = set;
                        m_currentSet.Clean();
                        exportReels(sheetName + "free_mod" + setIndex++.ToString(), targetBook);
                    }
                }
            }
        }

        protected override void exportReels(String sheetName, Excel.Workbook targetBook)
        {
            int tableIndex = parseInteger(sheetName);
            String tableName = "";

            tableName = m_currentSet.Name + tableIndex.ToString();

            // copy the match sheet template to a new worksheet
            copyMatchSheet(tableName, targetBook);
            // copy the pay sheet template to a new worksheet
            copyPaySheet(tableName, targetBook);

            Globals.Program.Application.ScreenUpdating = false;

            tableIndex++;
            Excel.Worksheet newSheet = createSheet(tableName, targetBook);
            this.m_currentSet.SendToWorksheet(newSheet, "A1");

            // copy the reel data to the corresponding match and pay sheets
            updateMatchLinks(newSheet, targetBook, tableName, m_setIndex);
            updatePayLinks(newSheet, targetBook, tableName);

            // get this baby out from under foot - move it to the end of the workbook
            moveSheetToEnd(newSheet, targetBook);

            // let the user see that we're working
            Globals.Program.Application.ScreenUpdating = true;
        }

        protected override bool checkValid()
        {
            bool result = true;

            if (m_baseReelset.Count == 0)
                result = false;

            if (m_freeReelset.Count == 0)
                result = false;

            if (m_baseModReelset.Count < m_reelWidth && m_baseModReelset.Count != 0)
                result = false;

            if (m_freeModReelset.Count < m_reelWidth && m_freeModReelset.Count != 0)
                result = false;

            if ((m_baseModReelset.Count % m_baseReelset.Count) != 0)
                result = false;

            return result;
        }

        protected override List<ReelSet> getSubSets(ReelSet set)
        {
            throw new NotImplementedException();
        }

        protected List<BallyReelSet> getSubSets(BallyReelSet set)
        {
            if (set.Count < 1)
                return null;

            List<BallyReelSet> subset = new List<BallyReelSet>();
            BallyReelSet temp;

            // fix this - this gets modifiers for each reel, not a set of modifier reels.
            // should be index + (set.Count % m_reelWidth) - 1 while index < set.Count / m_reelWidth
            int count = 0;
            int sets = set.Count / m_reelWidth;
            do
            {
                temp = new BallyReelSet();
                temp.Name = set.Name + (count + 1).ToString() + "_";
                for (int index = count; index < set.Count; index += sets )
                {
                    temp.AddReel(set.Reels[index]);
                }
                subset.Add(temp);
                count++;
            } while(count < sets);

            return subset;
        }

        public override void Parse(StreamReader inStream)
        {
            if (m_parseState == null)
                m_parseState = new ParserState();
            bool lineHasOpenBrace = false;
            bool lineHasCloseBrace = false;
            m_baseReelset.Name = "BR_";
            m_freeReelset.Name = "FR_";
            m_baseModReelset.Name = "BR_M_";
            m_freeModReelset.Name = "FR_M_";
            BallyReel tmpReel = new BallyReel();
            String line = "";

            using (inStream)
            {
                while ((line = inStream.ReadLine()) != null)
                {
                    if (line.Contains("/"))
                    {
                        int pos = line.IndexOf("/");
                        line = line.Remove(pos);
                    }

                    line = line.Trim();

                    if (line.Length == 0 || line == "")
                        continue;

                    // look for reels
                    if (line == "reels")
                    {
                        m_parseState.EnterBaseReelSet();
                        continue;
                    }
                    if (line == "freegame_reels")
                    {
                        m_parseState.EnterFreeReelSet();
                        continue;
                    }
                    if (line == "modifierset")
                    {
                        m_parseState.EnterModifierReelSet();
                        continue;
                    }
                    if (line == "paytableoptions")
                    {
                        break;
                    }

                    if (m_parseState.ReelSetStart || m_parseState.FreeSetStart)
                    {
                        // check for braces
                        if (line == m_util.openBrace)
                        {
                            m_parseState.EnterArrayLevel();
                            continue;
                        }

                        lineHasOpenBrace = line.Contains(m_util.openBrace);
                        lineHasCloseBrace = line.Contains(m_util.closeBrace);

                        if (!lineHasOpenBrace && !lineHasCloseBrace)
                            continue;

                        if ((line.Length > 1) && lineHasOpenBrace && (m_parseState.ReelSetStart || m_parseState.FreeSetStart || m_parseState.ModifierSetStart))
                        {
                            if (!m_parseState.ReelStart && !m_parseState.FreeStart && !m_parseState.ModifierStart)
                            {
                                if (m_parseState.ReelSetStart && !m_parseState.ModifierSetStart)
                                    m_parseState.EnterBaseReel();
                                if (m_parseState.FreeSetStart && !m_parseState.ModifierSetStart)
                                    m_parseState.EnterFreeReel();
                                if (m_parseState.ModifierSetStart)
                                    m_parseState.EnterModifierReel();
                            }
                        }

                        if (lineHasCloseBrace)
                        {
                            if (line == m_util.closeBrace)
                            {
                                m_parseState.LeaveArrayLevel();
                                if (m_parseState.ModifierSetStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)ReelReadState.MODIFIERSETSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveModifierReelSet();
                                }
                                else if (m_parseState.ReelSetStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)ReelReadState.REELSETSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveBaseReelSet();
                                }
                                else if (m_parseState.FreeSetStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)ReelReadState.FREEREELSETSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveFreeReelSet();
                                }

                                continue;
                            }

                            // could be end of a reelstop definition, or moving up a level
                            if (line == m_util.arrayEnd)
                            {
                                m_parseState.LeaveArrayLevel();
                                if (m_parseState.ModifierSetStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)ReelReadState.MODIFIERSETSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveModifierReelSet();
                                }
                                else if (m_parseState.ReelSetStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)ReelReadState.REELSETSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveBaseReelSet();
                                }
                                else if (m_parseState.FreeSetStart)
                                {
                                    if (m_parseState.StateEnteredLevel[(int)ReelReadState.FREEREELSETSTART] == m_parseState.ArrayDepth)
                                        m_parseState.LeaveFreeReelSet();
                                }

                                continue;
                            }
                        }

                        if (m_parseState.ReelStart || m_parseState.FreeStart || m_parseState.ModifierStart)
                        {
                            tmpReel = new BallyReel();
                            tmpReel.Parse(inStream, line, m_parseState);
                            if (m_parseState.CurrentSetType == ReelSetType.BASEMODREEL)
                            {
                                m_baseModReelset.AddReel(tmpReel);
                                m_parseState.ResetModifierReel();
                            }

                            if (m_parseState.CurrentSetType == ReelSetType.BASEREEL)
                            {
                                m_baseReelset.AddReel(tmpReel);
                                m_parseState.ResetBaseReel();
                            }

                            if (m_parseState.CurrentSetType == ReelSetType.FREEMODREEL)
                            {
                                m_freeModReelset.AddReel(tmpReel);
                                m_parseState.ResetModifierReel();
                            }

                            if (m_parseState.CurrentSetType == ReelSetType.FREEREEL)
                            {
                                m_freeReelset.AddReel(tmpReel);
                                m_parseState.ResetFreeReel();
                            }
                        }
                    }
                }
            }
            m_isValid = checkValid();
            m_reelWidth = m_baseReelset.Count;
        }
    }
}
