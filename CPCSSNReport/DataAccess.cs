using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;

namespace CPCSSNReport
{
    class DataAccess
    {
        OleDbConnection conn;
        Dictionary<string, int> dictDemoGroup;
        Dictionary<string, string> dictDemoStr;
        List<string> lstPractices;
        List<string> lstProviders;
        string connStrCurrDB, connStrPrevDB;
        string recentDataDrawnDate, lastDataDrawnDate;
        string sCutOffDate;
        string sBeginDate;
        bool Year1 = false;

        public string ContactGroup
        {
            get
            {
                string SQL ="";
                if (Year1)
                {
                    SQL = "(SELECT p1.*, pd1.Site_ID FROM Patient_1Yr p1, PatientDemographic pd1 WHERE p1.Patient_ID = pd1.Patient_ID) p";                    
                }
                else
                {
                    SQL = "(SELECT p1.*, pd1.Site_ID FROM Patient p1, PatientDemographic pd1 WHERE pd1.PatientStatus_calc='Active' AND p1.Patient_ID = pd1.Patient_ID) p";
                }
                return SQL;
            }
        }


        public DataAccess(string currDB, string prevDB)
        {
            connStrCurrDB = currDB;
            connStrPrevDB = prevDB;
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + connStrCurrDB);
            conn.Open();
        }

        public void LoadProviderInfo()
        {            
            string SQL = "SELECT Site_ID, Provider_ID FROM SiteProvider";
            lstPractices = new List<string>();
            lstProviders = new List<string>();
            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                HashSet<string> hsPratices = new HashSet<string>();
                while (reader.Read())
                {
                    string provider_id = reader.GetInt32(1).ToString();
                    lstProviders.Add(provider_id);
                    string practices_id = reader.GetInt32(0).ToString();
                    if (!lstPractices.Contains(practices_id))
                    {
                        lstPractices.Add(practices_id);
                    }
                }
                reader.Close();                
            }
            catch (OleDbException ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
            cmmd = null;

            string SQLDate = "SELECT CutOffDate FROM Cycle WHERE CutOffDate = (SELECT MAX(CutOffDate) FROM Cycle)";
            cmmd = new OleDbCommand(SQLDate, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                if (reader.Read())
                {
                    DateTime dtCutOff = reader.GetDateTime(0);
                    recentDataDrawnDate = dtCutOff.ToShortDateString();
                    dtCutOff = dtCutOff.AddDays(1);
                    sCutOffDate = dtCutOff.Year + "-" + dtCutOff.Month + "-" + dtCutOff.Day;
                    DateTime dtBegin = dtCutOff.AddMonths(-3);
                    sBeginDate = dtBegin.Year + "-" + dtBegin.Month + "-" + dtBegin.Day;
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmd = null;
        }

        public List<string> GetProviders()
        {
            return lstProviders;
        }

        public List<string> GetPractices()
        {
            return lstPractices;
        }

        public List<string> GetKeys()
        {
            return dictDemoStr.Keys.ToList<string>();
        }

        protected void GetDemoGroup(string filter)
        {
            string SQL = "SELECT p.Sex";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=5, p.Patient_ID, NULL)) AS Age5less";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=17 AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=6, p.Patient_ID, NULL)) AS Age6_17";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=17, p.Patient_ID, NULL)) AS Age17less";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=39 AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=18, p.Patient_ID, NULL)) AS Age18_39";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=64  AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=40, p.Patient_ID, NULL)) AS Age40_64";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=79 AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=65, p.Patient_ID, NULL)) AS Age65_79";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)>=80, p.Patient_ID, NULL)) AS Age80plus";
            SQL = SQL + ", COUNT(p.Patient_ID) AS AllAge";
            SQL = SQL + " FROM " + ContactGroup;
            SQL = SQL + " WHERE p.BirthYear IS NOT NULL AND p.Sex IS NOT NULL";
            SQL = SQL + filter;            
            SQL = SQL + " GROUP BY p.Sex";
            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    // require sex is not null
                    if (!reader.IsDBNull(0))
                    {
                        for (int i = 1; i < 9; i++)
                        {
                            string key = reader.GetString(0) + '_' + reader.GetName(i);
                            dictDemoGroup.Add(key, reader.GetInt32(i));
                        }
                    }
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
        }

        protected void GetIndexDiseaseDemoGroup(string filter)
        {
            string SQL = "SELECT p.Sex, i.Disease";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=17, p.Patient_ID, NULL)) AS Age17less";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=39 AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=18, p.Patient_ID, NULL)) AS Age18_39";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=64  AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=40, p.Patient_ID, NULL)) AS Age40_64";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)<=79 AND (YEAR(#" + sCutOffDate + "#)-BirthYear)>=65, p.Patient_ID, NULL)) AS Age65_79";
            SQL = SQL + ", COUNT(IIF((YEAR(#" + sCutOffDate + "#)-BirthYear)>=80, p.Patient_ID, NULL)) AS Age80plus";
            SQL = SQL + ", COUNT(p.Patient_ID) AS AllAge";
            SQL = SQL + " FROM " + ContactGroup + ", DiseaseCase i";
            SQL = SQL + " WHERE p.BirthYear IS NOT NULL";
            SQL = SQL + filter;
            SQL = SQL + " AND p.Patient_ID = i.Patient_ID GROUP BY i.Disease, p.Sex";
            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        for (int i = 2; i < 8; i++)
                        {
                            string indexdisease;
                            switch (reader.GetString(1))
                            {
                                case "Diabetes Mellitus":
                                    indexdisease = "DM";
                                    break;
                                case "Hypertension":
                                    indexdisease = "HTN";
                                    break;
                                case "Osteoarthritis":
                                    indexdisease = "OA";
                                    break;
                                case "Depression":
                                    indexdisease = "DP";
                                    break;
                                case "Epilepsy":
                                    indexdisease = "EPL";
                                    break;
                                case "Dementia":
                                    indexdisease = "DEM";
                                    break;
                                case "Parkinson's Disease":
                                    indexdisease = "PAK";
                                    break;
                                default:
                                    indexdisease = reader.GetString(1);
                                    break;
                            }
                            string key = reader.GetString(0) + '_' + indexdisease + '_' + reader.GetName(i);
                            dictDemoGroup.Add(key, reader.GetInt32(i));
                        }
                    }
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
        }

        protected void GetPrevalence(string filter)
        {
            Dictionary<string, List<int>> dictPrevalence = new Dictionary<string, List<int>>();
            string SQL = "SELECT i.Disease, i.Patient_ID FROM DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE i.Patient_ID = p.Patient_ID";
            SQL = SQL + " AND p.BirthYear IS NOT NULL AND p.Sex IS NOT NULL AND LEN(TRIM(p.Sex))>0";
            SQL = SQL + filter;

            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    List<int> lstIdxPt;
                    if (dictPrevalence.ContainsKey(reader.GetString(0)))
                    {
                        lstIdxPt = dictPrevalence[reader.GetString(0)];
                    }
                    else
                    {
                        lstIdxPt = new List<int>();
                        dictPrevalence.Add(reader.GetString(0), lstIdxPt);
                    }                    
                    lstIdxPt.Add(reader.GetInt32(1));
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmd = null;            


            Dictionary<string, HashSet<int>> dictPrePrevalence = new Dictionary<string, HashSet<int>>();
            OleDbConnection connPre = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + connStrPrevDB);
            connPre.Open();
            SQL = "SELECT i.Disease, i.Patient_ID FROM DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE i.Patient_ID = p.Patient_ID";
            SQL = SQL + " AND p.BirthYear IS NOT NULL AND p.Sex IS NOT NULL AND LEN(TRIM(p.Sex))>0";
            SQL = SQL + filter;
            OleDbCommand cmmdPre = new OleDbCommand(SQL, connPre);
            try
            {
                OleDbDataReader reader = cmmdPre.ExecuteReader();
                while (reader.Read())
                {
                    HashSet<int> hsIdxPt;
                    if (dictPrePrevalence.ContainsKey(reader.GetString(0)))
                    {
                        hsIdxPt = dictPrePrevalence[reader.GetString(0)];
                    }
                    else
                    {
                        hsIdxPt = new HashSet<int>();
                        dictPrePrevalence.Add(reader.GetString(0), hsIdxPt);
                    }
                    hsIdxPt.Add(reader.GetInt32(1));
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmdPre = null;

            string SQLDate = "SELECT CutOffDate FROM Cycle WHERE CutOffDate = (SELECT MAX(CutOffDate) FROM Cycle)";
            cmmdPre = new OleDbCommand(SQLDate, connPre);
            try
            {
                OleDbDataReader reader = cmmdPre.ExecuteReader();
                if (reader.Read())
                {
                    lastDataDrawnDate = reader.GetDateTime(0).ToShortDateString();
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            connPre.Close();
            connPre = null;

            foreach (string s in dictPrevalence.Keys)
            {
                List<int> lstIdxPt = dictPrevalence[s];
                dictDemoGroup.Add(s.Replace(' ', '_').Replace("'s",""), lstIdxPt.Count);                
                int newPt = 0;
                foreach (int idxPt in lstIdxPt)
                {
                    if (dictPrePrevalence.ContainsKey(s) && !dictPrePrevalence[s].Contains(idxPt))
                    {
                        newPt++;
                    }
                }
                dictDemoGroup.Add("New_" + s.Replace(' ', '_'), newPt);

                string indexdisease;
                switch (s)
                {
                    case "Diabetes Mellitus":
                        indexdisease = "DM";
                        break;
                    case "Hypertension":
                        indexdisease = "HTN";
                        break;
                    case "Osteoarthritis":
                        indexdisease = "OA";
                        break;
                    case "Depression":
                        indexdisease = "DP";
                        break;
                    case "Epilepsy":
                        indexdisease = "EPL";
                        break;
                    case "Dementia":
                        indexdisease = "DEM";
                        break;
                    case "Parkinson's Disease":
                        indexdisease = "PAK";
                        break;
                    default:
                        indexdisease = s;
                        break;
                }
                dictDemoGroup.Add("NEW_" + indexdisease + "_PT", newPt);
            }
        }

        public string GetValue(string s)
        {
            return dictDemoStr[s];
        }

        public void LoadData(string filter, string raw_filter)
        {
            if (dictDemoGroup == null)
            {
                dictDemoGroup = new Dictionary<string, int>();
                dictDemoStr = new Dictionary<string, string>();
            }
            else
            {
                dictDemoGroup.Clear();
                dictDemoStr.Clear();
            }
            GetPrevalence(filter);
            GetDemoGroup(filter);            
            GetIndexDiseaseDemoGroup(filter);
            GetHTNTarget(filter);
            GetDMTarget(filter);
            GetMedication(filter);
            GetLabScreenProc(filter);
            GetExamScreenProc(filter);
            GetSmokingStatus(filter);
            CalSum();
            CalDenominator();
            CalPct();
            GetDataDrawnDate();
            GetPracticeOf(raw_filter);
        }

        protected void GetHTNTarget(string filter)
        {
            HashSet<int> hsDMPatient = new HashSet<int>();
            string sqlDM = "SELECT Patient_ID FROM DiseaseCase WHERE Disease = 'Diabetes Mellitus'";
            OleDbCommand cmmd = new OleDbCommand(sqlDM, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    hsDMPatient.Add(reader.GetInt32(0));
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmd = null;

            HashSet<int> hsHTNOnlyMalePt = new HashSet<int>();
            HashSet<int> hsHTNOnlyFemalePt = new HashSet<int>();
            HashSet<int> hsHTNDMMalePt = new HashSet<int>();
            HashSet<int> hsHTNDmFemalePt = new HashSet<int>();

            string sqlHTN = "SELECT p.Patient_ID, p.Sex FROM DiseaseCase i, " + ContactGroup;
            sqlHTN = sqlHTN + " WHERE p.Patient_ID = i.Patient_ID";
            sqlHTN = sqlHTN + " AND i.Disease = 'Hypertension'";
            sqlHTN = sqlHTN + filter;
            cmmd = new OleDbCommand(sqlHTN, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    int sPatientID = reader.GetInt32(0);
                    if (!reader.IsDBNull(1))
                    {
                        string sSex = reader.GetString(1);

                        if (hsDMPatient.Contains(sPatientID))
                        {
                            if (sSex == "Male")
                            {
                                hsHTNDMMalePt.Add(sPatientID);
                            }
                            else
                            {
                                hsHTNDmFemalePt.Add(sPatientID);
                            }
                        }
                        else
                        {
                            if (sSex == "Male")
                            {
                                hsHTNOnlyMalePt.Add(sPatientID);
                            }
                            else
                            {
                                hsHTNOnlyFemalePt.Add(sPatientID);
                            }
                        }
                    }
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmd = null;

            string sqlView = "SELECT e1.* FROM Exam AS e1, (SELECT Max(DateCreated) AS MaxDateCreated, Patient_ID, Exam1 FROM Exam GROUP BY Patient_ID, Exam1) AS e2";
            sqlView = sqlView + " WHERE e1.Patient_ID = e2.Patient_ID and e1.Exam1 = e2.Exam1 and e1.DateCreated = e2.MaxDateCreated";
            string SQL = "SELECT p.Patient_ID, p.Sex, pe.Exam1, AVG(CDBL(pe.Result1_calc)), pe.Exam2, AVG(CDBL(pe.Result2_calc)) FROM "+ContactGroup+", DiseaseCase i, (" + sqlView + ") AS pe";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND pe.Exam1 = 'sBP (mmHg)'";
            SQL = SQL + " AND ISNUMERIC(pe.Result1_calc) AND ISNUMERIC(pe.Result2_calc)";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Patient_ID, p.Sex, pe.Exam1, pe.Exam2";

            cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                int nSBPLe140M = 0, nDBPLe90M = 0, nSBPDMLe130M = 0, nDBPDMLe80M = 0;
                int nSBPLe140F = 0, nDBPLe90F = 0, nSBPDMLe130F = 0, nDBPDMLe80F = 0;
                               

                while (reader.Read())
                {                    
                    if (!reader.IsDBNull(1))
                    {
                        int sPatientID = reader.GetInt32(0);
                        string sSex = reader.GetString(1);                        
                        double dSBP = reader.GetDouble(3);
                        double dDBP = reader.GetDouble(5);

                        if (hsDMPatient.Contains(sPatientID))
                        {
                            if (dDBP <= 80)
                            {
                                if (sSex == "Male")
                                    nDBPDMLe80M++;
                                else
                                    nDBPDMLe80F++;
                            }
                            if (dSBP <= 130)
                            {
                                if (sSex == "Male")
                                    nSBPDMLe130M++;
                                else
                                    nSBPDMLe130F++;
                            }
                        }
                        else
                        {
                            if (dDBP <= 90)
                            {
                                if (sSex == "Male")
                                    nDBPLe90M++;
                                else
                                    nDBPLe90F++;

                            }
                            if (dSBP <= 140)
                            {
                                if (sSex == "Male")
                                    nSBPLe140M++;
                                else
                                    nSBPLe140F++;
                            }
                        }                        
                    }
                }
                //dictDemoGroup.Add("Male_HTN_DM_DBP_LE_80", nDBPDMLe80M);
                //dictDemoGroup.Add("Male_HTN_DM_SBP_LE_130", nSBPDMLe130M);
                //dictDemoGroup.Add("Male_HTN_DBP_LE_90", nDBPLe90M);
                //dictDemoGroup.Add("Male_HTN_SBP_LE_140", nSBPLe140M);
                string value, pct;

                int nDMM = hsHTNDMMalePt.Count;
                int nDMF = hsHTNDmFemalePt.Count;
                int nM = hsHTNOnlyMalePt.Count;
                int nF = hsHTNOnlyFemalePt.Count;

                value = nDBPDMLe80M.ToString();
                if (nDMM ==0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nDBPDMLe80M / nDMM * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Male_HTN_DM_DBP_LE_80", value);

                value = nSBPDMLe130M.ToString();
                if (nDMM == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nSBPDMLe130M / nDMM * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Male_HTN_DM_SBP_LE_130", value);
                
                value = nDBPLe90M.ToString();
                if (nM == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nDBPLe90M / nM * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Male_HTN_DBP_LE_90", value);

                value = nSBPLe140M.ToString();
                if (nM == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nSBPLe140M / nM * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Male_HTN_SBP_LE_140", value);

                //dictDemoGroup.Add("Female_HTN_DM_DBP_LE_80", nDBPDMLe80F);
                //dictDemoGroup.Add("Female_HTN_DM_SBP_LE_130", nSBPDMLe130F);
                //dictDemoGroup.Add("Female_HTN_DBP_LE_90", nDBPLe90F);
                //dictDemoGroup.Add("Female_HTN_SBP_LE_140", nSBPLe140F);

                value = nDBPDMLe80F.ToString();
                if (nDMF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nDBPDMLe80F / nDMF * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Female_HTN_DM_DBP_LE_80", value);

                value = nSBPDMLe130F.ToString();
                if (nDMF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nSBPDMLe130F / nDMF * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Female_HTN_DM_SBP_LE_130", value);

                value = nDBPLe90F.ToString();
                if (nF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nDBPLe90F / nF * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Female_HTN_DBP_LE_90", value);

                value = nSBPLe140F.ToString();
                if (nF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)nSBPLe140F / nF * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Female_HTN_SBP_LE_140", value);

                //Sum

                value = (nDBPDMLe80M + nDBPDMLe80F).ToString();
                if (nDMM + nDMF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)(nDBPDMLe80M + nDBPDMLe80F) / (nDMM + nDMF) * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Sum_HTN_DM_DBP_LE_80", value);

                value = (nSBPDMLe130M + nSBPDMLe130F).ToString();
                if (nDMM + nDMF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)(nSBPDMLe130M + nSBPDMLe130F) / (nDMM + nDMF) * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Sum_HTN_DM_SBP_LE_130", value);

                value = (nDBPLe90M + nDBPLe90F).ToString();
                if (nM + nF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double) (nDBPLe90M + nDBPLe90F) / (nM + nF) * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Sum_HTN_DBP_LE_90", value);

                value = (nSBPLe140M + nSBPLe140F).ToString();
                if (nM + nF == 0)
                    pct = 0.ToString("0.#") + "%";
                else
                    pct = ((double)(nSBPLe140M + nSBPLe140F) / (nM + nF) * 100).ToString("0.#") + "%";
                value = value + "(" + pct + ")";
                dictDemoStr.Add("Sum_HTN_SBP_LE_140", value);

                dictDemoGroup.Add("Male_HTN_AT_DBP_TGT", nDBPDMLe80M + nDBPLe90M);
                dictDemoGroup.Add("Male_HTN_AT_SBP_TGT", nSBPDMLe130M + nSBPLe140M);

                dictDemoGroup.Add("Female_HTN_AT_DBP_TGT", nDBPDMLe80F + nDBPLe90F);
                dictDemoGroup.Add("Female_HTN_AT_SBP_TGT", nSBPDMLe130F + nSBPLe140F);
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
        }

        protected void GetDMTarget(string filter)
        {
            HashSet<int> hsTargetMetPtM = new HashSet<int>();
            HashSet<int> hsTargetMetPtF = new HashSet<int>();
            
            string sqlView = "SELECT e1.* FROM Exam AS e1, (SELECT Max(DateCreated) AS MaxDateCreated, Patient_ID, Exam1 FROM Exam GROUP BY Patient_ID, Exam1) AS e2";
            sqlView = sqlView + " WHERE e1.Patient_ID = e2.Patient_ID and e1.Exam1 = e2.Exam1 and e1.DateCreated = e2.MaxDateCreated";
            string SQL = "SELECT p.Patient_ID, p.Sex, pe.Exam1, AVG(CDBL(pe.Result1_calc)), pe.Exam2, AVG(CDBL(pe.Result2_calc)) FROM "+ContactGroup+", DiseaseCase i, (" + sqlView + ") AS pe";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND pe.Exam1 = 'sBP (mmHg)'";
            SQL = SQL + " AND ISNUMERIC(pe.Result1_calc) AND ISNUMERIC(pe.Result2_calc)";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Patient_ID, p.Sex, pe.Exam1, pe.Exam2";

            HashSet<int> hsDBPMetPt = new HashSet<int>();
            HashSet<int> hsSBPMetPt = new HashSet<int>();

            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                int nSBPLe130M = 0, nDBPLe80M = 0, nSBPLe130F = 0, nDBPLe80F = 0;
                while (reader.Read())
                {
                    if (!reader.IsDBNull(1))
                    {
                        int sPatientID = reader.GetInt32(0);
                        string sSex = reader.GetString(1);
                        string sExam = reader.GetString(2);
                        double dSBP = reader.GetDouble(3);
                        double dDBP = reader.GetDouble(5);
                        if (sSex == "Male")
                            hsTargetMetPtM.Add(sPatientID);
                        else
                            hsTargetMetPtF.Add(sPatientID);


                        if (dDBP <= 80)
                        {
                            if (sSex == "Male")
                            {
                                nDBPLe80M++;
                            }
                            else
                            {
                                nDBPLe80F++;
                            }
                            hsDBPMetPt.Add(sPatientID);
                        }

                        if (dSBP <= 130)
                        {
                            if (sSex == "Male")
                                nSBPLe130M++;
                            else
                                nSBPLe130F++;
                            hsSBPMetPt.Add(sPatientID);
                        }

                    }
                }
                reader.Close();
                dictDemoGroup.Add("Male_DM_DBP_LE_80", nDBPLe80M);
                dictDemoGroup.Add("Male_DM_SBP_LE_130", nSBPLe130M);

                dictDemoGroup.Add("Female_DM_DBP_LE_80", nDBPLe80F);
                dictDemoGroup.Add("Female_DM_SBP_LE_130", nSBPLe130F);
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }

            List<int> lstBPMetM = new List<int>();
            List<int> lstBPMetF = new List<int>();

            foreach (int s in hsTargetMetPtM)
            {
                if (hsSBPMetPt.Contains(s) && hsDBPMetPt.Contains(s))
                {
                    lstBPMetM.Add(s);
                }
            }

            foreach (int s in hsTargetMetPtF)
            {
                if (hsSBPMetPt.Contains(s) && hsDBPMetPt.Contains(s))
                {
                    lstBPMetF.Add(s);
                }
            }

            hsTargetMetPtM.Clear();
            hsTargetMetPtF.Clear();
            hsSBPMetPt.Clear();
            hsDBPMetPt.Clear();
            hsTargetMetPtM = null;
            hsTargetMetPtF = null;
            hsSBPMetPt = null;
            hsDBPMetPt = null;

            sqlView = "SELECT a.* FROM Lab a, (SELECT Name_calc, Patient_ID, MAX(DateCreated) AS recentDate FROM Lab WHERE Name_calc IN ('LDL','HBA1C') GROUP BY Name_calc, Patient_ID) b";
            sqlView = sqlView + " WHERE a.Name_calc IN ('LDL','HBA1C')";
            sqlView = sqlView + " AND a.DateCreated = b.recentDate AND a.Name_calc = b.Name_calc AND a.Patient_ID = b.Patient_ID";
            
            SQL = "SELECT p.Patient_ID, p.Sex, lr.Name_calc, lr.TestResult FROM "+ContactGroup+", DiseaseCase i, (" + sqlView + ") AS lr";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = lr.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + filter;

            HashSet<int> hsLDLMetPt = new HashSet<int>();
            HashSet<int> hsHBA1CMetPt = new HashSet<int>();

            cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                int nLDLLe2M = 0, nLDLLe2F = 0, nHBA1CLe7M = 0, nHBA1CLe7F = 0;
                while (reader.Read())
                {
                    if (!reader.IsDBNull(1))
                    {
                        int sPatientID = reader.GetInt32(0);
                        string sSex = reader.GetString(1);
                        string sExam = reader.GetString(2);
                        string sExamResult = reader.GetString(3);
                        float fExamResult = 0;

                        if (float.TryParse(sExamResult, out fExamResult))
                        {
                            if (sExam == "LDL" && fExamResult <= 2.0)
                            {
                                if (sSex == "Male")
                                    nLDLLe2M++;
                                else
                                    nLDLLe2F++;
                                hsLDLMetPt.Add(sPatientID);
                            }

                            if (sExam == "HBA1C" && fExamResult <= 7.0)
                            {
                                if (sSex == "Male")
                                    nHBA1CLe7M++;
                                else
                                    nHBA1CLe7F++;
                                hsHBA1CMetPt.Add(sPatientID);
                            }
                        }
                    }
                }
                reader.Close();

                dictDemoGroup.Add("Male_DM_LDL_LE_2", nLDLLe2M);
                dictDemoGroup.Add("Male_DM_HBA1C_LE_7", nHBA1CLe7M);

                dictDemoGroup.Add("Female_DM_LDL_LE_2", nLDLLe2F);
                dictDemoGroup.Add("Female_DM_HBA1C_LE_7", nHBA1CLe7F);
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }

            List<int> lstAllMetM = new List<int>();
            List<int> lstAllMetF = new List<int>();

            foreach (int s in lstBPMetM)
            {
                if (hsLDLMetPt.Contains(s) && hsHBA1CMetPt.Contains(s))
                {
                    lstAllMetM.Add(s);
                }
            }

            foreach (int s in lstBPMetF)
            {
                if (hsLDLMetPt.Contains(s) && hsHBA1CMetPt.Contains(s))
                {
                    lstAllMetF.Add(s);
                }
            }

            dictDemoGroup.Add("Male_DM_Met_3Targets", lstAllMetM.Count);
            dictDemoGroup.Add("Female_DM_Met_3Targets", lstAllMetF.Count);
        }

        protected void GetMedication(string filter)
        {
            //Hyertension Thiazides
            string SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_Thiazide_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,4) IN ('C03A', 'C07B', 'C07D', 'C07C') OR LEFT(Code_calc,5) IN ('C03EA', 'C09BA', 'C09DA') OR LEFT(Code_calc,7) IN ('C09DX01', 'C09DX03'))";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_ACEI_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,4) IN ('C09A', 'C09B')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_ARB_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;             
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,4) IN ('C09C', 'C09D')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_CCB_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;                     
            SQL = SQL + "  WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,3) IN ('C08') OR LEFT(Code_calc,4) IN ('C07F') OR LEFT(Code_calc,5) IN ('C09BB', 'C09DB'))";
            SQL = SQL + " AND (LEFT(Code_calc,5) NOT IN ('C08DA'))";  // remove verapamil, not used for anti-hypertensive
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_BB_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;               
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,3) IN ('C07')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_Other_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,3) IN ('C02'))";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex";
            SQL = SQL + ", (IIF(Thiazide>=1,1,0) + IIF(ACEI>=1,1,0) + IIF(ARB>=1,1,0) + IIF(CCB>=1,1,0) + IIF(BB>=1,1,0)) AS HTN_MedNum";
            SQL = SQL + ", Count(Patient_ID) AS HTN_PtNum";
            SQL = SQL + " FROM (";
            SQL = SQL + " SELECT p.Sex, m.Patient_ID";
            SQL = SQL + ", COUNT(IIF(LEFT(Code_calc,4) IN ('C03A', 'C07B') OR LEFT(Code_calc,5) IN ('C03EA', 'C03BA', 'C03BB'), 1, NULL)) AS Thiazide";
            SQL = SQL + ", COUNT(IIF(LEFT(Code_calc,4) IN ('C09A', 'C09B'), 1, NULL)) AS ACEI";
            SQL = SQL + ", COUNT(IIF(LEFT(Code_calc,4) IN ('C09C', 'C09D'), 1, NULL)) AS ARB";
            SQL = SQL + ", COUNT(IIF(LEFT(Code_calc,4) IN ('C08C', 'C08D', 'C08E') OR LEFT(Code_calc,5) IN ('C09BB', 'C09DB'), 1, NULL)) AS CCB";
            SQL = SQL + ", COUNT(IIF(LEFT(Code_calc,3) IN ('C07'), 1, NULL)) AS BB";
            SQL = SQL + " FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS MedPt";
            SQL = SQL + " GROUP BY Sex, (IIF(Thiazide>=1,1,0) + IIF(ACEI>=1,1,0) + IIF(ARB>=1,1,0) + IIF(CCB>=1,1,0) + IIF(BB>=1,1,0))";

            string SQLStr = "SELECT Sex, IIF(HTN_MedNum >=4, 'HTN_Med_4plus', 'HTN_Med_' & HTN_MedNum) AS HTN_MedNum_Grp, SUM(HTN_PtNum) AS HTN_MedNum_SUM FROM ("
                + SQL + ") AS MedSum GROUP BY Sex, IIF(HTN_MedNum >=4, 'HTN_Med_4plus', 'HTN_Med_' & HTN_MedNum)";
            SQLExec(SQLStr);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS OA_NSAID_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Osteoarthritis'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,4) IN ('M01A', 'M01B') OR LEFT(Code_calc,5) IN ('M02AA'))";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS OA_Acetaminophen_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Osteoarthritis'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,7) IN ('N02BE01', 'N02AA59', 'N02BE51')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS OA_Steroid_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Osteoarthritis'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('H02AB') AND INSTR(m.Name_Orig, 'ml')>0";            
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS OA_Opiate_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Osteoarthritis'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('N02AA', 'N02AX')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_InhAC_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('R03BB')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_InhSteroid_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('R03BA', 'R03AK')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_InhB2A_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('R03AC', 'R03CC', 'R03AK')";
            SQL = SQL + " AND LEFT(Code_calc,7) NOT IN ('R03AK01', 'R03AK02')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_LABA_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,7) IN ('R03AK06','R03AK07','R03CC12','R03CC13','R03AC14','R03AC13','R03AC12','R03AC18')"; 
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_SABA_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,7) IN ('R03AC02','R03AC03','R03AC04','R03AC05','R03AC08','R03AC10','R03AC11','R03AC15','R03AC16','R03AC17','R03AK03','R03AK04','R03AK05','R03CC02','R03CC03','R03CC04','R03CC07','R03CC08','R03CC10','R03CC11','R03CC14')"; 
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_OralSteroid_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('H02AB') AND INSTR(m.Name_Orig, 'ml')=0";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_Other_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('R03DA')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Metformin_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,5) IN ('A10BA') OR LEFT(Code_calc,7) IN ('A10BD02','A10BD03','A10BD05','A10BD07','A10BD08'))";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Sulfonylure_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,5) IN ('A10BB', 'A10BC') OR LEFT(Code_calc,7) IN ('A10BD01','A10BD02','A10BD04','A10BD06'))";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Glitazones_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,5) IN ('A10BG') OR LEFT(Code_calc,7) IN ('A10BD03','A10BD04','A10BD05','A10BD06','A10BD09'))";            
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Gliptins_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND (LEFT(Code_calc,5) IN ('A10BH') OR LEFT(Code_calc,7) IN ('A10BD07','A10BD08','A10BD09'))";                  
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Acarbose_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('A10BF')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Insulin_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,4) IN ('A10A')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_Other_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('A10BX')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DP_Tricyclics_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Depression'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('N06AA')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DP_SSRI_SNRI_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Depression'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('N06AB','N06AX')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DP_Benzo_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Depression'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('N05BA')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DP_Antipsychotic_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Depression'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,4) IN ('N05A')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DP_Other_Taking FROM (";
            SQL = SQL + "SELECT p.Sex, m.Patient_ID FROM Medication m, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = m.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Depression'";
            SQL = SQL + " AND (m.StopDate IS NULL OR m.StopDate>=#" + sBeginDate + "#)";
            SQL = SQL + " AND LEFT(Code_calc,5) IN ('N06AG','N06AF')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, m.Patient_ID";
            SQL = SQL + ") AS PtMed GROUP BY Sex";
            SQLExec(SQL);
        }

        protected void GetLabScreenProc(string filter)
        {
            string SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_HBA1C_3MONTHS FROM (";
            SQL = SQL + "SELECT p.Sex, l.Patient_ID FROM Lab l, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = l.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND DATEDIFF('m',l.PerformedDate,#" + sCutOffDate + "#)<=3 AND l.Name_calc='HBA1C'";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, l.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_LIPID_1YEAR FROM (";
            SQL = SQL + "SELECT p.Sex, l.Patient_ID FROM Lab l, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = l.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND DATEDIFF('m',l.PerformedDate,#" + sCutOffDate + "#)<=12 AND l.Name_calc IN ('HDL','LDL','CALCULATED LDL','TRIGLYCERIDES','TOTAL CHOLESTEROL')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, l.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_MICRAL_CREAT_1YEAR FROM (";
            SQL = SQL + "SELECT p.Sex, l.Patient_ID FROM Lab l, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = l.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND DATEDIFF('m',l.PerformedDate,#" + sCutOffDate + "#)<=12 AND l.Name_calc IN ('URINE ALBUMIN CREATININE RATIO')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, l.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_FBS_1YEAR FROM (";
            SQL = SQL + "SELECT p.Sex, l.Patient_ID FROM Lab l, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = l.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND DATEDIFF('m',l.PerformedDate,#" + sCutOffDate + "#)<=12 AND l.Name_calc IN ('FASTING GLUCOSE')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, l.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_LIPID_1YEAR FROM (";
            SQL = SQL + "SELECT p.Sex, l.Patient_ID FROM Lab l, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = l.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND DATEDIFF('m',l.PerformedDate,#" + sCutOffDate + "#)<=12 AND l.Name_calc IN ('HDL','LDL','CALCULATED LDL','TRIGLYCERIDES','TOTAL CHOLESTEROL')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, l.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            //SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_MICRAL_CREAT_1YEAR FROM (";
            //SQL = SQL + "SELECT p.Sex, l.Patient_ID FROM LabResult l, DiseaseCase i, Patient p, PatientDemographic pd WHERE p.Patient_ID = pd.Patient_ID";
            //SQL = SQL + " AND p.Patient_ID = i.Patient_ID";
            //SQL = SQL + " AND p.Patient_ID = l.Patient_ID";
            //SQL = SQL + " AND i.DiseaseCase = 'Hypertension'";
            //SQL = SQL + " AND DATEDIFF('yyyy',l.TestDoneDate,#" + sCutOffDate + "#)<=1 AND l.LabTest IN ('URINE ALBUMIN CREATININE RATIO')";
            //SQL = SQL + filter;
            //SQL = SQL + " GROUP BY p.Sex, l.Patient_ID";
            //SQL = SQL + ") AS PtLab GROUP BY Sex";
            //SQLExec(SQL);            
        }

        protected void GetExamScreenProc(string filter)
        {
            string SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_BP_3MONTHS FROM (";
            SQL = SQL + "SELECT p.Sex, pe.Patient_ID FROM Exam pe, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND DATEDIFF('m',pe.DateCreated,#" + sCutOffDate + "#)<=3 AND (pe.Exam1 ='sBP (mmHg)' OR pe.Exam2 ='dBP (mmHg)')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, pe.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_WEIGHT_1YEAR FROM (";
            SQL = SQL + "SELECT p.Sex, pe.Patient_ID FROM Exam pe, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND DATEDIFF('m',pe.DateCreated,#" + sCutOffDate + "#)<=12 AND pe.Exam1 IN ('Weight (kg)')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, pe.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS DM_HEIGHT FROM (";
            SQL = SQL + "SELECT p.Sex, pe.Patient_ID FROM Exam pe, DiseaseCase i, " + ContactGroup; 
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus'";
            SQL = SQL + " AND pe.Exam1 IN ('Height (cm)')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, pe.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);
            
            SQL = "SELECT Sex, Format(AVG(DM_BMI_P),'#.0') AS DM_BMI FROM (";
            SQL = SQL + "SELECT p.Sex, CDBL(Result1_calc) AS DM_BMI_P FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + ", (SELECT Patient_ID, MAX(DateCreated) AS RecentDate FROM Exam WHERE Exam1='BMI (kg/m^2)' AND ISNUMERIC(Result1_calc) GROUP BY Patient_ID) pew";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pew.Patient_ID";
            SQL = SQL + " AND pew.RecentDate = pe.DateCreated";
            SQL = SQL + " AND (Year(#" + sCutOffDate + "#) - p.BirthYear)>=18";            
            SQL = SQL + " AND pe.Exam1='BMI (kg/m^2)'";
            SQL = SQL + " AND ISNUMERIC(Result1_calc)";
            SQL = SQL + filter;
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus') GROUP BY p.Sex";            
            SQLExecDbl(SQL);            

            SQL = "SELECT Format(AVG(DM_BMI_P),'#.0') AS Sum_DM_BMI FROM (";
            SQL = SQL + "SELECT p.Sex, CDBL(Result1_calc) AS DM_BMI_P FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + ", (SELECT Patient_ID, MAX(DateCreated) AS RecentDate FROM Exam WHERE Exam1='BMI (kg/m^2)' AND ISNUMERIC(Result1_calc) GROUP BY Patient_ID) pew";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pew.Patient_ID";
            SQL = SQL + " AND pew.RecentDate = pe.DateCreated";
            SQL = SQL + " AND (Year(#" + sCutOffDate + "#) - p.BirthYear)>=18";
            SQL = SQL + " AND pe.Exam1='BMI (kg/m^2)'";
            SQL = SQL + " AND ISNUMERIC(Result1_calc)";
            SQL = SQL + filter;
            SQL = SQL + " AND i.Disease = 'Diabetes Mellitus')";
            SQLExecDbl(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_BP_6MONTHS FROM (";
            SQL = SQL + "SELECT p.Sex, pe.Patient_ID FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND DATEDIFF('m',pe.DateCreated,#" + sCutOffDate + "#)<=6 AND (pe.Exam1 ='sBP (mmHg)' OR pe.Exam2 ='dBP (mmHg)')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, pe.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_WEIGHT_1YEAR FROM (";
            SQL = SQL + "SELECT p.Sex, pe.Patient_ID FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND DATEDIFF('m',pe.DateCreated,#" + sCutOffDate + "#)<=12 AND pe.Exam1 IN ('Weight (kg)')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, pe.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, COUNT(Patient_ID) AS HTN_HEIGHT FROM (";
            SQL = SQL + "SELECT p.Sex, pe.Patient_ID FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND i.Disease = 'Hypertension'";
            SQL = SQL + " AND pe.Exam1 IN ('Height (cm)')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, pe.Patient_ID";
            SQL = SQL + ") AS PtLab GROUP BY Sex";
            SQLExec(SQL);

            SQL = "SELECT Sex, Format(AVG(DM_BMI_P),'#.0') AS HTN_BMI FROM (";
            SQL = SQL + "SELECT p.Sex, CDBL(Result1_calc) AS DM_BMI_P FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + ", (SELECT Patient_ID, MAX(DateCreated) AS RecentDate FROM Exam WHERE Exam1='BMI (kg/m^2)' AND ISNUMERIC(Result1_calc) GROUP BY Patient_ID) pew";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pew.Patient_ID";
            SQL = SQL + " AND pew.RecentDate = pe.DateCreated";
            SQL = SQL + " AND (Year(#" + sCutOffDate + "#) - p.BirthYear)>=18";
            SQL = SQL + " AND pe.Exam1='BMI (kg/m^2)'";
            SQL = SQL + " AND ISNUMERIC(Result1_calc)";
            SQL = SQL + filter;
            SQL = SQL + " AND i.Disease = 'Hypertension') GROUP BY p.Sex";
            SQLExecDbl(SQL);                        

            SQL = "SELECT Format(AVG(DM_BMI_P),'#.0') AS Sum_HTN_BMI FROM (";
            SQL = SQL + "SELECT p.Sex, CDBL(Result1_calc) AS DM_BMI_P FROM Exam pe, DiseaseCase i, " + ContactGroup;
            SQL = SQL + ", (SELECT Patient_ID, MAX(DateCreated) AS RecentDate FROM Exam WHERE Exam1='BMI (kg/m^2)' AND ISNUMERIC(Result1_calc) GROUP BY Patient_ID) pew";
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pe.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = pew.Patient_ID";
            SQL = SQL + " AND pew.RecentDate = pe.DateCreated";
            SQL = SQL + " AND (Year(#" + sCutOffDate + "#) - p.BirthYear)>=18";
            SQL = SQL + " AND pe.Exam1='BMI (kg/m^2)'";
            SQL = SQL + " AND ISNUMERIC(Result1_calc)";
            SQL = SQL + filter;
            SQL = SQL + " AND i.Disease = 'Hypertension')";
            SQLExecDbl(SQL);
        }

        protected void GetSmokingStatus(string filter)
        {
            string SQL = "SELECT Sex, COUNT(Patient_ID) AS COPD_Smoking_Status FROM (";
            SQL = SQL + "SELECT p.Sex, r.Patient_ID FROM RiskFactor r, DiseaseCase i, " + ContactGroup;
            SQL = SQL + " WHERE p.Patient_ID = i.Patient_ID";
            SQL = SQL + " AND p.Patient_ID = r.Patient_ID";
            SQL = SQL + " AND i.Disease = 'COPD'";
            SQL = SQL + " AND r.Name_calc IN ('Smoking')";
            SQL = SQL + filter;
            SQL = SQL + " GROUP BY p.Sex, r.Patient_ID";
            SQL = SQL + ") AS PtSmoke GROUP BY Sex";
            SQLExec(SQL);
        }

        protected void SQLExec(string SQL)
        {
            Log(SQL, true);
            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    if (reader.FieldCount == 2)
                    {
                        if (!reader.IsDBNull(0))
                        {
                            dictDemoGroup.Add(reader.GetString(0) + '_' + reader.GetName(1), reader.GetInt32(1));
                        }
                    }
                    else if (reader.FieldCount == 3)
                    {
                        if (reader.GetString(1) != "HTN_Med_0")
                        {
                            if (!reader.IsDBNull(0))
                            {
                                dictDemoGroup.Add(reader.GetString(0) + '_' + reader.GetString(1), (int)reader.GetDouble(2));
                            }
                        }
                    }
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmd = null;
        }

        protected void SQLExecDbl(string SQL)
        {
            Log(SQL, true);
            OleDbCommand cmmd = new OleDbCommand(SQL, conn);
            try
            {
                OleDbDataReader reader = cmmd.ExecuteReader();
                while (reader.Read())
                {
                    if (reader.FieldCount == 2)
                    {
                        if (!reader.IsDBNull(0))
                        {
                            dictDemoStr.Add(reader.GetString(0) + '_' + reader.GetName(1), reader.GetString(1));
                        }
                    }
                    else if (reader.FieldCount == 1)
                    {
                        if (!reader.IsDBNull(0))
                        {
                            dictDemoStr.Add(reader.GetName(0), reader.GetString(0));
                        }
                        else
                        {
                            dictDemoStr.Add(reader.GetName(0), "N/A");
                        }
                    }
                }
                reader.Close();
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.StackTrace);
            }
            cmmd = null;
        }

        public void Close()
        {
            conn.Close();
        }

        protected void CalSum()
        {
            Dictionary<string, int> tempDict = new Dictionary<string, int>();
            foreach (string key in dictDemoGroup.Keys)
            {
                if (key.StartsWith("Male_") && key.IndexOf("BMI") < 0)
                {
                    string postfix = key.Substring("Male_".Length);
                    string femaleKey = "Female_" + postfix;
                    string sumKey = "Sum_" + postfix;
                    if (!tempDict.ContainsKey(sumKey))
                    {
                        int female_count = 0;
                        if (dictDemoGroup.ContainsKey(femaleKey))
                        {
                            female_count = dictDemoGroup[femaleKey];
                        }
                        int newValue = dictDemoGroup[key] + female_count;
                        tempDict.Add(sumKey, newValue);
                    }
                }

                if (key.StartsWith("Female_") && key.IndexOf("BMI") < 0)
                {
                    string postfix = key.Substring("Female_".Length);
                    string maleKey = "Male_" + postfix;
                    string sumKey = "Sum_" + postfix;
                    if (!tempDict.ContainsKey(sumKey))
                    {
                        int male_count = 0;
                        if (dictDemoGroup.ContainsKey(maleKey))
                        {
                            male_count = dictDemoGroup[maleKey];
                        }
                        int newValue = dictDemoGroup[key] + male_count;
                        tempDict.Add(sumKey, newValue);
                    }
                }
            }

            foreach (string key in tempDict.Keys)
            {
                dictDemoGroup.Add(key, tempDict[key]);
            }
        }

        protected void CalDenominator()
        {
            //DM
            if (dictDemoGroup.ContainsKey("Male_DM_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_DM", dictDemoGroup["Male_DM_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_DM", 0);

            if (dictDemoGroup.ContainsKey("Female_DM_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_DM", dictDemoGroup["Female_DM_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_DM", 0);

            if (dictDemoGroup.ContainsKey("Sum_DM_AllAge"))
            {
                dictDemoGroup.Add("DNT_SUM_DM", dictDemoGroup["Sum_DM_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_DM", 0);

            //HTN
            if (dictDemoGroup.ContainsKey("Male_HTN_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_HTN", dictDemoGroup["Male_HTN_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_HTN", 0);

            if (dictDemoGroup.ContainsKey("Female_HTN_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_HTN", dictDemoGroup["Female_HTN_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_HTN", 0);

            if (dictDemoGroup.ContainsKey("Sum_HTN_AllAge"))
            {                
                dictDemoGroup.Add("DNT_SUM_HTN", dictDemoGroup["Sum_HTN_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_HTN", 0);

            //COPD
            if (dictDemoGroup.ContainsKey("Male_COPD_AllAge"))
            {                
                dictDemoGroup.Add("DNT_MALE_COPD", dictDemoGroup["Male_COPD_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_COPD", 0);

            if (dictDemoGroup.ContainsKey("Female_COPD_AllAge"))
            {                
                dictDemoGroup.Add("DNT_FEMALE_COPD", dictDemoGroup["Female_COPD_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_COPD", 0);

            if (dictDemoGroup.ContainsKey("Sum_COPD_AllAge"))
            {
                dictDemoGroup.Add("DNT_SUM_COPD", dictDemoGroup["Sum_COPD_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_COPD", 0);

            //OA
            if (dictDemoGroup.ContainsKey("Male_OA_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_OA", dictDemoGroup["Male_OA_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_OA", 0);


            if (dictDemoGroup.ContainsKey("Female_OA_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_OA", dictDemoGroup["Female_OA_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_OA", 0);

            if (dictDemoGroup.ContainsKey("Sum_OA_AllAge"))
            {                
                dictDemoGroup.Add("DNT_SUM_OA", dictDemoGroup["Sum_OA_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_OA", 0);

            //Depression
            if (dictDemoGroup.ContainsKey("Male_DP_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_DP", dictDemoGroup["Male_DP_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_DP", 0);

            if (dictDemoGroup.ContainsKey("Female_DP_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_DP", dictDemoGroup["Female_DP_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_DP", 0);

            if (dictDemoGroup.ContainsKey("Sum_DP_AllAge"))
            {
                dictDemoGroup.Add("DNT_SUM_DP", dictDemoGroup["Sum_DP_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_DP", 0);

            //Epilepsy            
            if (dictDemoGroup.ContainsKey("Male_EPL_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_EPL", dictDemoGroup["Male_EPL_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_EPL", 0);

            if (dictDemoGroup.ContainsKey("Female_EPL_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_EPL", dictDemoGroup["Female_EPL_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_EPL", 0);

            if (dictDemoGroup.ContainsKey("Sum_EPL_AllAge"))
            {
                dictDemoGroup.Add("DNT_SUM_EPL", dictDemoGroup["Sum_EPL_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_EPL", 0);

            //Dementia            
            if (dictDemoGroup.ContainsKey("Male_DEM_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_DEM", dictDemoGroup["Male_DEM_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_DEM", 0);

            if (dictDemoGroup.ContainsKey("Female_DEM_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_DEM", dictDemoGroup["Female_DEM_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_DEM", 0);

            if (dictDemoGroup.ContainsKey("Sum_DEM_AllAge"))
            {
                dictDemoGroup.Add("DNT_SUM_DEM", dictDemoGroup["Sum_DEM_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_DEM", 0);

            //Parkinson's Disease            
            if (dictDemoGroup.ContainsKey("Male_PAK_AllAge"))
            {
                dictDemoGroup.Add("DNT_MALE_PAK", dictDemoGroup["Male_PAK_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_MALE_PAK", 0);

            if (dictDemoGroup.ContainsKey("Female_PAK_AllAge"))
            {
                dictDemoGroup.Add("DNT_FEMALE_PAK", dictDemoGroup["Female_PAK_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_FEMALE_PAK", 0);

            if (dictDemoGroup.ContainsKey("Sum_PAK_AllAge"))
            {
                dictDemoGroup.Add("DNT_SUM_PAK", dictDemoGroup["Sum_PAK_AllAge"]);
            }
            else
                dictDemoGroup.Add("DNT_SUM_PAK", 0);
        }

        protected void CalPct()
        {
            foreach (string key in dictDemoGroup.Keys)
            {
                string value = dictDemoGroup[key].ToString();
                if ((key.StartsWith("Male") || key.StartsWith("Female") || key.StartsWith("Sum")) && key.IndexOf("BMI") < 0)
                {
                    int pos = key.IndexOf('_');
                    pos = key.IndexOf('_', pos + 1);
                    if (pos >= 0)
                    {
                        string prefix = key.Substring(0, pos);
                        if (dictDemoGroup.ContainsKey("DNT_" + prefix.ToUpper()))
                        {
                            string pct;
                            if (dictDemoGroup["DNT_" + prefix.ToUpper()] == 0)
                                pct = 0.ToString("0.#") + "%";
                            else
                                pct= ((double)dictDemoGroup[key] / dictDemoGroup["DNT_" + prefix.ToUpper()] * 100).ToString("0.#") + "%";
                            value = value + "(" + pct + ")";
                        }
                    }
                }
                dictDemoStr.Add(key, value);
            }
        }

        protected void GetDataDrawnDate()
        {
            for (int i = 1; i <= 9; i++)
            {
                dictDemoStr.Add("Data_Drawn_Date" + i, recentDataDrawnDate);
            }

            for (int i = 1; i <= 8; i++)
            {
                dictDemoStr.Add("Last_Data_Drawn_Date" + i, lastDataDrawnDate);
            }
        }

        protected void GetPracticeOf(string raw_filter)
        {
            string reportTo = WordAccess.ConvertId(raw_filter);
            for (int i = 1; i <= 9; i++)
            {
                dictDemoStr.Add("Practice_Of_" + i, reportTo);
            }
        }


        public static void Log(string s)
        {
            Log(s, true);
        }

        public static void Log(string s, bool aIsLog)
        {
            if (aIsLog)
            {
                DateTime dtNow = DateTime.Now;
                StreamWriter log = new StreamWriter(new FileStream("CPCSSN_RPT.log", FileMode.Append));
                log.WriteLine(s);
                log.Close();
            }
        }
    }
}
