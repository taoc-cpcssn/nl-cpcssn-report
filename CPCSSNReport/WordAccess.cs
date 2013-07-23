using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;

namespace CPCSSNReport
{
    class WordAccess
    {
        DataAccess data;
        string strTemplate;
        string strExcelTemplate;
        string strOutputPath;
        List<string> bookmarks;

        public WordAccess(string currentDB, string prevDB, string templateStr, string excelTemplateStr, string outputStr)
        {
            data = new DataAccess(currentDB, prevDB);
            data.LoadProviderInfo();
            strTemplate = templateStr;
            strExcelTemplate = excelTemplateStr;
            strOutputPath = outputStr;

            bookmarks = new List<string>();
            FileStream stream_in = new FileStream(@"bookmarks.txt", FileMode.Open);
            StreamReader reader = new StreamReader(stream_in);
            string line = reader.ReadLine();
            while (line != null && line.Trim().Length > 0)
            {
                if (line.IndexOf(':') >= 0)
                {
                    bookmarks.Add(line);
                }
                line = reader.ReadLine();
            }
            reader.Close();
        }

        public void CreateDocByProvider(string filter)
        {
            List<string> lstProvider = data.GetProviders();
            foreach (string s in lstProvider)
            {
                if (filter == "All" || filter == s)
                {
                    CreateDoc(s, "ByProvider");
                }
            }
        }

        public void CreateDocByGroup(HashSet<string> group)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in group)
            {
                if (s != "All")
                {
                    sb.Append(s + ",");
                }
            }
            sb.Remove(sb.Length - 1, 1);
            CreateDoc(sb.ToString(), "ByGroup");
        }

        public void CreateDocByPractice(string filter)
        {
            List<string> lstPractices = data.GetPractices();
            foreach (string s in lstPractices)
            {
                if (filter == "All" || filter == s)
                {
                    CreateDoc(s, "ByPractice");
                }
            }
        }

        public void CreateDocByAll()
        {
            CreateDoc("", "");
        }

        public void CreateSheetByProvider(string filter)
        {
            List<string> lstProvider = data.GetProviders();
            foreach (string s in lstProvider)
            {
                if (filter == "All" || filter == s)
                {
                    AppendSheet(s, "ByProvider");
                }
            }
        }

        public void CreateSheetByPractice(string filter)
        {
            List<string> lstPractices = data.GetPractices();
            foreach (string s in lstPractices)
            {
                if (filter == "All" || filter == s)
                {
                    AppendSheet(s, "ByPractice");
                }
            }
        }

        public void CreateSheetByGroup(HashSet<string> group)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in group)
            {
                if (s != "All")
                {
                    sb.Append("'" + s + "',");
                }
            }
            sb.Remove(sb.Length - 1, 1);
            AppendSheet(sb.ToString(), "ByGroup");
        }

        public void CreateSheetByAll()
        {
            AppendSheet("", "");
        }        

        public void CreateDoc(string filter, string type)
        {            
            string dbfilter, rpt_postfix;
            if (type == "ByProvider")
            {
                dbfilter = ByProvider(filter);
                rpt_postfix = "provider_" + filter;
            }
            else if (type == "ByPractice")
            {
                dbfilter = ByPractice(filter);
                rpt_postfix = "practice_" + filter;
            }
            else if (type == "ByGroup")
            {                
                dbfilter = ByGroup(filter);
                rpt_postfix = "group (" + filter.Replace("'", "") + ")";               
            }
            else
            {
                dbfilter = "";
                rpt_postfix = "all_summary";
            }

            data.LoadData(dbfilter, filter);

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;
            object oTemplate = strTemplate;
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
                ref oMissing, ref oMissing);

            List<string> lstKeyBookmark = data.GetKeys();
            //foreach (string s in lstKeyBookmark)
            //{
            //    object oMyAddress = s;
            //    if (s == "Data_Drawn_Date1")
            //    {
            //        Debug.Assert(1 == 1);
            //    }
            //    if (oDoc.Bookmarks.Exists(s))
            //    {
            //        Word.Bookmark oBookMark = oDoc.Bookmarks.get_Item(ref oMyAddress);
            //        oBookMark.Range.Text = data.GetValue(s);
            //    }
            //    else
            //    {
            //        Debug.Assert(1 == 1);
            //    }
            //}

            //FileStream stream_out = new FileStream(@"C:\bookmarks.txt", FileMode.Create);
            //StreamWriter writer = new StreamWriter(stream_out);

            foreach (Word.Bookmark oBookMark in oDoc.Bookmarks)
            {
                //writer.WriteLine(oBookMark.Name);
                if (lstKeyBookmark.Contains(oBookMark.Name))
                {
                    oBookMark.Range.Text = data.GetValue(oBookMark.Name);
                }
                else if (oBookMark.Name.IndexOf("BMI") < 0 && !oBookMark.Name.StartsWith("Female") && !oBookMark.Name.StartsWith("Male") && !oBookMark.Name.StartsWith("Sum"))
                {
                    oBookMark.Range.Text = "0";
                }
                else if (oBookMark.Name.IndexOf("BMI") < 0)
                {
                    oBookMark.Range.Text = "0(0%)";
                }
                else
                {
                    oBookMark.Range.Text = "N/A";
                }
            }
            //writer.Close();
            object strFilePath = strOutputPath + "\\Report_" + rpt_postfix;           
            object wfFormat = Word.WdSaveFormat.wdFormatDocument;
            oDoc.SaveAs(ref strFilePath, ref wfFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            //Close this form.            
            //this.Close();
        }

        public string ByProvider(string s)
        {
            return " AND p.Patient_ID IN (SELECT Patient_ID FROM PatientProvider WHERE Provider_ID=" + s + ")";
        }

        public string ByPractice(string s)
        {
            return " AND p.Site_ID =" + s;
        }

        public string ByGroup(string s)
        {
            return " AND p.Patient_ID IN (SELECT Patient_ID FROM PatientProvider WHERE Provider_ID IN (" + s + "))";
        }

        public List<string> GetProviders()
        {
            return data.GetProviders();
        }

        public List<string> GetPractices()
        {
            return data.GetPractices();
        }

        public void OpenBook()
        {
            if (oXL != null || oWB != null)
            {
                return;
            }
            try
            {                
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                object oTemplate = strExcelTemplate; //"c:\\Site Comparison Data (for auto poulation).xls";
                oWB = oXL.Workbooks.Add(oTemplate);
                //AppendSheet(oWB);
                //oWB.Save();
                //oXL.Quit();
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                throw theException;
            }

        }

        public void CloseBook()
        {
            object strFilePath = strOutputPath + "\\Data_all_summary";
            object oMissing = System.Reflection.Missing.Value;
            object xfFormat = Excel.XlFileFormat.xlXMLSpreadsheet;
            oWB.SaveAs(strFilePath,
                Excel.XlFileFormat.xlWorkbookDefault, oMissing, oMissing,
                false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                oMissing, oMissing, oMissing, oMissing, oMissing);            
            oWB.Close(false, oMissing, oMissing);
            oWB = null;
            oXL.Quit();
            oXL = null;
        }

        public void QuitBook()
        {
            object oMissing = System.Reflection.Missing.Value;
            if (oWB != null)
            {
                oWB.Close(false, oMissing, oMissing);
                oWB = null;
            }

            if (oXL != null)
            {
                oXL.Quit();
                oXL = null;
            }
        }

        public static string ConvertId(string id)
        {
            if (id.Length > 3)
            {
                string networkid = string.Format("0{0}", id.Substring(0, 1));
                string siteid = string.Format("{0}", id.Substring(1, 2));
                string providerid = string.Format("{0}", id.Substring(3, 2));
                return networkid + "-" + siteid + "-" + providerid;
            }
            else
            {
                return "07-0" + id;
            }
        }

        public void AppendSheet(string filter, string type)
        {            
            string dbfilter, rpt_postfix;
            if (type == "ByProvider")
            {
                dbfilter = ByProvider(filter);
                rpt_postfix = "provider_" + ConvertId(filter);
            }
            else if (type == "ByPractice")
            {
                dbfilter = ByPractice(filter);
                rpt_postfix = "practice_" + ConvertId(filter);
            }
            else if (type == "ByGroup")
            {
                dbfilter = ByGroup(filter);
                string title = filter.Replace("'", "");
                if (title.Length > 22)
                {
                    title = title.Substring(0, title.LastIndexOf(',', 22)) + "...";
                }
                rpt_postfix = "group (" + title + ")";                
            }
            else
            {
                dbfilter = "";
                rpt_postfix = "all_summary";
            }

            data.LoadData(dbfilter, filter);

            Excel._Worksheet oSheet;
            object oMissing = System.Reflection.Missing.Value;

            oSheet = (Excel._Worksheet)oWB.Sheets["Template"];            
            Excel._Worksheet newSheet = null;
            foreach (Excel._Worksheet tmpSheet in oWB.Sheets)
            {
                if (tmpSheet.Name == rpt_postfix)
                {
                    newSheet = tmpSheet;
                    break;
                }
            }
            if (newSheet == null)
            {
                oSheet.Copy(oMissing, oSheet);
                newSheet= (Excel._Worksheet)oWB.Sheets["Template (2)"];
                newSheet.Name = rpt_postfix;
            }            
            
            List<string> lstKeyBookmark = data.GetKeys();

            foreach (string bookmark in bookmarks)
            {                
                string[] ss = bookmark.Split(':');
                string value;
                if (lstKeyBookmark.Contains(ss[0]))
                {
                    value = data.GetValue(ss[0]);
                }
                else if (ss[0].IndexOf("BMI") < 0 && !ss[0].StartsWith("Female") && !ss[0].StartsWith("Male") && !ss[0].StartsWith("Sum"))
                {
                    value = "0";
                }
                else if (ss[0].IndexOf("BMI") < 0)
                {
                    value = "0(0%)";
                }
                else
                {
                    value = "N/A";
                }
     
                string[] vv = value.Split('(', ')');
                
                int x, y;
                GetCoordinate(ss[1], out x, out y);
                newSheet.Cells[x, y] = vv[0];
                if (vv[0] == "6")
                {
                    Debug.Assert(1 == 1);
                }

                if (ss.Length == 3 && vv.Length >= 2)
                {
                    GetCoordinate(ss[2], out x, out y);
                    newSheet.Cells[x, y] = vv[1];                    
                }                
            }
        }

        public void GetCoordinate(string s, out int x, out int y)
        {
            if (char.IsLetter(s[1]))
            {
                string y_v = s.Substring(0, 2).ToUpper();
                char c1 = y_v[0];
                char c2 = y_v[1];
                y = (c1 - 'A' + 1) * 26 + c2 - 'A' + 1;
                int.TryParse(s.Substring(2), out x);
            }
            else
            {
                string y_v = s.Substring(0, 1).ToUpper();
                char c1 = y_v[0];
                y = c1 - 'A' + 1;
                int.TryParse(s.Substring(1), out x);
            }
        }
       
        Excel._Workbook oWB;
        Excel.Application oXL;
    }    
}
