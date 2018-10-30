using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace NUnitSample.Configurations
{
    class Excel
    {
        String Path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        int scenariocolumnnumber;
        int speccustomercolumn;
        int testcasescolumn;
        ArrayList scenarionumers = new ArrayList();

       
        public void OpenExcelFile(String path)
        {
            this.Path = path;
            wb = excel.Workbooks.Open(Path);
        }

        public void OpenSheet(string Sheet)
        {
            ws = wb.Worksheets[Sheet];
        }

        public void SaveAndCloseExcel()
        {
            wb.Save();
            wb.Close();
            wb = null;
            ws = null;
            
        }
        public string GetInputData(string Sheet, string Scenario, string testcase, string inputcol)
        {
            ws=wb.Worksheets[Sheet];
            int testcaserownum = 0;
            scenariocolumnnumber = 0;
            testcasescolumn = 0;
            speccustomercolumn = 0;
            ArrayList scenarioname = new ArrayList();
            scenariocolumnnumber = GetColumnNumber("Scenario", "1");
            GetScenarios(Scenario);
            testcasescolumn = GetColumnNumber("Testcase", "1");
            speccustomercolumn = GetColumnNumber(inputcol, "1");

            for (int tcNum = 0; tcNum < scenarionumers.Count; tcNum++)
            {
                if (String.Equals(ws.Cells[scenarionumers[tcNum], testcasescolumn].Value, testcase, StringComparison.OrdinalIgnoreCase))
                {
                    testcaserownum = (int)scenarionumers[tcNum];
                    break;
                }
                if ((tcNum == scenarionumers.Count - 1) && (testcaserownum == 0))
                {
                    Console.WriteLine("Test case '" + testcase + "'  not found");
                }
            }
            return ws.Cells[testcaserownum, speccustomercolumn].Value;
        }
        public void WriteExcelFile(string Sheet, string Scenario, string testcase, string inputcol, string input)
        {
            ws=wb.Worksheets[Sheet];
            int testcaserownum = 0;
            scenariocolumnnumber = 0;
            testcasescolumn = 0;
            speccustomercolumn = 0;
            ArrayList scenarioname = new ArrayList();
            scenariocolumnnumber = GetColumnNumber("Scenario", "1");
            GetScenarios(Scenario);
            testcasescolumn = GetColumnNumber("Testcase", "1");
            speccustomercolumn = GetColumnNumber(inputcol, "1");
            for (int tcNum = 0; tcNum < scenarionumers.Count; tcNum++)
            {
                if (String.Equals(ws.Cells[scenarionumers[tcNum], testcasescolumn].Value, testcase, StringComparison.OrdinalIgnoreCase))
                {
                    testcaserownum = (int)scenarionumers[tcNum];
                    break;
                }
                if ((tcNum == scenarionumers.Count - 1) && (testcaserownum ==0))
                {
                    Console.WriteLine("Test case '" + testcase + "'  not found");
                }
            }             
            ws.Cells[testcaserownum, speccustomercolumn] = input;
            wb.Save();
        }

        public int GetColumnNumber(string colValue, string row)
        {
            Boolean foundcolumn = false;
            int columnnumber = 0;
            if (!(ws.UsedRange.Columns.Count > 0))
            {
                return columnnumber;
            }
            if (row == null || row == "")
            {
                row = "1";
            }
            for (int i = 1; i <= ws.UsedRange.Columns.Count; i++)
            {
                if (String.Equals(ws.Cells[row, i].Value, colValue, StringComparison.OrdinalIgnoreCase))
                {
                    foundcolumn = true;
                    columnnumber = i;
                }
                if ((i == ws.UsedRange.Columns.Count) && (foundcolumn == false))
                {
                    columnnumber = 0;
                }
            }
            if (columnnumber < 1)
            {
                Console.WriteLine(colValue + " not found in the file");
                columnnumber = 0;
            }
            return columnnumber;
        }


        public void GetScenarios(string Scenario)
        {
            scenarionumers.Clear();
            for (int i = 1; i <= ws.UsedRange.Rows.Count; i++)
            {
                if (String.Equals(ws.Cells[i, scenariocolumnnumber].Value, Scenario, StringComparison.OrdinalIgnoreCase))
                {
                    scenarionumers.Add(i);
                }
                if ((i == ws.UsedRange.Rows.Count) && ((scenarionumers.Count) < 1))
                {
                    Console.WriteLine("Scenario '" + Scenario + "' not found under scenario tab");
                }
            }
        }

        public static void KillExcelProcesses()
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }

    }
}
