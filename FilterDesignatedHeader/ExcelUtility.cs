using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using System.Runtime.InteropServices;
using System.Diagnostics;
//using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

namespace FilterDesignatedHeader
{
    public class ExcelUtility
    {
        public static IEnumerable<Excel.Application> GetExcelInterface()
        {
            try
            {
                //List<object> osObjectList = new List<object>();
                List<Excel.Application> osObjectList = new List<Excel.Application>();
                Hashtable addedFileNames = new Hashtable();

                Type type = Type.GetTypeFromProgID("Excel.Application");
                string clsid = type.GUID.ToString();
                string lookUpCandidateName = String.Format("!{0}{1}{2}", "{", clsid, "}").ToUpper();

                List<(string Name, object Value)> runningObjects = GetRunningObjectList();
                List<string> objName = new List<string>();
                Excel.Application excelObj = null;
                foreach (var (Name, Value) in runningObjects)
                {
                    objName.Add(Name);
                    string candidateName = Name.ToUpper();

                    if (candidateName.StartsWith(lookUpCandidateName))
                    //if (candidateName.EndsWith("XLSX"))
                    {
                        excelObj = Value as Excel.Application;
                        osObjectList.Add(excelObj);
                    }
                }
                return osObjectList;
            }
            catch (Exception)
            {
                return null;
            }
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable prot);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out System.Runtime.InteropServices.ComTypes.IBindCtx ppbc);
        public static List<(string Name, object Value)> GetRunningObjectList()
        {
            try
            {
                List<(string Name, object Value)> result = new List<(string Name, object Value)>();
                IntPtr numFetched = new IntPtr();
                System.Runtime.InteropServices.ComTypes.IMoniker[] monikers = new System.Runtime.InteropServices.ComTypes.IMoniker[1];
                GetRunningObjectTable(0, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable runningObjectTable);
                runningObjectTable.EnumRunning(out System.Runtime.InteropServices.ComTypes.IEnumMoniker monikerEnumerator);
                monikerEnumerator.Reset();
                while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
                {
                    CreateBindCtx(0, out System.Runtime.InteropServices.ComTypes.IBindCtx ctx);
                    monikers[0].GetDisplayName(ctx, null, out string runningObjectName);
                    //monikers[0].GetClassID(out Guid runningObjectPID);
                    runningObjectTable.GetObject(monikers[0], out object runningObjectVal);
                    result.Add((runningObjectName, runningObjectVal));
                }
                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }


        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        public void ClearExcelProcess()
        {
            var list = ExcelUtility.GetExcelInterface().ToArray();
            for (int i = 0; i < list.Length; i++)
            {
                Excel.Application excel = list[i];
                if (excel.Workbooks.Count < 1)
                {
                    excel.Visible = true;
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);

                    Process process = GetExcelProcess(excel);
                    if (process.ProcessName == "EXCEL")
                    {
                        process.Kill();
                    }
                }
            }
        }
    }
}
