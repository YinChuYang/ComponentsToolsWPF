using ComponentsToolsWPF.Extensions;
using Microsoft.Office.Interop.Excel;
using MsdevManager;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
namespace ComponentsToolsWPF.ExcelPack {
    /// <summary>
    /// 工作簿类
    /// </summary>
    internal class WorkbookUserClass {

        private Excel.Application excelApp = null;
        Workbook workbook = null;
        public WorkbookUserClass() {

        }

        public WorkbookUserClass(string filePath) {
            if (excelApp == null) {
                workbook = GetExcelWorkbook(filePath);
                string name = workbook.Name;
            }
        }

        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="filePath"></param>
        public Workbook OpenExcelFile(string filePath) {
            if (excelApp == null) {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
            }
            return workbook = excelApp.Workbooks.Open(filePath);
        }
        public List<string> getsheetsName() {
            List<string> sheetsName = new List<string>();
            foreach (Worksheet sheet in workbook.Sheets) {
                sheetsName.Add(sheet.Name);
            }
            return sheetsName;
        }
        /// <summary>
        /// 获取一个区域的单元格
        /// </summary>
        /// <param name="R1"></param>
        /// <param name="C1"></param>
        /// <param name="R2"></param>
        /// <param name="C2"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public Range getRange(int R1, int C1, int R2, int C2, Worksheet sheet) {
            try {
                return sheet.Range[sheet.Cells[R1, C1], sheet.Cells[R2, C2]];
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public Worksheet getSheetByIndex(int index) {
            return (Worksheet)workbook.Worksheets[index];
        }
        public Worksheet getSheetByName(string sheetName) {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets) {
                if (worksheet.Name == sheetName) {
                    return worksheet;
                }
            }
            return null;
        }
        public string getWorkbookName() {
            return workbook.Name;
        }
        public string getFilePath() {
            return workbook.FullName;
        }
        public int getEndRow(Range range, Worksheet sheet) {
            return range.End[XlDirection.xlUp].Row;
        }
        /// <summary>
        /// 获取最后一行
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public int getlastRow(Worksheet sheet) {
            return sheet.GetLastRow();
        }
        public Worksheet getActiveSheet() {
            if (excelApp == null) {
                return null;
            }
            return (Worksheet)workbook.ActiveSheet;
        }
        public Workbook NewWorkbook() {
            return excelApp.Workbooks.Add();
        }




        #region 非业务代码


        /// <summary>
        /// 得到workbook, 有则返回, 无则创建
        /// 逻辑 : 
        /// 1. 获得所有的workbook的程序集
        /// 2. 如果没有, 则获得ExcelApp
        /// 3. 如果也没有excelapp, 则创建
        /// 4. 显示后打开文件
        /// 规避重复打开excelApp
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public Workbook GetExcelWorkbook(string filePath) {
            string fileName = Path.GetFileName(filePath);
            Hashtable runningObjects = Msdev.GetExcelInstances(false);
            IDictionaryEnumerator rotEnumerator = runningObjects.GetEnumerator();
            while (rotEnumerator.MoveNext()) {
                string 啊啊 = rotEnumerator.Key.ToString();
                if (filePath.Equals(啊啊)) {
                    workbook = (Workbook)rotEnumerator.Value;
                    excelApp = workbook.Application;
                    //IntPtr appHandle = getExcelAppHandle(filePath);
                    //Msdev.ShowExcel(workbook);
                    IntPtr appHandle = Msdev.getExcelWindowHandleByWindowName(fileName);
                    Msdev.ShowExcel(appHandle);
                    break;
                }
            }
            if (workbook == null) {

                try {
                    excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch (Exception) {
                }

                if (excelApp == null) {
                    excelApp = new Excel.Application();
                }
                excelApp.Visible = true;
                workbook = OpenExcelFile(filePath);
            }
            return workbook;
        }


        #endregion
    }
}
