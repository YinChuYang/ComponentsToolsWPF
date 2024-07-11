using ComponentsToolsWPF.ExcelPack;
using ComponentsToolsWPF.Extensions;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComponentsToolsWPF.UpDataDLL {
    internal class WorkbookUserSetvice {

        WorkbookUserClass workbookClass;



        public WorkbookUserSetvice() {
            workbookClass = new WorkbookUserClass();
        }

        /// <summary>
        /// 获取零件在工作表的区域
        /// </summary>
        public Range GetPartRegion(string partName, string PartID, Workbook workbook) {
            Range range = GetPartRange(partName, workbook);
            if (range != null) {
                return range.CurrentRegion;
            }
            return null;
        }


        private Range GetPartRange(string partName, Workbook workbook) {
            Range range;
            foreach (Worksheet sheet in workbook.Sheets) {
                range = GetVlaueRange("零件ID：" + partName, sheet);
                if (range != null) {
                    return range;
                }
            }
            return null;
        }

        /// <summary>
        /// 查找值
        /// </summary>
        /// <param name="value"></param>
        /// <param name="sheet"></param>
        /// <returns> 单元格</returns>
        public Range GetVlaueRange(string value, Worksheet sheet) {
            Excel.Range foundCells;
            Excel.Range searchRange = sheet.UsedRange;

            // 查找值
            foundCells = searchRange.Find(
                value,
                Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole,
                Excel.XlSearchOrder.xlByColumns,
                Excel.XlSearchDirection.xlNext,
                false, Type.Missing, Type.Missing) ;

            if (foundCells != null) {
                sheet.Select();
                foundCells.Select();
            }
            return foundCells;
        }
        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public Workbook GetWorkbook(string filePath) {
            return workbookClass.GetExcelWorkbook(filePath);
        }


        /// <summary>
        ///     获取一个区域的单元格
        /// </summary>
        /// <param name="R1"></param>
        /// <param name="C1"></param>
        /// <param name="R2"></param>
        /// <param name="C2"></param>
        /// <param name="sheet"></param>
        public void getRegion(int R1, int C1, int R2, int C2, Worksheet sheet) {
            workbookClass.OpenExcelFile(@"F:\Desktop\测试表格.xlsx");
            var arr = workbookClass.getRange(R1, C1, R2, C2, sheet);
            int aa = workbookClass.getlastRow(sheet);
        }


    }
}
