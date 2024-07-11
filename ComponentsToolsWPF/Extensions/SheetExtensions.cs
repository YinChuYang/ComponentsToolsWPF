using Microsoft.Office.Interop.Excel;
using System;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComponentsToolsWPF.Extensions {
    public static class SheetExtensions {

        public static int GetLastRow(this Excel.Worksheet sheet) {
            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            return lastRow;
        }

        /// <summary>
        /// 查找值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="value">单元格值,需要精确</param>
        /// <returns></returns>
        public static Range GetVlaueRange(this Worksheet sheet, string value) {
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
                false, Type.Missing, Type.Missing);

            if (foundCells != null) {


                try {
                    sheet.Select();
                }
                catch (Exception) {
                    MessageBox.Show("Excel表格在编辑状态, 请先退出编辑状态后再继续");
                    throw;
                }

                
                foundCells.Select();
            }
            return foundCells;
        }

        public static void InsertLine(this Excel.Worksheet sheet, int row) {
            Range range = (Range)(sheet.Cells[row, 1]);
            range.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        }

        public static int GetUpRow(this Excel.Worksheet sheet, int row, int cloumn) {
            Range range = (Range)(sheet.Cells[row, cloumn]);
            return range.End[XlDirection.xlUp].Row;
        }

        public static int GetDownRow(this Excel.Worksheet sheet, int row, int cloumn) {
            Range range = (Range)(sheet.Cells[row, cloumn]);
            return range.End[XlDirection.xlDown].Row;
        }

        public static int GetEndColumn(this Excel.Worksheet sheet, int row) {
            Range range = (Range)(sheet.Cells[row, 200]);
            return range.End[XlDirection.xlToLeft].Column;
        }


    }
}
