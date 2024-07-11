using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ComponentsToolsWPF.Extensions {
    public static class RangeExtension {


        public static void UserMoveCells(this Range activeCells, Range destination) {
            activeCells.Cut();
            destination.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        }



        public static void UserAddLink(this Range range, string filePath) {
            Excel.Worksheet sheet = (Excel.Worksheet)range.Parent;
            sheet.Hyperlinks.Add(range, filePath, Type.Missing, range.Value, Type.Missing);
        }


        public static void UserSpin90(this Range range) {
            range.Orientation = 90;
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public static void UserCenter(this Range range) {
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public static void UserNoneLine(this Range range) {
            range.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlLineStyleNone;
        }
        public static void UserStandardLine(this Range range) {
            //range.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            range.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;

            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;

            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

            range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;

            range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
        }

        public static void UserMergeCells(this Range range) {
            //var mergeCells = range.MergeCells;
            //string sss = range.Address;
            Range _ = range.Resize[1, 1];
            if ((bool)_.MergeCells) {
                _.UnMerge();
            }
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            range.ReadingOrder = -5002;
            range.MergeCells = false;
            range.Merge();
        }


        public static void UserAddComboBox(this Range range, string[] list) {
            range.Validation.Delete();
            range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, string.Join(",", list), Type.Missing);
            range.Validation.IgnoreBlank = true;
            range.Validation.InCellDropdown = true;
        }

        public static void UserDeleteComboBox(this Range range) {
            range.Validation.Delete();
        }
    }
}
