using ComponentsToolsWPF.Extensions;
using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;

namespace ComponentsToolsWPF.Pojo {
    internal abstract class UserModelClassBase : IUserModelClassIntterface {

        protected string partID { get; set; }
        protected string name { get; set; }
        protected string filePath { get; set; }
        protected Range range { get; set; }
        protected string[] customNames { get; set; }
        protected List<string[]> modelConfigs { get; set; }
        protected Workbook workbook { get; set; }
        protected ModelDoc2 _model { get; set; }


        private bool amendModelSizeNameBool = false;

        protected UserModelClassBase() {
        }

        public abstract bool UpData();

        public abstract bool ReadData(ModelDoc2 swModel);


        protected void WritePartTitle(Worksheet sheet, int row) {
            sheet.Cells[RangeCustomRows.IDRow + row - 1, 1] = "零件ID：" + partID;
            ((Range)sheet.Cells[RangeCustomRows.IDRow + row - 1, 1]).UnMerge();
            ((Range)sheet.Cells[RangeCustomRows.IDRow + row - 1, 1]).Select();
            sheet.Cells[RangeCustomRows.PartNameRow + row - 1, 1] = "系列零件设计表是为：" + name;
            ((Range)sheet.Cells[RangeCustomRows.PartNameRow + row - 1, 1]).UnMerge();
            sheet.Cells[RangeCustomRows.PartPathRow + row - 1, 1] = "文件地址：" + filePath;
            ((Range)sheet.Cells[RangeCustomRows.PartPathRow + row - 1, 1]).UnMerge();
            ((Range)sheet.Cells[RangeCustomRows.PartPathRow + row - 1, 1]).UserAddLink(filePath);
            sheet.Cells[RangeCustomRows.CustomNameRow + row - 1, 1] = "自定义名称";
        }

        protected string[] GetModelSizeNameArray() {
            object[,] _ = ((Range)range.Rows[RangeCustomRows.ModelSizeNameRow]).Value2 as object[,];
            string[] sheetModelSizeName;
            if (_ == null) {
                return new string[] { "" };
            }
            else {
                sheetModelSizeName = new string[_.Length];
            }
            for (int i = 0; i < sheetModelSizeName.Length; i++) {
                if (_[1, i + 1] != null) {
                    sheetModelSizeName[i] = _[1, i + 1].ToString();
                }
                else {
                    sheetModelSizeName[i] = "";
                }
            }
            amendModelSizeNameBool = false;
            return sheetModelSizeName;
        }

        /// <summary>
        /// 获取设计表中的自定义名称
        /// </summary>
        /// <returns></returns>
        protected string ReadCustomSizeNames() {
            int row = range.Row;
            int count = range.Rows.Count;
            int column = range.Columns.Count;
            if (count < (int)RangeCustomRows.CustomNameRow) {
                return "";
            }
            string[] customNames = new string[column - 1];
            for (int i = 2; i <= column; i++) {
                customNames[i - 2] = ((Range)range.Cells[(int)RangeCustomRows.ModelSizeNameRow, i]).Value + ":"
                    + ((Range)range.Cells[(int)RangeCustomRows.CustomNameRow, i]).Value;
            }
            return string.Join("|", customNames);
        }

        /// <summary>
        /// 获取零件在设计表中的区域
        /// </summary>
        /// <param name="book"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        protected Range SetRanges(Workbook book, string id) {
            string partNameTItle = "零件ID：" + id;
            Range cells;
            try {
                cells = book.GetPartRange(partNameTItle);
            }
            catch (Exception) {

                throw;
            }

            if (cells == null) {
                return null;
            }

            return cells.CurrentRegion;
        }



    }
}