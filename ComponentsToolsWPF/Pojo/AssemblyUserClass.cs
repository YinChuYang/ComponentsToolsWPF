using ComponentsToolsWPF.ExcelPack;
using ComponentsToolsWPF.Extensions;
using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace ComponentsToolsWPF.Pojo {
    internal class AssemblyUserClass : UserModelClassBase {



        bool amendModelSizeNameBool = true;

        public AssemblyUserClass(ModelDoc2 swModel, Workbook workbook) {
            this.partID = swModel.UserPartID();
            this.name = swModel.GetTitle();
            this.filePath = swModel.GetPathName();
            this.workbook = workbook;
            _model = swModel;
            customNames = new string[] { "自定义名称" };
            try {
                range = SetRanges(workbook, this.partID);
            }
            catch (Exception) {
                throw;
            }
            modelConfigs = new List<string[]>();
            Console.WriteLine(
                "  模型名: " + name
                );
        }

        public override bool ReadData(ModelDoc2 swModel) {
            if (range == null) {
                //没有找到对应的表格
                MessageBox.Show("没有找到对应的表格,请先上传到设计表");
                return false;
            }
            string _customNames = ReadCustomSizeNames();
            swModel.UserAddCustomProperty("参数别名", _customNames);

            List<ComponentUserClass> componentsList = GetSheetConfigurationParams();
            bool pass = UpDataModelConfiguration(componentsList);

            return false;
        }

        private bool UpDataModelConfiguration(List<ComponentUserClass> componentsList) {
            AssemblyDoc assemblyDoc = _model as AssemblyDoc;
            string activeConfigurationName = ((Configuration)_model.GetActiveConfiguration()).Name;
            foreach (ComponentUserClass componentConfiguration in componentsList) {
                string configurationName = componentConfiguration.configName;
                if (activeConfigurationName != configurationName && !_model.ShowConfiguration2(configurationName)) {
                    //激活配置失败 创建配置
                    _model.AddConfiguration3(configurationName, "", "", 256);
                }
                object[] components = assemblyDoc.GetComponents(true) as object[];
                //Console.WriteLine(components.Length);
                foreach (Component2 component in components) {
                    string componentName = component.Name2;
                    int index = componentConfiguration.componentNames.IndexOf(componentName);
                    if (index > -1) {
                        component.ReferencedConfiguration = componentConfiguration.activeConfigName[index];
                    }
                }
            }
            //_model.ShowConfiguration2(activeConfigurationName);
            return true;
        }

        public override bool UpData() {
            try {
                workbook.Application.ScreenUpdating = false;
                //获取参数哈希表
                List<ComponentUserClass> configurations = GetModelComponentParams();
                //获取表格数据
                range = getModelRegion();
                if (range == null) {
                    return false;
                }
                range.Select();
                int row = range.Row;
                Worksheet sheet = (Worksheet)range.Parent;
                range.UserDeleteComboBox();
                WritePartTitle(sheet, row);
                List<string> sheetModelSizeName = GetModelSizeNameArray().ToList();
                //写入数据
                foreach (ComponentUserClass configuration in configurations) {
                    string configurationName = configuration.configName;
                    int configNameRow = GetConfigurationNameRow(configurationName, range);
                    sheet.Cells[configNameRow, 1] = configurationName;
                    //写入配置
                    for (int i = 0; i < configuration.componentNames.Count; i++) {
                        Debug.WriteLine(configuration.componentNames[i]);
                        int C = sheetModelSizeName.IndexOf(configuration.componentNames[i]);

                        if (C > -1) {
                            //找到了
                            sheet.Cells[configNameRow, C + 1] = configuration.activeConfigName[i];
                            ((Range)sheet.Cells[configNameRow, C + 1]).UserAddComboBox(configuration.componentConfigurationNames[i]);
                        }
                        else {
                            //没找到, 插入一列
                            int modelSizeNameEndColumn = sheet.GetEndColumn(row + (int)RangeCustomRows.ModelSizeNameRow - 1);
                            sheet.Cells[(int)RangeCustomRows.ModelSizeNameRow + range.Row - 1, modelSizeNameEndColumn + 1] = configuration.componentNames[i];
                            sheet.Cells[configNameRow, modelSizeNameEndColumn + 1] = configuration.activeConfigName[i];
                            //添加下拉框
                            ((Range)sheet.Cells[configNameRow, modelSizeNameEndColumn + 1]).UserAddComboBox(configuration.componentConfigurationNames[i]);
                            sheetModelSizeName.Add(configuration.componentNames[i]);
                            this.range = range.CurrentRegion;
                            //amendModelSizeNameBool = true;
                        }
                    }
                }
                range = range.CurrentRegion;

                //合并单元格
                ((Range)range.Rows[RangeCustomRows.PartNameRow]).UserMergeCells();
                ((Range)range.Rows[RangeCustomRows.PartPathRow]).UserMergeCells();
                ((Range)range.Rows[RangeCustomRows.IDRow]).UserMergeCells();
                ((Range)range.Rows[RangeCustomRows.ModelSizeNameRow]).UserSpin90();
                //range.UserNoneLine();

                //endRow = range.Rows.Count - (int)RangeCustomRows.CustomNameRow;
                for (int i = range.Row; i < range.Rows.Count + range.Row; i++) {
                    ((Range)sheet.Rows[i]).UserNoneLine();
                    if (i >= (int)RangeCustomRows.CustomNameRow + range.Row - 1) {
                        //居中
                        ((Range)sheet.Rows[i]).UserCenter();
                    }
                }
                range.UserStandardLine();
            }
            catch (Exception) {

                throw;
            }
            finally {
                workbook.Application.ScreenUpdating = true;
            }

            return true;
        }

        private int GetConfigurationNameRow(string configurationName, Range range) {
            int row = range.Row;
            int RowCount = range.Rows.Count;
            int endRow = RowCount + row - 1;
            Worksheet sheet = (Worksheet)range.Parent;
            if (RowCount <= 5) {
                endRow = row + (int)RangeCustomRows.ConfigNameRow - 1;
                sheet.InsertLine(endRow);
                sheet.Cells[endRow, 1] = configurationName;
                this.range = range.CurrentRegion;
                return endRow;
            }

            for (int j = (int)RangeCustomRows.ConfigNameRow + row - 1; j <= endRow; j++) {
                object _object = ((Range)sheet.Cells[j, 1]).Value;
                if (_object == null) {
                    continue;
                }
                if (configurationName == _object.ToString()) {
                    return j;
                }
            }
            //插入一行
            sheet.InsertLine(endRow + 1);
            sheet.Cells[endRow + 1, 1] = configurationName;
            range = range.CurrentRegion;
            return endRow + 1;
        }

        private Range getModelRegion() {
            //选择添加到表格还是新建表格
            if (range == null) {
                Console.WriteLine("没找到表格");
                ActiveSheetWindow1 activeSheetWindow = new ActiveSheetWindow1(workbook, _model);
                workbook.Application.WindowState = Excel.XlWindowState.xlMinimized;
                activeSheetWindow.ShowDialog();
                workbook.Application.WindowState = Excel.XlWindowState.xlNormal;
                string sheetName = activeSheetWindow.GetActiveSheetName();
                if (sheetName == null) {
                    return null;
                }
                Worksheet sheet;
                try {
                    sheet = workbook.Sheets[sheetName] as Worksheet;
                    sheet.Select();
                }
                catch {
                    sheet = (Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                    sheet.Select();
                    sheet.Name = sheetName;
                }
                int row = sheet.GetUpRow(65000, 1) + 19;

                range = sheet.Cells[row, 1] as Range;
                range.Select();
            }
            return range;
        }


        private List<ComponentUserClass> GetSheetConfigurationParams() {
            int row = range.Row;
            int rowCount = range.Rows.Count;
            int endRow = rowCount + row - 1;
            Worksheet sheet = (Worksheet)range.Parent;

            if (rowCount <= (int)RangeCustomRows.CustomNameRow) {
                return null;
            }
            object[,] arr = range.Value2 as object[,];
            if (arr == null) {
                return null;
            }
            List<ComponentUserClass> configurations = new List<ComponentUserClass>();
            for (int i = (int)RangeCustomRows.ConfigNameRow; i <= rowCount; i++) {
                ComponentUserClass component = new ComponentUserClass();
                component.SetValue(arr, i);
                configurations.Add(component);
            }
            return configurations;
        }


        /// <summary>
        /// 获取模型中的参数
        /// </summary>
        /// <returns></returns>
        private List<ComponentUserClass> GetModelComponentParams() {
            string activaConfigurationName = ((Configuration)_model.GetActiveConfiguration()).Name;
            //获取参数哈希表
            //优化的话, 可以暂停模型更新
            AssemblyDoc assemblyDoc = _model as AssemblyDoc;
            string[] configurationNames = _model.GetConfigurationNames() as string[];
            List<ComponentUserClass> configurations = new List<ComponentUserClass>();
            foreach (string configName in configurationNames) {
                //激活配置
                ComponentUserClass componentUser = new ComponentUserClass();
                componentUser.configName = configName;
                _model.ShowConfiguration2(configName);
                object[] components = assemblyDoc.GetComponents(true) as object[];
                foreach (Component2 component in components) {
                    componentUser.SetValue(component);
                }
                configurations.Add(componentUser);
            }
            _model.ShowConfiguration2(activaConfigurationName);
            return configurations;
        }

        private void WriteAssemblyConfiguration(Worksheet sheet, ref string[] sheetModelSizeName, string[][] data, ref int[] sizeNameArrTab) {
            int row = range.Row;
            int endRow;
            endRow = sheet.Rows.Count + row - 1;

            //获取零件列表
            string 是是 = _model.GetTitle();
            AssemblyDoc assemblyDoc = _model as AssemblyDoc;
            object[] components = assemblyDoc.GetComponents(true) as object[];

            if (endRow < (int)RangeCustomRows.CustomNameRow + row - 1) { endRow = 5 + row - 1; }


            for (int i = 0; i < data.GetLength(0); i++) {
                bool pass = false;

                string configName = data[i][data[i].GetLength(0) - 1].GetValue(":");

                if (amendModelSizeNameBool) {
                    //修改表格后,尺寸名称可能有增加, 所以做个标记,表示表格有修改
                    sheetModelSizeName = GetModelSizeNameArray();
                }

                for (int j = (int)RangeCustomRows.ConfigNameRow + row - 1; j <= endRow; j++) {
                    object _object = ((Range)sheet.Cells[j, 1]).Value;
                    string sheetConfigname = "";
                    if (_object is double) {
                        sheetConfigname = ((double)_object).ToString();
                    }
                    else {
                        sheetConfigname = _object as string;
                    }
                    if (sheetConfigname == configName) {
                        //找到了
                        pass = true;
                        WriteConfigsToSheet_Assembly((Range)sheet.Cells[j, 1], sheetModelSizeName.ToList(), ref sizeNameArrTab);
                        break;
                    }
                }
                if (!pass) {
                    //没找到, 插入一行
                    endRow++;
                    sheet.InsertLine(endRow);
                    WriteConfigsToSheet_Assembly((Range)sheet.Cells[endRow, 1], sheetModelSizeName.ToList(), ref sizeNameArrTab);
                    range = range.CurrentRegion;
                }

            }

        }

        private void WriteConfigsToSheet_Assembly(Range cells, List<string> sheetModelSizeName, ref int[] sizeNameArrTab) {
            Worksheet sheet = (Worksheet)cells.Parent;
            int column;
            int row = cells.Row;
            AssemblyDoc assemblyDoc = _model as AssemblyDoc;
            object[] components = assemblyDoc.GetComponents(true) as object[];
            for (int i = 0; i < components.Length; i++) {
                column = sheet.GetEndColumn(row);
                Component2 component = (Component2)components[i];
                ModelDoc2 modelDoc = component.GetModelDoc2() as ModelDoc2;
                string[] names = modelDoc.GetConfigurationNames() as string[];
                Console.WriteLine(component.Name2);
                string _ = component.Name2;
                int _Index = _.LastIndexOf('-');
                _ = "$配置@" + _.Substring(0, _Index) + "<" + _.Substring(_Index + 1) + ">";
                int C = sheetModelSizeName.IndexOf(_);

                if (C > -1) {
                    //找到了
                    //sheet.Cells[row, C+1] = component.ReferencedConfiguration;
                    ((Range)sheet.Cells[row, C + 1]).UserAddComboBox(names);
                    if (C < sizeNameArrTab.Length)
                        sizeNameArrTab[C] = 1;
                }
                else {
                    //没找到, 插入一列
                    sheet.Cells[(int)RangeCustomRows.ModelSizeNameRow + range.Row - 1, column + 1] = _;
                    sheet.Cells[row, column + 1] = component.ReferencedConfiguration;
                    //添加下拉框
                    ((Range)sheet.Cells[row, column + 1]).UserAddComboBox(names);
                    amendModelSizeNameBool = true;
                }

            }
            if (amendModelSizeNameBool) {
                range = range.CurrentRegion;
            }

        }

    }
}
