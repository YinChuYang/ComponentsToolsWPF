using ComponentsToolsWPF.ExcelPack;
using ComponentsToolsWPF.Extensions;
using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using Xarial.XCad.Documents;
using Excel = Microsoft.Office.Interop.Excel;
namespace ComponentsToolsWPF.Pojo {
    enum RangeCustomRows : int {
        IDRow = 1,
        PartNameRow = 2,
        PartPathRow = 3,
        ModelSizeNameRow = 4,
        CustomNameRow = 5,
        ConfigNameRow = 6
    }

    internal class PartUserClass : UserModelClassBase {

        bool amendModelSizeNameBool = false;
        public PartUserClass() {
        }

        public PartUserClass(ModelDoc2 swModel, Workbook book) {
            this.partID = swModel.UserPartID();
            this.name = swModel.GetTitle();
            this.filePath = swModel.GetPathName();
            this.range = null;
            this.customNames = swModel.GetCustomProperty();
            this.workbook = book;
            this._model = swModel;
            Range cells;
            try {
                cells = SetRanges(book, this.partID);
            }
            catch (Exception) {

                throw;
            }
            if (cells != null) {
                range = cells;
            }
            SetModelConfig(swModel);
            Console.WriteLine(
                "  模型名: " + name 
                );
        }




        public override bool ReadData(ModelDoc2 swModel) {
            if (range == null) {
                //没有找到对应的表格
                MessageBox.Show("没有找到  >>" + swModel.GetTitle() + "<<  对应的设计表, 请先上传");
                return false;
            }

            //拿自定义名称
            string _customNames = ReadCustomSizeNames();
            swModel.UserAddCustomProperty("参数别名", _customNames);

            //拿到表格数据
            List<string[]> sheetConfigurations = ReadConfigurationData();
            //写入到模型中
            //Configuration currentConfiguration = swModel.GetActiveConfiguration() as Configuration;
            //string[] names = (string[])swModel.GetConfigurationNames();
            //swModel.ShowConfiguration2(names[names.Length - 1]);
            foreach (string[] configItem in sheetConfigurations) {
                string configName = configItem[0].GetValue(":");
                if (configName != "") {
                    //用表格中的配置名去查找模型中的配置
                    Configuration configuration = swModel.GetConfigurationByName(configName) as Configuration;
                    if (configuration == null) {
                        configuration = (Configuration)swModel.AddConfiguration3(configName, "", "", 256);
                    }
                    string[] paramsName = new string[configItem.Length - 1];
                    string[] paramsValue = new string[configItem.Length - 1];
                    if (verifyConfiguration(configuration, configItem)) {
                        for (int i = 1; i < configItem.Length; i++) {
                            paramsName[i - 1] = configItem[i].GetName(":");
                            paramsValue[i - 1] = configItem[i].GetValue(":");
                        }
                        configuration.SetParameters(paramsName, paramsValue);
                        if (swModel.ConfigurationManager.SetConfigurationParams(configName, paramsName, paramsValue)) {
                            Console.WriteLine("设置配置:" + configName + "-成功");
                        }
                        else {
                            Console.WriteLine("设置配置:" + configName + "失败");
                        }
                    }
                    else {
                        Console.WriteLine("设置配置:" + configName + "失败");
                    }
                }
            }

            return true;
        }

        private bool verifyConfiguration(Configuration currentConfiguration, string[] configItem) {
            int number = 0, index = 0;
            object _paramsName, _paramsValue;
            string[] paramsName, paramsValue;
            currentConfiguration.GetParameters(out _paramsName, out _paramsValue);
            if (_paramsName != null) {
                paramsName = _paramsName as string[];
                paramsValue = _paramsValue as string[];
            }
            else {
                paramsName = new string[0];
                paramsValue = new string[0];
            }
            double _int;
            string A, B;
            if (configItem.Length > 0) {
                for (int i = 0; i < paramsName.Length; i++) {
                    string str = ForEachParamsName(paramsName[i], configItem);
                    if (str == null) {
                        continue;
                    }
                    try {
                        if (str.GetName(":") == "$说明") {
                            index++;
                            number++;
                            continue;
                        }
                    }
                    catch (Exception) {
                        continue;
                    }

                    A = str.GetValue(":");
                    B = paramsValue[i];
                    bool pass;
                    pass = double.TryParse(A, out _int);
                    pass = double.TryParse(B, out _int);
                    if (double.TryParse(A, out _int) == double.TryParse(B, out _int)) {
                        index++;
                    }
                    else {
                        MessageBox.Show("配置名: ->" + configItem[0].GetValue(":") +
                            "<-   参数名: { " + str.GetName(":") + "}   参数值:  > " +
                            A + " <  与配置中的值类型不匹配");
                        return false;
                    }
                    number++;

                }
                if (index == number) {
                    return true;
                }
            }
            return false;
        }

        private string ForEachParamsName(string name, string[] paramsName) {
            for (int i = 0; i < paramsName.Length; i++) {
                string str = paramsName[i];
                if (paramsName[i].GetName(":").Equals(name)) {
                    return paramsName[i];
                }
            }
            return null;
        }


        private List<string[]> ReadConfigurationData() {

            int row = range.Row;
            int column = range.Columns.Count;
            Worksheet sheet = (Worksheet)range.Parent;
            int rowCount = range.Rows.Count;

            string[] modelSizeNames = GetModelSizeNameArray();

            List<string[]> sheetConfigs = new List<string[]>();
            string[] _;
            for (int i = row + (int)RangeCustomRows.ConfigNameRow - 1; i <= row + rowCount - 1; i++) {
                _ = new string[modelSizeNames.Length];
                for (int j = modelSizeNames.Length; j > 0; j--) {
                    Range temp = (Range)sheet.Cells[i, j];
                    object tempValue = temp.Value;
                    if (tempValue != null) {
                        _[j - 1] = modelSizeNames[j - 1] + ":" + tempValue.ToString();
                        if (j == 1) {
                            _[j - 1] = "配置名:" + tempValue.ToString();
                        }
                    }
                    else {
                        _[j - 1] = modelSizeNames[j - 1] + ":";
                    }
                }
                sheetConfigs.Add(_);
            }
            return sheetConfigs;
        }



        /// <summary>
        /// 写入数据到表格
        /// </summary>
        public override bool UpData() {
            if (range == null) {
                //选择添加到表格还是新建表格
                Console.WriteLine("没找到表格");
                ActiveSheetWindow1 activeSheetWindow = new ActiveSheetWindow1(workbook, _model);
                workbook.Application.WindowState = Excel.XlWindowState.xlMinimized;
                activeSheetWindow.ShowDialog();
                workbook.Application.WindowState = Excel.XlWindowState.xlNormal;
                string sheetName = activeSheetWindow.GetActiveSheetName();
                if (sheetName == "") {
                    return false;
                }
                Worksheet sheet;
                try {
                    sheet = workbook.Sheets[sheetName] as Worksheet;
                }
                catch {
                    sheet = (Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                    sheet.Name = sheetName;
                }
                int row = sheet.GetUpRow(65000, 1) + 19;
                range = sheet.Cells[row, 1] as Range;
                range.Select();
            }
            //写入数据
            workbook.Application.ScreenUpdating = false;
            try {
                PartWriteData();
            }
            catch (Exception) {

                throw;
            }
            finally {
                workbook.Application.ScreenUpdating = true;
            }

            return true;
        }


        private void PartWriteData() {
            Excel.Worksheet sheet = (Excel.Worksheet)range.Parent;
            int column = range.Columns.Count;
            int row = range.Row;

            WritePartTitle(sheet, row);

            string[] sheetModelSizeName;
            sheetModelSizeName = GetModelSizeNameArray();
            WriteCustomModelSizeName(sheetModelSizeName);

            string[][] data = modelConfigs.ToArray();

            int _ = sheet.GetEndColumn(row + (int)RangeCustomRows.ModelSizeNameRow - 1);
            int[] sizeNameArrTab = new int[_ - 1];

            range.UserDeleteComboBox();

            WritePartConfiguration(sheet, ref sheetModelSizeName, data, ref sizeNameArrTab);

            DeleteInvalidConfiguration(sheet, sizeNameArrTab);

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





        private int WritePartConfiguration(Worksheet sheet, ref string[] sheetModelSizeName, string[][] data, ref int[] sizeNameArrTab) {
            int column;
            int endRow;
            int row = range.Row;
            endRow = range.Rows.Count + row - 1;
            if (endRow < (int)RangeCustomRows.CustomNameRow + row - 1) { endRow = 5 + row - 1; }

            column = sheet.GetEndColumn(row + (int)RangeCustomRows.ModelSizeNameRow - 1);

            //找配置列表
            for (int i = 0; i < data.GetLength(0); i++) {
                string configName = data[i][data[i].GetLength(0) - 1].GetValue(":");
                bool pass = false;
                if (amendModelSizeNameBool) {
                    //修改表格后,尺寸名称可能有增加, 所以做个标记,表示表格有修改
                    sheetModelSizeName = GetModelSizeNameArray();
                }
                for (int j = (int)RangeCustomRows.ConfigNameRow + row - 1; j <= endRow; j++) {
                    object _object = default(object);
                    string sheetConfigname = "";
                    try {
                        _object = ((Range)sheet.Cells[j, 1]).Value;
                    }
                    catch (Exception) {
                        _object = "";
                    }
                    int _int;
                    double _double;
                    bool ss = _object is double;
                    if (_object is double) {
                        _double = (double)_object;
                        sheetConfigname = _double.ToString();
                    }
                    else {
                        sheetConfigname = _object as string;
                    }

                    if (sheetConfigname == configName) {
                        //找到了
                        pass = true;
                        WriteConfigsToSheet(((Range)sheet.Cells[j, 1]), data[i], sheetModelSizeName.ToList(), ref sizeNameArrTab);
                        break;
                    }
                }
                if (!pass) {
                    //没找到, 插入一行
                    endRow++;
                    sheet.InsertLine(endRow);
                    WriteConfigsToSheet((Range)sheet.Cells[endRow, 1], data[i], sheetModelSizeName.ToList(), ref sizeNameArrTab);
                    range = range.CurrentRegion;
                }
            }
            //调整格式

            return column;
        }

        private void DeleteInvalidConfiguration(Worksheet sheet, int[] sizeNameArrTab) {
            //删除从未使用过得尺寸名称
            int row = range.Row;
            int column = sheet.GetEndColumn(row + (int)RangeCustomRows.ModelSizeNameRow - 1);
            int endRow = range.Rows.Count + row - 1;
            int ModelSizeNameRow = (int)RangeCustomRows.ModelSizeNameRow + row - 1;
            for (int i = sizeNameArrTab.Length - 1; i > 0; i--) {
                if (sizeNameArrTab[i] == 0) {
                    ((Range)sheet.Range[sheet.Cells[ModelSizeNameRow, i + 1], sheet.Cells[endRow, i + 1]]).ClearContents();
                    Range activeCells = (Range)sheet.Range[sheet.Cells[ModelSizeNameRow, i + 2], sheet.Cells[endRow, column]];
                    Range destination = (Range)sheet.Range[sheet.Cells[ModelSizeNameRow, i + 1], sheet.Cells[endRow, i + 1]];
                    activeCells.UserMoveCells(destination);
                }
            }
        }




        /// <summary>
        /// 写入配置到表格中
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="data"></param>
        /// <param name="sheetModelSizeName"></param>
        private void WriteConfigsToSheet(Range cells, string[] data, List<string> sheetModelSizeName, ref int[] sizeNameArrTab) {
            Worksheet sheet = (Worksheet)range.Parent;
            int row = cells.Row;
            int titleRow = range.Row;
            int endColumn = sheet.GetEndColumn((int)RangeCustomRows.ModelSizeNameRow + titleRow - 1);
            amendModelSizeNameBool = false;
            for (int i = 0; i < data.Length - 1; i++) {
                string sizeName = data[i].GetName(":");
                int C = sheetModelSizeName.IndexOf(sizeName);
                //int C= Array.IndexOf(sheetModelSizeName, sizeName);
                if (C > -1) {
                    sheet.Cells[row, C + 1] = data[i].GetValue(":");
                    if (C < sizeNameArrTab.Length)
                        sizeNameArrTab[C] = 1;
                }
                else {
                    //插入一列
                    sheet.Cells[(int)RangeCustomRows.ModelSizeNameRow + titleRow - 1, endColumn + 1] = data[i].GetName(":");
                    endColumn = endColumn + 1;
                    sheet.Cells[row, endColumn] = data[i].GetValue(":");
                    amendModelSizeNameBool = true;
                }
            }
            sheet.Cells[row, 1] = data[data.Length - 1].GetValue(":");
            if (amendModelSizeNameBool) {
                range = range.CurrentRegion;
            }

        }
        //将data写入到表格中去


        /// <summary>
        /// 将自定义模型名写入到表格中
        /// </summary>
        /// <param name="sheetModelSizeName"></param>
        private void WriteCustomModelSizeName(string[] sheetModelSizeName) {
            for (int i = 0; i < customNames.Length; i++) {
                int C = Array.IndexOf(sheetModelSizeName, customNames[i].GetName(":"));
                if (C > 0) {
                    range.Cells[RangeCustomRows.CustomNameRow, C + 1] = customNames[i].GetValue(":");
                }
            }
        }
        /// <summary>
        /// 检索数组中的值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="Value"></param>
        /// <param name="array"></param>
        /// <returns></returns>
        private int RetrievalArray<T>(T Value, T[] array) {
            for (int i = 0; i < array.Length; i++) {
                if (Value.Equals(array[i])) {
                    return i;
                }
            }
            return -1;
        }


        /// <summary>
        /// add model config
        /// 添加模型配置
        /// </summary>
        /// <param name="keyValuePair"></param>
        public void addModelConfig(string[] keyValuePair) {
            modelConfigs.Add(keyValuePair);
        }


        /// <summary>
        /// 获取模型配置
        /// </summary>
        /// <param name="swModel"></param>
        public void SetModelConfig(ModelDoc2 swModel) {
            //获取配置管理器
            ConfigurationManager configMgr = swModel.ConfigurationManager;
            string[] vConfName = swModel.GetConfigurationNames() as string[];
            object vParamName;
            object vParamValue;
            modelConfigs = new List<string[]> { };
            for (int i = 0; i < vConfName.Length; i++) {

                configMgr.GetConfigurationParams(vConfName[i], out vParamName, out vParamValue);
                string[] paramName = (string[])vParamName;
                string[] paramValue = (string[])vParamValue;
                string[] KeyValuePairs = new string[paramName.Length + 1];
                int j;
                for (j = 0; j < paramName.Length; j++) {
                    // Console.WriteLine("配置名: " + vConfName[i] + " 参数名: " + paramName[j] + " 参数值: " + paramValue[j]);
                    KeyValuePairs[j] = paramName[j] + ":" + paramValue[j];
                }
                KeyValuePairs[j] = "配置名:" + vConfName[i];
                modelConfigs.Add(KeyValuePairs);

            }

        }


    }
}
