using ComponentsToolsWPF.ExcelPack;
using ComponentsToolsWPF.Extensions;
using ComponentsToolsWPF.Pojo;
using ComponentsToolsWPF.SolidWorks;
using ComponentsToolsWPF.ToolsPack;
using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Windows;

namespace ComponentsToolsWPF.UpDataDLL {


    public class UpDataService {
        SolidWorksDoc swDoc;
        WorkbookUserSetvice workbookClass;
        public UpDataService() {
            swDoc = new SolidWorksDoc();
            workbookClass = new WorkbookUserSetvice();
        }

        public bool OpenDesignSheet() {
            Workbook workbook;
            ModelDoc2 swModel = swDoc.GetActiveModel();
            string DesignPath = GetPartDesignPath(swModel);
            if (DesignPath == "") {
                return false;
            }
            workbook = workbookClass.GetWorkbook(DesignPath);
            string PartName = swModel.UserPartID();
            Range range = workbookClass.GetPartRegion(PartName, PartName, workbook);
            //string ss =range.Address;
            if (range == null) {
                Console.WriteLine("没有找到零件");
                return false;
            }
            range.Select();
            return true;
        }

        private string GetPartDesignPath(ModelDoc2 swModel) {

            string customPropertyValue = swDoc.GetCustomProperty("设计表地址", swModel);
            if (customPropertyValue == "") {
                //如果没有的话需要选择一个excel文件 并设置自定义属性
                FileSelect fileSelect = new FileSelect();
                customPropertyValue = fileSelect.getFileSelectPath();
                if (!swDoc.AddCustomProperty("设计表地址", customPropertyValue, swModel)) {
                    return "";
                }
            }

            return customPropertyValue;
        }

        #region 读取数据

        /// <summary>
        /// 下载数据到模型
        /// </summary>
        /// <returns></returns>
        public bool DownLoadDataToModel() {
            ModelDoc2 swModel = swDoc.GetActiveModel();

            if (swModel.GetSaveFlag()) {
                MessageBox.Show("请先保存文件");
                return false;
            }
            Workbook workbook = null;
            string excelPfth = swDoc.GetCustomProperty("设计表地址", swModel);
            if (excelPfth == "") {
                MessageBox.Show("该零件没有设计表, 请先上传到设计表");
                return false;
            }
            workbook = workbookClass.GetWorkbook(excelPfth);

            int componentType = swModel.GetType();

            Configuration currentConfiguration = swModel.GetActiveConfiguration() as Configuration;
            string[] names = (string[])swModel.GetConfigurationNames();
            swModel.ShowConfiguration2(names[names.Length - 1]);

            try {
                switch (componentType) {
                    case (int)swDocumentTypes_e.swDocPART:
                        //零件
                        PartDownloadLoadedSheet(swModel, workbook);
                        break;
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        //装配体
                        List<string> componentNames = new List<string>();
                        AssemblyDownloadLoadedSheet(swModel, workbook, componentNames);
                        UserModelClassBase userModelClassBase = new AssemblyUserClass(swModel, workbook);
                        if (!userModelClassBase.ReadData(swModel)) {
                            return false;
                        }
                        break;
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        //图纸
                        break;
                    default:
                        break;
                }
            }
            catch (Exception) {
                return false;
                throw;
            }
            swModel.ShowConfiguration2(currentConfiguration.Name);
            swModel.Save();
            return false;
        }

        private bool AssemblyDownloadLoadedSheet(ModelDoc2 swModel, Workbook workbook, List<string> componentNames) {
            AssemblyDoc swAssembly = swModel as AssemblyDoc;
            object[] components = swAssembly.GetComponents(true) as object[];
            Console.WriteLine(swModel.GetTitle());
            for (int i = components.Length - 1; i >= 0; i--) {
                Component2 item = components[i] as Component2;
                ModelDoc2 model = item.GetModelDoc2() as ModelDoc2;
                Console.WriteLine(item.Name2);

                if (item.IGetChildrenCount() <= 0) {
                    //零件
                    if (componentNames.Contains(model.GetTitle())) {
                        continue;
                    }
                    componentNames.Add(model.GetTitle());
                    PartDownloadLoadedSheet(model, workbook);
                }
                else {
                    //装配体
                    if (componentNames.Contains(model.GetTitle())) {
                        continue;
                    }
                    componentNames.Add(model.GetTitle());

                    AssemblyDownloadLoadedSheet(item.IGetModelDoc(), workbook, componentNames);
                    UserModelClassBase userModelClassBase = new AssemblyUserClass(swModel, workbook);
                    if (!userModelClassBase.ReadData(swModel)) {
                        continue;
                    }
                }

            }


            return true;
        }

        /// <summary>
        /// 下载零件数据  
        /// </summary>
        /// <param name="swModel"></param>
        /// <param name="workbook"></param>
        private void PartDownloadLoadedSheet(ModelDoc2 swModel, Workbook workbook) {
            try {
                PartUserClass componentUser = new PartUserClass(swModel, workbook);
                componentUser.ReadData(swModel);
            }
            catch (Exception) {
                throw;
            }
            return;
        }

        #endregion


        #region 上传数据

        /// <summary>
        /// 上传数据到设计表
        /// </summary>
        /// <returns></returns>
        public bool ModelDocUpLoadedSheet() {

            ModelDoc2 swModel = swDoc.GetActiveModel();

            if (swModel.GetSaveFlag()) {
                MessageBox.Show("请先保存文件");
                return false;
            }
            Workbook workbook = null;
            //读取自定义属性
            string excelPfth = swDoc.GetCustomProperty("设计表地址", swModel);
            if (excelPfth == "") {
                //如果没有的话需要选择一个excel文件 并设置自定义属性
                FileSelect fileSelect = new FileSelect();
                excelPfth = fileSelect.getFileSelectPath();
                if (excelPfth == "") {
                    return false;
                }
                swModel.UserAddCustomProperty("设计表地址", excelPfth);
            }
            workbook = workbookClass.GetWorkbook(excelPfth);
            int componentType = swModel.GetType();

            try {
                switch (componentType) {
                    case (int)swDocumentTypes_e.swDocPART:
                        //零件
                        PartConfigurationUpLoadedSheet(swModel, workbook);
                        break;
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        //装配体

                        List<string> componentNames = new List<string>();
                        UserModelClassBase userModelClassBase = new AssemblyUserClass(swModel, workbook);
                        if (!userModelClassBase.UpData()) {
                            return false;
                        }
                        AssemblyUpLoadedSheet(swModel as AssemblyDoc, workbook, ref componentNames);
                        break;
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        //图纸
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex) {

                return false;
                throw;
            }



            swModel.Save();
            return true;
        }

        private bool AssemblyUpLoadedSheet(AssemblyDoc swAssembly, Workbook workbook, ref List<string> componentNames) {
            object[] components = swAssembly.GetComponents(true) as object[];
            foreach (Component2 item in components) {
                ModelDoc2 swModel = (ModelDoc2)item.GetModelDoc2();
                Console.WriteLine(swModel.GetTitle());
                if (item.IGetChildrenCount() <= 0) {
                    Console.WriteLine(item.Name2 + "  为零件" + item.IGetChildrenCount());
                    if (componentNames.Contains(swModel.GetTitle())) {
                        continue;
                    }
                    swModel.UserAddCustomProperty("设计表地址", workbook.Path + "\\" + workbook.Name);
                    PartConfigurationUpLoadedSheet(swModel, workbook);
                    componentNames.Add(swModel.GetTitle());
                }
                else {
                    Console.WriteLine(item.Name2 + "  为装配体" + item.IGetChildrenCount());
                    if (componentNames.Contains(swModel.GetTitle())) {
                        continue;
                    }
                    swModel.UserAddCustomProperty("设计表地址", workbook.Path + "\\" + workbook.Name);
                    AssemblyConfigurationUpLoadedSheet(swModel, workbook);
                    componentNames.Add(swModel.GetTitle());
                    AssemblyUpLoadedSheet(swModel as AssemblyDoc, workbook, ref componentNames);
                }
            }

            return true;
        }



        /// <summary>
        /// 上传装配体数据
        /// </summary>
        /// <param name="swModel"></param>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private bool AssemblyUpLoadedSheet(ModelDoc2 swModel, Workbook workbook, string excelPath) {
            List<ModelDoc2> components = GetAllPart(swModel as AssemblyDoc);
            foreach (ModelDoc2 model in components) {
                model.UserAddCustomProperty("设计表地址", excelPath);
                PartConfigurationUpLoadedSheet(model, workbook);
            }
            return true;
        }
        /// <summary>
        /// 获取装配体中的所有零件
        /// </summary>
        /// <param name="swModel"></param>
        /// <returns></returns>
        private List<ModelDoc2> GetAllPart(AssemblyDoc swModel) {
            object[] components = swModel.GetComponents(false) as object[];
            List<ModelDoc2> reComponents = new List<ModelDoc2>();
            foreach (object item in components) {
                Component2 component = item as Component2;
                if (component != null) {
                    ModelDoc2 componentModel = (ModelDoc2)component.GetModelDoc2();
                    //Console.WriteLine(componentModel.GetPathName());
                    if (componentModel.GetType() == 1) {
                        if (!reComponents.Contains(componentModel)) {
                            reComponents.Add(componentModel);
                        }
                    }
                }
            }
            return reComponents;
        }

        /// <summary>
        /// 零件上传数据
        /// </summary>
        /// <param name="swModel"></param>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private bool PartConfigurationUpLoadedSheet(ModelDoc2 swModel, Workbook workbook) {
            swModel.UpDataConfiguration();
            try {
                PartUserClass componentUser = new PartUserClass(swModel, workbook);
                componentUser.UpData();
            }
            catch (Exception) {

                throw;
            }
            swModel.EditRebuild3();
            return true;
        }

        private bool AssemblyConfigurationUpLoadedSheet(ModelDoc2 swModel, Workbook workbook) {
            //swModel.UpDataConfiguration();
            //swModel.UserAssemblyUpDataConfiguration();
            try {
                UserModelClassBase componentUser = new AssemblyUserClass(swModel, workbook);
                componentUser.UpData();
            }
            catch (Exception) {

                throw;
            }
            swModel.EditRebuild3();
            return true;
        }

        private string AddPartID(ModelDoc2 swModel) {
            string componentID = swDoc.GetCustomProperty("零件ID", swModel);
            if (componentID == "") {
                Random random = new Random();
                componentID = random.Next(10000000, 99999999) + "ID";
                swDoc.AddCustomProperty("零件ID", componentID, swModel);
            }
            return componentID;
        }

        #endregion
    }
}
