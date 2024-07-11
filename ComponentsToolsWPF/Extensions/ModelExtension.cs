using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ComponentsToolsWPF.Extensions {
    public static class ModelExtension {
        public static string UserPartID(this ModelDoc2 swModel) {
            string componentID = swModel.UserGetPartID();
            if (componentID == "") {
                Random random = new Random();
                componentID = random.Next(10000000, 99999999) + "ID";
                CustomPropertyManager customPropertyManagers = swModel.Extension.CustomPropertyManager[""];
                int states = customPropertyManagers.Add3("零件ID", 30, componentID, 2);
                switch (states) {
                    case 0:
                        //添加成功
                        return componentID;
                    default:
                        //添加失败
                        Console.WriteLine("添加零件ID失败");
                        return "";
                }
            }
            return componentID;
        }




        public static int UserAddCustomProperty(this ModelDoc2 swModel, string propertyName, string propertyValue) {
            CustomPropertyManager customPropertyManagers = swModel.Extension.CustomPropertyManager[""];
            return customPropertyManagers.Add3(propertyName, 30, propertyValue, 2);
        }

        private static string UserGetPartID(this ModelDoc2 swModel) {
            CustomPropertyManager swCustProp;
            CustomPropertyManager customPropertyManagers = swModel.Extension.CustomPropertyManager[""];
            string propValue = "";
            string propType = "";
            customPropertyManagers.Get4("零件ID", false, out propValue, out propType);
            return propValue;
        }

        public static string[] GetCustomProperty(this ModelDoc2 swModel) {

            CustomPropertyManager swCustProp = default(CustomPropertyManager);
            CustomPropertyManager customPropertyManagers = swModel.Extension.CustomPropertyManager[""];
            string propValue = "";
            string propType = "";
            customPropertyManagers.Get4("参数别名", false, out propValue, out propType);
            return propValue.GetCustomNames('|');

        }




        public static void 获取所有特征(this ModelDoc2 swModel) {
            Feature swFeat;
            swFeat = (Feature)swModel.FirstFeature();

            while (swFeat != null) {
                Console.WriteLine(swFeat.Name);


                swFeat = swFeat.GetNextFeature() as Feature;
            }

        }


        public static ArrayList UpDataConfiguration(this ModelDoc2 swModel) {
            //Configuration swConf = (Configuration)swModel.GetActiveConfiguration();

            //获得文档的第一个特征
            Feature startFeat = (Feature)swModel.FirstFeature();
            ArrayList vFeats;

            while (startFeat != null) {
                string s = startFeat.Name;
                if (startFeat.Name == "原点") {
                    startFeat = (Feature)startFeat.GetNextFeature();
                    break;
                }
                startFeat = (Feature)startFeat.GetNextFeature();
            }

            vFeats = GetAllFeatures(startFeat);
            ArrayList vDispDims;
            vDispDims = GetAllDimensions(vFeats);

            string[] paramName = new string[vDispDims.Count];
            string[] paramValue = new string[vDispDims.Count];
            for (int i = 0; i < paramValue.Length; i++) {
                DisplayDimension swDispDim = (DisplayDimension)vDispDims[i];
                Dimension swDim = swDispDim.GetDimension2(0);
                double[] val = (double[])swDim.GetValue3(1, "");
                paramName[i] = swDim.GetNameForSelection();
                paramValue[i] = val[0].ToString();
            }
            IConfiguration configuration = (IConfiguration)swModel.GetActiveConfiguration();
            configuration.SetParameters(paramName, paramValue);
            swModel.EditRebuild3();
            return vFeats;
        }



        private static ArrayList GetAllDimensions(ArrayList vFeats) {
            ArrayList swDimsColl = new ArrayList();
            for (int i = 0; i < vFeats.Count; i++) {
                Feature swFeat = (Feature)vFeats[i];
                DisplayDimension swDispDim = (DisplayDimension)swFeat.GetFirstDisplayDimension();
                while (swDispDim != null) {
                    if (!swDimsColl.Contains(swDispDim)) {
                        swDimsColl.Add(swDispDim);
                    }
                    swDispDim = (DisplayDimension)swFeat.GetNextDisplayDimension(swDispDim);

                }
            }
            return swDimsColl;
        }

        /// <summary>
        /// 获取所有特征
        /// </summary>
        /// <param name="startFeat"></param>
        /// <returns></returns>
        private static ArrayList GetAllFeatures(Feature startFeat) {
            ArrayList swProcFeatsColl = new ArrayList();
            Feature swFeat = startFeat;

            while (swFeat != null) {
                if (swFeat.GetTypeName2() != "HistoryFolder") {
                    if (!swProcFeatsColl.Contains(swFeat)) {
                        swProcFeatsColl.Add(swFeat);
                    }
                    CollectAllSubFeatures(swFeat, swProcFeatsColl);
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }
            return swProcFeatsColl;
        }

        /// <summary>
        /// 获取所有子特征
        /// </summary>
        /// <param name="parentFeat"></param>
        /// <param name="procFeatsColl"></param>
        private static void CollectAllSubFeatures(Feature parentFeat, ArrayList procFeatsColl) {
            Feature swSubFeat;
            swSubFeat = (Feature)parentFeat.GetFirstSubFeature();
            while (swSubFeat != null) {
                if (!procFeatsColl.Contains(swSubFeat)) {
                    procFeatsColl.Add(swSubFeat);
                }
                CollectAllSubFeatures(swSubFeat, procFeatsColl);
                swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
            }
        }



        public static void UserAssemblyUpDataConfiguration(this ModelDoc2 swModel) {
            List<string> paramName = new List<string>();
            List<string> paramValue = new List<string>();
            AssemblyDoc swAssembly = swModel as AssemblyDoc;
            Object components = swAssembly.GetComponents(true) as object[];

            foreach (Component2 component in (object[])components) {
                string _ = component.Name2;
                int _Index = _.LastIndexOf('-');
                _ = "$配置@" + _.Substring(0, _Index) + "<" + _.Substring(_Index + 1) + ">";
                paramName.Add(_);
                paramValue.Add(component.ReferencedConfiguration);

            }

            ConfigurationManager configurationManager = swModel.ConfigurationManager;
            Configuration 当前配置 = swModel.GetActiveConfiguration()as Configuration;
            bool pass =  configurationManager.SetConfigurationParams(当前配置.Name, paramName.ToArray(), paramValue.ToArray());
            //IConfiguration configuration = (IConfiguration)swModel.GetActiveConfiguration();
            //configuration.SetParameters(paramName, paramValue);
            swModel.EditRebuild3();
        }


    }
}
