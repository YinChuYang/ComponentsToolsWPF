
using SolidWorks.Interop.sldworks;


namespace ComponentsToolsWPF.SolidWorks {
    public class SolidWorksDoc {
        //IXApplication swApp;
        SldWorks swApp;
        ModelDoc2 swModel;

        public SolidWorksDoc() {
            //swApp = SolidWorksAppSingleton.GetApplicationInstance();
            swApp = SwAppSingleton.GetSwApplication();
        }

        public ModelDoc2 GetActiveModel() {
            swModel = (ModelDoc2)swApp.IActiveDoc2;
            return swModel;
        }


        /// <summary>
        /// 获取自定义属性
        /// </summary>
        /// <param name="name"></param>
        /// <param name="swModel"></param>
        /// <returns></returns>
        public string GetCustomProperty(string name, ModelDoc2 swModel) {
            CustomPropertyManager swCustProp = default(CustomPropertyManager);
            CustomPropertyManager customPropertyManagers = swModel.Extension.CustomPropertyManager[""];
            string propValue = "";
            string propType = "";
            customPropertyManagers.Get4(name, false, out propValue, out propType);
            return propValue;
        }

        /// <summary>
        /// 添加自定义属性
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="swModel"></param>
        /// <returns></returns>
        public bool AddCustomProperty(string name, string value, ModelDoc2 swModel) {

            CustomPropertyManager customPropertyManagers = swModel.Extension.CustomPropertyManager[""];
            int states = customPropertyManagers.Add3(name, 30, value, 2);
            //0:成功, 1:失败, 2:有相同,不同值, 3:类型不匹配
            switch (states) {
                case 0:
                    //添加成功
                    return true;
                default:
                    //添加失败
                    return false;
            }
        }

        /// <summary>
        /// 重新建模
        /// </summary>
        /// <returns></returns>
        public bool UserEditRebuild() {
            return swModel.EditRebuild3();
        }

        /// <summary>
        /// 设置配置
        /// </summary>
        /// <param name="configName">配置名</param>
        /// <param name="paramNames">参数名数组</param>
        /// <param name="values">参数值数组</param>
        /// <param name="swModel">模型</param>
        /// <returns>bool</returns>
        public bool setConfiguration(string configName, string[] paramNames, string[] values, ModelDoc2 swModel) {
            return swModel.ConfigurationManager.SetConfigurationParams(configName, paramNames, values);
        }

        public ConfigurationManager GetConfigurationmMgr( ModelDoc2 swModel) {
            //获取配置管理器
            return swModel.ConfigurationManager;
        }

    }
}
