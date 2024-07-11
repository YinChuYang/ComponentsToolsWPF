using ComponentsToolsWPF.Extensions;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentsToolsWPF.Pojo {
    internal class ComponentUserClass {
        public string configName { get; set; }
        public List<string> activeConfigName = new List<string> { };
        public List<string> componentNames = new List<string> { };
        public List<string[]> componentConfigurationNames = new List<string[]> { };

        public ComponentUserClass() { }

        public ComponentUserClass(string componentName, string[] customNames) {
            this.componentNames.Add(componentName);
            this.componentConfigurationNames.Add(customNames);
        }

        public void SetValue(Component2 component) {
            string _componentName = component.Name2;
            int _Index = _componentName.LastIndexOf('-');
            componentNames.Add("$配置@" + _componentName.Substring(0, _Index) + "<" + _componentName.Substring(_Index + 1) + ">");
            activeConfigName.Add(component.ReferencedConfiguration);
            ModelDoc2 modelDoc = component.GetModelDoc2() as ModelDoc2;
            componentConfigurationNames.Add(modelDoc.GetConfigurationNames() as string[]);
        }
        public void SetValue(object[,] arr, int row) {
            configName = arr[row, 1].UserToString();
            int s = arr.GetLength(1);
            for (int i = 2; i <= arr.GetLength(1); i++) {
                activeConfigName.Add(arr[row, i].UserToString());
                string _ = arr[(int)RangeCustomRows.ModelSizeNameRow, i].UserToString();
                if (_.IndexOf("$配置@") >= 0) {
                    _ = _.Substring(_.IndexOf("@") + 1);
                    _ = _.Substring(0, _.Length - 1);
                    _  = _.Replace("<", "-");
                }
               
                componentNames.Add(_);
            }


        }


    }
}
