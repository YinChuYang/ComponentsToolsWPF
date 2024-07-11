using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentsToolsWPF.Extensions {
    public static  class AssemblyExtension {
        public static void UserAssemblyUpDataConfiguration(this AssemblyDoc swAssembly) {
            List<string> paramName = new List<string>();
            List<string> paramValue = new List<string>();
            Object components = swAssembly.GetComponents(true) as object[];

            foreach (Component2 component in (object[])components) {
                ModelDoc2 swModel = component.GetModelDoc2() as ModelDoc2;

                string _ = component.Name2;
                int _Index = _.LastIndexOf('-');
                _ = "$配置@" + _.Substring(0, _Index) + "<" + _.Substring(_Index + 1) + ">";
                paramName.Add(_);
                paramValue.Add(component.ReferencedConfiguration);

            }
            ModelDoc2 modelDoc2 = swAssembly as ModelDoc2;
            IConfiguration configuration = (IConfiguration)modelDoc2.GetActiveConfiguration();
            configuration.SetParameters(paramName, paramValue);
            modelDoc2.EditRebuild3();
        }
    }
}
