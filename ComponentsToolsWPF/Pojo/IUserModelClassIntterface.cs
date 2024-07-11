using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;

namespace ComponentsToolsWPF.Pojo {
    internal interface IUserModelClassIntterface {

        bool ReadData(ModelDoc2 swModel);
        bool UpData();
    }
}