using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComponentsToolsWPF.Extensions {
    public static class WorkbookExtension {
        public static Range GetPartRange(this Workbook workbook, string partName) {
            Range range;
            foreach (Worksheet sheet in workbook.Sheets) {
                try {
                    range = sheet.GetVlaueRange(partName);
                }
                catch (Exception) {

                    throw;
                }

                if (range != null) {
                    return range;
                }
            }
            return null;
        }
        public static string[] UserGetShttesName(this Workbook workbook) {
            List<string> sheetsName = new List<string>();
            foreach (Worksheet sheet in workbook.Worksheets) {
                sheetsName.Add(sheet.Name);
            }
            return sheetsName.ToArray();
        }


    }
}
