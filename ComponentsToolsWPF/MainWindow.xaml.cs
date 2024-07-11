using ComponentsToolsWPF.ExcelPack;
using ComponentsToolsWPF.SolidWorks;
using ComponentsToolsWPF.ToolsPack;
using ComponentsToolsWPF.UpDataDLL;
using Microsoft.Office.Interop.Excel;
using MsdevManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xarial.XCad.Documents;

namespace ComponentsToolsWPF {
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow() {
            InitializeComponent();
        }

        private void OpenExcelFileButton_Click(object sender, RoutedEventArgs e) {
            UpDataService upDataService = new UpDataService();
            upDataService.OpenDesignSheet();
        }

        private void Test_Click(object sender, RoutedEventArgs e) {

            //SolidWorksDoc worksDoc = new SolidWorksDoc();
            //IXDocument model = worksDoc.GetActiveModel();

            //WorkbookUserClass workbook = new WorkbookUserClass(@"F:\Desktop\测试表格.xlsx");
            //Worksheet sheet = workbook.getSheetByName("Sheet1");

            UpDataService upDataService = new UpDataService();
            //bool pass=  upDataService.ModelDocUpLoadedSheet();    //上传零件设计表
            /*bool pass = upDataService.DownLoadDataToModel();*/    //下载零件设计表
            Console.WriteLine("测试功能");

        }

        private void ClearPluginRegistry_Click(object sender, RoutedEventArgs e) {
            string pluginName = PluginNameBox.Text;
            if (pluginName == "") {
                return;
            } else {
                ToolsPack.RegistryClass.ClearPluginRegistry(pluginName);
            }
        }

        private void UpData_Click(object sender, RoutedEventArgs e) {
            UpDataService upDataService = new UpDataService();
            bool pass = upDataService.ModelDocUpLoadedSheet();
        }

        private void DownloadData_Click(object sender, RoutedEventArgs e) {
            UpDataService upDataService = new UpDataService();
            bool pass = upDataService.DownLoadDataToModel();
        }
    }
}
