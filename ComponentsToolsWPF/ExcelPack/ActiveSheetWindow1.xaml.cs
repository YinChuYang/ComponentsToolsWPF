using ComponentsToolsWPF.Extensions;
using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ComponentsToolsWPF.ExcelPack {
    /// <summary>
    /// ActiveSheetWindow1.xaml 的交互逻辑
    /// </summary>
    public partial class ActiveSheetWindow1 : System.Windows.Window
    {
        private string activeSheetName;
        string filePath;
        string fileName;
        public ActiveSheetWindow1() {
            InitializeComponent();
        }

        public string GetActiveSheetName() {
            return  activeSheetName;
        }

        public ActiveSheetWindow1(Workbook workbook,ModelDoc2 model) {
            InitializeComponent();
            //filePath = workbook.Path;
            //fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
            filePath = model.GetPathName();
            fileName = model.GetTitle();
            FilePathBox.Content = fileName;
            int N = 0;
            string[] sheetNames = workbook.UserGetShttesName();
            foreach (string sheetName in sheetNames) {
                RadioButton radioButton = new RadioButton();
                radioButton.Content = sheetName;
                radioButton.FontSize = 18;
                radioButton.MinHeight = 40;
                radioButton.FontWeight = FontWeights.Bold;
                radioButton.VerticalContentAlignment = VerticalAlignment.Center;
                radioButton.Padding = new Thickness(10, -1, 0, 0);
                radioButton.Margin = new Thickness(20, 10, 0, 0);
                radioButton.IsChecked = N == 0 ? true : false;
                SheetsNamesBox.Children.Add(radioButton);

                N++;
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e) {
            if (ActivesheetBox.Text != "") {
                activeSheetName = ActivesheetBox.Text;
            }
            else {
                foreach (RadioButton sheetNameBox in SheetsNamesBox.Children) {
                    if (sheetNameBox.IsChecked == true) {
                        activeSheetName = sheetNameBox.Content.ToString();
                    }
                }  
            }
            this.Close();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e) {

        }

        private void Label_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            FilePathBox.Content = (string)FilePathBox.Content == fileName ? filePath : fileName;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {

        }
    }
}
