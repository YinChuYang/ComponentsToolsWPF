using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ComponentsToolsWPF.ToolsPack {
    internal class FileSelect {


        public FileSelect() {

        }
        [STAThread]
        public string getFileSelectPath() {
            // 创建 OpenFileDialog 对象
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置对话框的标题
            openFileDialog.Title = "选择Excel文件用于存储参数";

            // 设置对话框的初始目录
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 设置对话框的过滤器，以限制可以选择的文件类型
            openFileDialog.Filter = "Excel (*.xls*)|*.xls*";

            // 是否允许选择多个文件
            openFileDialog.Multiselect = false;

            // 显示对话框，并检查用户是否点击了“确定”按钮
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                // 用户选择了文件，可以通过 openFileDialog.FileName 获取所选文件的完整路径
                return openFileDialog.FileName;
            }
            return "";
        }

    }
}
