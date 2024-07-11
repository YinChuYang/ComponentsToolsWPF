using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xarial.XCad.SolidWorks;
using Xarial.XCad;
using System.Windows.Forms;

namespace ComponentsToolsWPF.SolidWorks {
    internal class SolidWorksAppSingleton {

        private static IXApplication _appInstance;
        private static readonly object _lockObject = new object();

        private SolidWorksAppSingleton() { }

        public static IXApplication GetApplicationInstance() {
            lock (_lockObject) {
                if (_appInstance == null) {
                    // 如果实例不存在，则创建新的IXApplication实例
                    var process = System.Diagnostics.Process.GetProcessesByName("SLDWORKS");
                    if (!process.Any()) {
                        MessageBox.Show("请先打开solidworks", "提示");
                        return null;
                    }
                    _appInstance = SwApplicationFactory.FromProcess(process.First());
                }

                return _appInstance;
            }
        }

        public static void ReleaseApplicationInstance() {
            lock (_lockObject) {
                if (_appInstance != null) {
                    // 在应用程序结束时释放IXApplication实例
                    _appInstance.Close();
                    _appInstance = null;
                }
            }
        }
    }

}
