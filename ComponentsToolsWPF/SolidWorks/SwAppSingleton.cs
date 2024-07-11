using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Xarial.XCad;
using Xarial.XCad.SolidWorks;

namespace ComponentsToolsWPF.SolidWorks {
    internal class SwAppSingleton {

        private static SldWorks _appInstance;
        private static readonly object _lockObject = new object();

        private SwAppSingleton() {

        }

        public static SldWorks GetSwApplication() {
            lock (_lockObject) {
                if (_appInstance == null) {
                    try {
                        // 尝试连接到已经运行的SolidWorks实例
                        _appInstance = (SldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application");
                        return _appInstance;
                    }
                    catch {
                        // 如果失败，创建新的SolidWorks实例
                        //_appInstance = new SldWorks();
                        //_appInstance.Visible = true; // 可见性可选，根据需要设置
                        return null ;
                    }
                }
                return _appInstance;
            }
        }

    }
}
