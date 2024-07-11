using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentsToolsWPF.Extensions {
    internal static class ObjectExtension {

        public static string UserToString(this object obj) {
            if (obj == null) {
                return "";
            }
            else {
                return obj.ToString();
            }
        }
        
    }
}
