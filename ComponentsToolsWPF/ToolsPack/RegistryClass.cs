using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentsToolsWPF.ToolsPack {
    internal static class RegistryClass {
        public static void ClearPluginRegistry(string pluginName) {
            string registryPath = @"SOFTWARE\SolidWorks\AddIns\" + pluginName;

            try {
                Registry.LocalMachine.DeleteSubKeyTree(registryPath);
                Console.WriteLine("Plugin registry cleared successfully.");
            } catch (Exception ex) {
                Console.WriteLine("Error clearing plugin registry: " + ex.Message);
            }
        }
    }
}
