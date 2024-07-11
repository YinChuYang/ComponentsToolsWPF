using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentsToolsWPF.Extensions {
    public static class StringExtensions {
        public static string GetName(this string input, string separator) {
            // 获取第一个冒号之前的部分
            int index = input.IndexOf(separator);
            if (index != -1) {
                return input.Substring(0, index).Trim();
            }
            else {
                // 如果没有冒号，直接返回原始字符串
                return input;
            }
        }
        public static string GetValue(this string input, string separator) {
            // 获取第一个冒号之后的部分
            int index = input.IndexOf(separator);
            if (index != -1 && index + 1 < input.Length) {
                return input.Substring(index + 1).Trim();
            }
            else {
                // 如果没有冒号，返回空字符串
                return string.Empty;
            }
        }

        //为string添加扩展方法，输入分隔符，返回列表    示例"aaa|bbb|ccc"

        public static string[] GetCustomNames(this string input ,char separator) {
            return input.Split(separator);
        }

    }
}
