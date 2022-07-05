using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace comTest
{
    class INI
    {
        #region "声明变量"

        /// <summary>
        /// 写入INI文件
        /// </summary>
        /// <param name="section">节点名称[如[TypeName]]</param>
        /// <param name="key">键</param>
        /// <param name="val">值</param>
        /// <param name="filepath">文件路径</param>
        /// <returns></returns>
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        /// <summary>
        /// 读取INI文件
        /// </summary>
        /// <param name="section">节点名称</param>
        /// <param name="key">键</param>
        /// <param name="def">值</param>
        /// <param name="retval">stringbulider对象</param>
        /// <param name="size">字节大小</param>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        private static string strFilePath = Application.StartupPath + "\\config.ini";//获取INI文件路径
        private static string strSec = ""; //INI文件名

        /// <summary>
        /// 自定义读取INI文件中的内容方法
        /// </summary>
        /// <param name="Section">键</param>
        /// <param name="key">值</param>
        /// <returns></returns>
        public static string GetValue(string key)
        {

            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString("config", key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        public static void WriteIni(string key, string value)
        {
            WritePrivateProfileString("config", key, value.Trim(), strFilePath);
        }
        #endregion
    }
}
