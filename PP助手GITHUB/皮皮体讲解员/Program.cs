using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Threading.Tasks;
using System.Windows.Forms;
using IWshRuntimeLibrary;

namespace 皮皮助手
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                CreateShortcut();
                Application.Run(new FrmPP());

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }






        /// <summary>
        /// 创建快捷方式
        /// </summary>
        /// <param name="directory">快捷方式所处的文件夹</param>
        /// <param name="shortcutName">快捷方式名称</param>
        /// <param name="targetPath">目标路径</param>
        /// <param name="description">描述</param>
        /// <param name="iconLocation">图标路径，格式为"可执行文件或DLL路径, 图标编号"，
        /// 例如System.Environment.SystemDirectory + "\\" + "shell32.dll, 165"</param>
        /// <remarks></remarks>
        public static void CreateShortcut(string iconLocation = null)
        {

            string directory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            //if (System.IO.Directory.Exists(directory + "\\皮皮助手.lnk"))
            //{
            //    return;
            //}

            string shortcutPath = Path.Combine(directory, "皮皮助手.lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);//创建快捷方式对象
            shortcut.TargetPath = Application.ExecutablePath;//指定目标路径
            shortcut.WorkingDirectory = Application.StartupPath;//设置起始位置
            shortcut.WindowStyle = 1;//设置运行方式，默认为常规窗口
            shortcut.Save();//保存快捷方式
        }

      
    }
}
