using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace 皮皮助手
{
    public static class XLog
    {
        static string path = System.Windows.Forms.Application.StartupPath;
        public static void LogPPT(string txt)
        {
            try
            {
                using (StreamWriter fs = new StreamWriter(path + "/ppt.log", true))
                {
                    fs.WriteLine(txt);
                    fs.Close();
                }
            }
            catch (Exception)
            {
                ;
            }
        }
        public static void LogSpk(string txt)
        {
            try
            {
                using (StreamWriter fs = new StreamWriter(path + "/spk.log", true))
                {
                    fs.WriteLine(txt);
                    fs.Close();
                }
            }
            catch (Exception)
            {
                ;
            }
        }
        public static void LogSystem(string txt)
        {
            try
            {
                using (StreamWriter fs = new StreamWriter(path + "/system.log", true))
                {
                    fs.WriteLine(txt);
                    fs.Close();
                }
            }
            catch (Exception)
            {
                ;
            }
        }
    }
}

