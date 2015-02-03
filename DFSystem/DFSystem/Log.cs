using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace DFSystem
{
    class Log
    {
         /* 记录日志 */
        public static void RecordLog(string errorText)
        {
            FileStream fs = new FileStream(@"C:\DFS\ErrorLog.txt", FileMode.Append, FileAccess.Write, System.IO.FileShare.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(DateTime.Now);
            sw.WriteLine(errorText);
            sw.Flush();
            sw.Close();
            fs.Close();
        }
    }
}
