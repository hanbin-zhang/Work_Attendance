using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorkAttendance
{
    public static class Comm
    {
        public static string ConnString = new StreamReader(@"paths\connectionString.txt").ReadLine();

        public static void WriteTextLog(string source, string strMessage)
        {
            string path = System.IO.Path.Combine(Application.StartupPath , @"Logs\");
            try
            {
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);

                DateTime time = DateTime.Now;

                string fileFullPath = path + "WorkAttendance." + time.ToString("yyyy-MM-dd") + ".txt";
                StringBuilder str = new StringBuilder();
                str.Append("Time:    " + time.ToString() + "\r\n");
                str.Append("Source:  " + source + "\r\n");
                str.Append("Message: " + strMessage + "\r\n");
                str.Append("-----------------------------------------------------------\r\n\r\n");
                System.IO.StreamWriter sw;
                if (!File.Exists(fileFullPath))
                {
                    sw = File.CreateText(fileFullPath);
                }
                else
                {
                    sw = File.AppendText(fileFullPath);
                }
                sw.WriteLine(str.ToString());
                sw.Close();
            }
            catch { }
        }
    }


}
