using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace TicketFromEmail365
{
    class MyLogger
    {
        public MyLogger()
        {

        }

        public static void writeSingleLine(string line)
        {
            string singleLine = System.DateTime.Now.ToString() + "     " + line + System.Environment.NewLine;

            System.IO.File.AppendAllText(@"./Log.txt", singleLine);
        }

        public static void writeMultipleLines(string[] lines)
        {
            foreach (string s in lines)
            {
                string singleLine = System.DateTime.Now.ToString() + "     " + s + System.Environment.NewLine;

                System.IO.File.AppendAllText(@"./Log.txt", singleLine);
            }

        }

        public static bool checkLogFile()
        {
            if (!File.Exists(@"./Log.txt"))
            {
                try
                {
                    System.IO.File.AppendAllText(@"./Log.txt", "Date Time:                Event:" + System.Environment.NewLine);
                }
                catch
                {
                    return false;
                }
            }
            try
            {
                System.IO.File.AppendAllText(@"./Log.txt", "Log File Initialized" + System.Environment.NewLine);
                return true;
            }
            catch
            {
                return false;
            }
            
        }
    }
}
