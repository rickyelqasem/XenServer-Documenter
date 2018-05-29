using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Halfmode_Xenserver_Documenter
{
    class writelog
    {
        public static void entry(string log, string evententry)
        {
            bool logentry = false;

            StreamWriter SW;

            while (!logentry)
            {
                try
                {
                    SW = File.AppendText(log);
                    SW.WriteLine(evententry);
                    SW.Close();
                    logentry = true;
                }
                catch
                {
                }
            }
        }
    }
}
