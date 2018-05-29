using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using XenAPI;
using System.Reflection;
using System.Diagnostics;
using System.Collections;


namespace Halfmode_Xenserver_Documenter
{
    class templateCollector
    {
        public Object templatecollect(Session session, List<XenRef<VM>> vmRefs)
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string log;
            string entry;
            int tp = 1;
            try
            {
                



                foreach (XenRef<VM> vmRef in vmRefs)
                {

                    VM vm = VM.get_record(session, vmRef);

                    if (vm.is_a_template)
                    {
                        vmc.Add("Template Name:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.name_label));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " Template Collection Error " + tp;
                            writelog.entry(log, entry);
                            tp++;
                        }
                        vmc.Add("Template Description:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.name_description));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " Template Collection Error " + tp;
                            writelog.entry(log, entry);
                            tp++;
                        }
                        
                    }
                    tp = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Template Collection Finished ";
                writelog.entry(log, entry);
                
                
            }
            catch
            {

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Template Collection Failed";
                writelog.entry(log, entry);
                
                
            }
            if ((vmc.Count & 2) == 0)
            {
            }
            else
            {
                vmc.Add(" ");
            }
            return vmc;
        }
    }
}
