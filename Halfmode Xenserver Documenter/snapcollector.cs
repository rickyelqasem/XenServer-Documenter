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
    class snapcollector
    {
        public object snapcollect(Session session, List<XenRef<VM>> vmRefs)
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            int sn = 1;
            string log;
            string entry;

            try
            {
                
                 

                foreach (XenRef<VM> vmRef in vmRefs)
                {

                    VM vm = VM.get_record(session, vmRef);
                    if (!vm.is_a_template && !vm.is_control_domain)
                    {
                        vmc.Add("Virtual Machine Name:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.name_label));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " Snapshot Collection Error " + sn;
                            writelog.entry(log, entry);
                            sn++;
                        }
                        vmc.Add("Number Snapshots:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.snapshots.Count));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " Snapshot Collection Error " + sn;
                            writelog.entry(log, entry);
                            sn++;
                        }
                        if (vm.snapshots.Count >= 1)
                        {
                            for (int snapcount = 0; snapcount <= vm.snapshots.Count - 1; )
                            {
                                string snapref = Convert.ToString(vm.snapshots[snapcount].ServerOpaqueRef);

                                List<XenRef<VM>> snpRefs = VM.get_all(session);

                                foreach (XenRef<VM> snpRef in snpRefs)
                                {
                                    VM snapvm = VM.get_record(session, snpRef);
                                    if (snpRef.ServerOpaqueRef == vm.snapshots[snapcount].ServerOpaqueRef)
                                    {
                                        vmc.Add("Snapshot Name:");
                                        try
                                        {
                                            vmc.Add("Not displayed in demo version");
                                        }
                                        catch
                                        {
                                            vmc.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " Snapshot Collection Error " + sn;
                                            writelog.entry(log, entry);
                                            sn++;
                                        }
                                    }
                                }
                                snapcount++;

                            }

                        }


                    }
                    sn = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Snapshot Collection Finished";
                writelog.entry(log, entry);
                
                
            }
            catch
            {

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Snapshot Collection Falied";
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
