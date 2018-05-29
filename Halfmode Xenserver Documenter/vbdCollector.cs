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
    class vbdCollector
    {

        public object vbdcollect(Session session, List<XenRef<VM>> vmRefs)
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            int te = 1;
            string log;
            string entry;

            try
            {
                              


                foreach (XenRef<VM> vmRef in vmRefs)
                {

                    VM vm = VM.get_record(session, vmRef);
                    if (!vm.is_a_template && !vm.is_control_domain)
                    {
                        List<XenRef<VBD>> vbdRefs = VBD.get_all(session);
                        vmc.Add("Virtual Machine Name:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.name_label));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Error " + te;
                            writelog.entry(log, entry);
                            te++;
                        }
                        vmc.Add("Number of VBD devices:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.VBDs.Count));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Error " + te;
                            writelog.entry(log, entry);
                            te++;
                        }
                        try
                        {
                            for (int i = 0; i <= vm.VBDs.Count - 1; )
                            {
                                if (vm.VBDs[i].ServerOpaqueRef == "OpaqueRef:NULL")
                                {
                                }
                                else
                                {
                                    VBD vbd = VBD.get_record(session, (String)vm.VBDs[i].ServerOpaqueRef);

                                    vmc.Add("Device Name:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vbd.device));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Error " + te;
                                        writelog.entry(log, entry);
                                        te++;
                                    }
                                    vmc.Add("Device is currently attached:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vbd.currently_attached));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Error " + te;
                                        writelog.entry(log, entry);
                                        te++;
                                    }
                                    vmc.Add("Device Read/Write Mode:");
                                    try
                                    {
                                        vmc.Add("Not displayed in demo version");
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Error " + te;
                                        writelog.entry(log, entry);
                                        te++;
                                    }
                                    vmc.Add("Device Type:");
                                    try
                                    {
                                        vmc.Add("Not displayed in demo version");
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Error " + te;
                                        writelog.entry(log, entry);
                                        te++;
                                    }
                                    i++;
                                }
                            }
                        }
                            catch
                        {
                            }
                        
                    }
                    te=1;
                }

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Finished";
                writelog.entry(log, entry);
                
                

            }


            catch
            {

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " VBD Collection Failed";
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
