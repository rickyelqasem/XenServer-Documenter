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
    class VMCollector
    {
        public Object vmcollect(Session session, List<XenRef<VM>> vmRefs)
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            int vme = 1;
            string log;
            string entry;

            try
            {
                int vmcount = 0;



                foreach (XenRef<VM> vmRef in vmRefs)
                {


                    VM vm = VM.get_record(session, vmRef);

                    string tempdecr = vm.name_description;

                    //do not list templates or controller domain
                    if (!vm.is_a_template && !vm.is_control_domain)
                    {
                        vmcount = vmcount + 1;
                        vmc.Add("Virtual Machine Name:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.name_label));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("Decription:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.name_description));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("VM Power State:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.power_state));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("Amount of VCPUs:");
                        try
                        {
                            vmc.Add("Not displayed in demo version");
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("Max Amount of Memory:");
                        try
                        {
                            long maxmem = vm.memory_static_max;
                            maxmem = maxmem / 1048576;
                            vmc.Add("Not displayed in demo version");
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("Min Amount of Memory:");
                        try
                        {
                            long minmem = vm.memory_static_min;
                            minmem = minmem / 1048576;
                            vmc.Add(Convert.ToString(minmem));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        VM_metrics ms = VM_metrics.get_record(session, vm.metrics);
                        vmc.Add("Date of install:");
                        try
                        {
                            vmc.Add(Convert.ToString(ms.install_time));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("Startup time:");
                        try
                        {
                            vmc.Add(Convert.ToString(ms.start_time));
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }
                        vmc.Add("Time of last shutdown:");
                        try
                        {
                            vmc.Add(Convert.ToString(vm.other_config["last_shutdown_time"]));
                        }
                        catch
                        {
                            vmc.Add("Never Shutdown");
                        }
                        vmc.Add("Reason for Shutdown:");
                        try
                        {
                            vmc.Add(vm.other_config["last_shutdown_reason"]);
                        }
                        catch
                        {
                            vmc.Add("Never Shutdown");
                        }
                        vmc.Add("UUID");
                        try
                        {
                            vmc.Add(vm.uuid);
                        }
                        catch
                        {
                            vmc.Add("Data not available");
                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                            writelog.entry(log, entry);
                            vme++;
                        }



                        if (vm.guest_metrics.ServerOpaqueRef == "OpaqueRef:NULL")
                        {
                        }
                        else
                        {

                            VM_guest_metrics xms = VM_guest_metrics.get_record(session, vm.guest_metrics);

                            vmc.Add("OS version and System disk: ");
                            try
                            {
                                vmc.Add("Not displayed in demo version");
                            }
                            catch
                            {
                                vmc.Add("Data not available");
                                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                                writelog.entry(log, entry);
                                vme++;
                            }

                        }

                        
                        
                            VM vmRecord = VM.get_record(session, vmRef);
                            if (vmRecord.guest_metrics == "OpaqueRef:NULL")
                            {
                            }
                            else
                            {
                                XenRef<VM_guest_metrics> gmsref = vmRecord.guest_metrics;
                                VM_guest_metrics gms = VM_guest_metrics.get_record(session, gmsref);
                                Dictionary<String, String> dict = null;
                                dict = gms.networks;
                                String vmIP = null;
                                if (dict.Count >= 1)
                                {
                                    vmc.Add("Number of Xentools Visible VIFs:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(dict.Count));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                                        writelog.entry(log, entry);
                                        vme++;
                                    }

                                    foreach (String keyStr in dict.Keys)
                                    {
                                        vmc.Add("IP Address:");
                                        
                                        try
                                        {
                                            vmIP = (String)(dict[keyStr]);
                                            vmc.Add("Not displayed in demo version");
                                        }
                                        catch
                                        {
                                            vmc.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Error " + vme;
                                            writelog.entry(log, entry);
                                            vme++;
                                        }
                                    }
                                }

                                
                            }
                    }
                    vme = 1;
                }

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Finished ";
                writelog.entry(log, entry);
                
                
            }
            catch
            {

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " VM Collection Failed ";
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
