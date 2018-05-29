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
    class ResidentVMs
    {

        public object rescollect(Session session, List<XenRef<VM>> vmRefs)
        {
            ArrayList hostres = new ArrayList();

            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            int re = 1;
            string log;
            string entry;
            try
            {
               

                List<XenRef<Host>> hostRefs = Host.get_all(session);
                

                foreach (XenRef<Host> hostRef in hostRefs)
                {
                    string resvm;

                    Host host = Host.get_record(session, hostRef);
                    hostres.Add("XenServer Name:");
                    try
                    {
                        hostres.Add(Convert.ToString(host.name_label));
                    }
                    catch
                    {
                        hostres.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " ResidentVM Collection Error " + re;
                        writelog.entry(log, entry);
                        re++;
                    }
                    hostres.Add("Number of Powered on VMS:");
                    try
                    {
                        hostres.Add(Convert.ToString(host.resident_VMs.Count - 1));
                    }
                    catch
                    {
                        hostres.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " ResidentVM Collection Error " + re;
                        writelog.entry(log, entry);
                        re++;
                    }
                    try
                    {
                        for (int i = 0; i <= host.resident_VMs.Count - 1; )
                        {
                            resvm = Convert.ToString(host.resident_VMs[i].ServerOpaqueRef);

                            foreach (XenRef<VM> vmRef in vmRefs)
                            {

                                VM vm = VM.get_record(session, vmRef);
                                if (Convert.ToString(vmRef) == resvm)
                                {
                                    if (!vm.is_a_template && !vm.is_control_domain)
                                    {
                                        hostres.Add("Running VM:");
                                        try
                                        {
                                            hostres.Add(Convert.ToString(vm.name_label));
                                        }
                                        catch
                                        {
                                            hostres.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " ResidentVM Collection Error " + re;
                                            writelog.entry(log, entry);
                                            re++;
                                        }
                                    }
                                }

                            }
                            i = i + 1;
                        }
                    }
                    catch
                    {
                    }

                    re = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " ResidentVM Collection Finished";
                writelog.entry(log, entry);
                
                
            }
            catch
            {
                
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " ResidentVM Collection Failed";
                        writelog.entry(log, entry);
                        
                
            }
            if ((hostres.Count & 2) == 0)
            {
            }
            else
            {
                hostres.Add(" ");
            }
            return hostres;
        }
    }
}