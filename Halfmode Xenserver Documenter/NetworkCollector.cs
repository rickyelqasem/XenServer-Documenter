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
    class NetworkCollector
    {
        public object netcollect(Session session)
        {
            int ne = 1;
            string log;
            string entry;
            
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            try
            {

                List<XenRef<Network>> netRefs = Network.get_all(session);
                List<XenRef<PIF>> pifRefs = PIF.get_all(session);
                List<XenRef<VIF>> vifRefs = VIF.get_all(session);
                List<XenRef<VM>> vmRefs = VM.get_all(session);
                List<XenRef<Host>> hostRefs = Host.get_all(session);

                



                foreach (XenRef<Network> netRef in netRefs)
                {


                    Network net = Network.get_record(session, netRef);


                    vmc.Add("Network Name:");
                    try
                    {
                        vmc.Add(Convert.ToString(net.name_label));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;

                    }
                    vmc.Add("Network Description:");
                    try
                    {
                        vmc.Add(Convert.ToString(net.name_description));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;

                    }
                    vmc.Add("Number of connected PIFs:");
                    try
                    {
                        vmc.Add(Convert.ToString(net.PIFs.Count));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;

                    }
                    if (net.PIFs.Count >= 1)
                    {


                        for (int pi = 0; pi <= net.PIFs.Count - 1; )
                        {

                            if (net.PIFs[pi].ServerOpaqueRef == "OpaqueRef:NULL")
                            {
                            }
                            else
                            {


                                PIF pif = PIF.get_record(session, net.PIFs[pi].ServerOpaqueRef);
                                Host host = Host.get_record(session, pif.host.ServerOpaqueRef);

                                vmc.Add("Host Member:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                pi++;
                            }
                        }
                    }
                    vmc.Add("Number of connected VIFs:");
                    try
                    {
                        vmc.Add(Convert.ToString(net.VIFs.Count));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    if (net.VIFs.Count >= 1)
                    {


                        for (int vi = 0; vi <= net.VIFs.Count - 1; )
                        {
                            if (net.VIFs[vi].ServerOpaqueRef == "OpaqueRef:NULL")
                            {
                            }
                            else
                            {

                                VIF vif = VIF.get_record(session, net.VIFs[vi].ServerOpaqueRef);
                                VM vm = VM.get_record(session, vif.VM.ServerOpaqueRef);

                                vmc.Add("VM Member:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;

                                }

                                vi++;
                            }
                        }
                    }
                    ne = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Finished ";
                writelog.entry(log, entry);
               

            }
            catch
            {

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Network Collection Failed ";
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
