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

    class Hostcollector
    {
        public object hostcollect(Session session)
        {
        ArrayList vmc = new ArrayList();
        int ne = 1;
        string log;
        string entry;
        string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        

            try
            {
                //Session session = new Session(server, Convert.ToInt32(port));
                //session.login_with_password(username, password);
                
                
                    List<XenRef<Host>> hostRefs = Host.get_all(session);
                    List<XenRef<Host_cpu>> hcpuRefs = Host_cpu.get_all(session);
                
               
                foreach (XenRef<Host> hostRef in hostRefs)
                {
                    
                        Host host = Host.get_record(session, hostRef);
                        XenRef<Host_metrics> gmsref = Host.get_metrics(session, host.opaque_ref);
                        Host_metrics gms = Host_metrics.get_record(session, gmsref);
                   
                    vmc.Add("Xenserver Name:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.name_label));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }

                    vmc.Add("Hostname:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.hostname));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }

                    vmc.Add("Is Host alive:");
                    try
                    {
                        vmc.Add(Convert.ToString(gms.live));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }

                    vmc.Add("IP address:");
                    try
                    {
                        vmc.Add("Not displayed in demo version");
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }

                    vmc.Add("Total Host Memory:");
                    try
                    {
                        vmc.Add("Not displayed in demo version");
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("Host Memory Free:");
                    try
                    {
                        vmc.Add("Not displayed in demo version");
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("Number CPU cores:");
                    try
                    {
                        vmc.Add("Not displayed in demo version");
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    try
                    {
                    foreach (XenRef<Host_cpu> hcpuRef in hcpuRefs)
                    {
                        
                            Host_cpu hcpu = Host_cpu.get_record(session, hcpuRef);


                            if (hcpu.host.ServerOpaqueRef == hostRef)
                            {
                                vmc.Add("CPU Core Number:");
                                try
                                {
                                    vmc.Add(Convert.ToString(hcpu.number));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++; 
                                }
                                vmc.Add("CPU Core Vendor:");
                                try
                                {
                                    vmc.Add(Convert.ToString(hcpu.vendor));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("CPU Core Model Name:");
                                try
                                {
                                    vmc.Add(Convert.ToString(hcpu.modelname));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("CPU Core Speed:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("CPU Core Stepping Revesion:");
                                try
                                {
                                    vmc.Add(hcpu.stepping);
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("CPU Core Family:");
                                try
                                {
                                    vmc.Add(Convert.ToString(hcpu.family));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                            }
                     
                        }
                    
                    }
                    catch
                    {
                    }

                    vmc.Add("ISCSI IQN Name:");
                    try
                    {
                        vmc.Add("Not displayed in demo version");
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("Number of Allowed Operations:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.allowed_operations.Count));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    try
                    {
                        for (int i = 0; i <= host.allowed_operations.Count - 1; )
                        {
                            vmc.Add("Allowed Operation:");
                            try
                            {
                                vmc.Add(Convert.ToString(host.allowed_operations[i]));
                            }
                            catch
                            {
                                vmc.Add("Data not available");
                                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                writelog.entry(log, entry);
                                ne++;
                            }
                            i = i + 1;
                        }
                    }
                    catch
                    {
                    }

                    vmc.Add("Xenserver version:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.software_version["product_version"]));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("Build Number:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.software_version["build_number"]));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("Kernel version:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.software_version["linux"]));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("API version:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.API_version_vendor));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    vmc.Add("Number of Physical Block Devices (PBD):");
                    try
                    {
                        vmc.Add(Convert.ToString(host.PBDs.Count));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    try
                    {
                        for (int i = 0; i <= host.PBDs.Count - 1; )
                        {

                            if (host.PBDs[i].ServerOpaqueRef == "OpaqueRef:NULL")
                            {
                            }
                            else
                            {


                                PBD pbd = PBD.get_record(session, host.PBDs[i].ServerOpaqueRef);

                                Dictionary<String, String> dict = null;
                                dict = pbd.device_config;
                                String pdninfo = null;
                                foreach (String keyStr in dict.Keys)
                                {
                                    if (keyStr == "location" || keyStr == "device")
                                    {
                                        pdninfo = (String)(dict[keyStr]);
                                        vmc.Add("Physical Block Devices:");
                                        try
                                        {
                                            vmc.Add(Convert.ToString(pdninfo));
                                        }
                                        catch
                                        {
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                            writelog.entry(log, entry);
                                            ne++;
                                        }
                                    }
                                }
                                i = i + 1;
                            }
                        }
                    }
                    catch
                    {
                    }
                    vmc.Add("Number of Physical NICs (PIF):");
                    try
                    {
                        vmc.Add(Convert.ToString(host.PIFs.Count));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                        writelog.entry(log, entry);
                        ne++;
                    }
                    try
                    {
                        for (int i = 0; i <= host.PIFs.Count - 1; )
                        {
                            if (host.PIFs[i].ServerOpaqueRef == "OpaqueRef:NULL")
                            {
                            }
                            else
                            {
                                PIF vbd = PIF.get_record(session, Convert.ToString(host.PIFs[i].ServerOpaqueRef));
                                PIF_metrics pif = PIF_metrics.get_record(session, vbd.metrics);


                                vmc.Add("PIF IP Address:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Netmask:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.netmask));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF DHCP/Static:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.ip_configuration_mode));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF MAC Address:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF DNS servers:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Gateway:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.gateway));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                if (vbd.other_config.Count >= 1)
                                {
                                    vmc.Add("PIF Domain:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vbd.other_config["domain"]));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                        writelog.entry(log, entry);
                                        ne++;
                                    }
                                }
                                vmc.Add("PIF VLAN ID:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.VLAN));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF MTU:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.MTU));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Device:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.device));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("Is PIF attached:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vbd.currently_attached));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }

                                vmc.Add("PIF Device Name:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.device_name));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Device ID:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.device_id));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Vendor Name:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.vendor_name));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Vendor ID:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.vendor_id));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Speed:");
                                if (Convert.ToString(pif.speed) == "65535")
                                {
                                    vmc.Add("PIF not connected");
                                }
                                else
                                {
                                    try
                                    {
                                        vmc.Add(Convert.ToString(pif.speed));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                        entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                        writelog.entry(log, entry);
                                        ne++;
                                    }

                                }
                                vmc.Add("PIF is set to duplex:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.duplex));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF PCI Bus:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.pci_bus_path));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++;
                                }
                                vmc.Add("PIF Carrier:");
                                try
                                {
                                    vmc.Add(Convert.ToString(pif.carrier));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Error " + ne;
                                    writelog.entry(log, entry);
                                    ne++; ;
                                }
                                i = i + 1;



                            }
                        }
                    }
                        catch
                    {
                        }
                    ne = 1;
                    
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Finished ";
                writelog.entry(log, entry);
                
            }
            catch
            {


                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Host Collection Failed";
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

