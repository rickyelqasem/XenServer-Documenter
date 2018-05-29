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
    class vifCollector
    {
        public Object vifcollect(Session session)
        {
            ArrayList vmc = new ArrayList();

            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            int vi = 1;
            string log;
                string entry;


                    try
                        {
                      

                        foreach (VIF vif in VIF.get_all_records(session).Values)
                        {
                            // resolve the host reference in the pif
                            VM vmvifs = VM.get_record(session, vif.VM);
                            if (!vmvifs.is_a_template && !vmvifs.is_control_domain)
                            {
                                vmc.Add("Virtual Machine Name:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vmvifs.name_label));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " VIF Collection Error " + vi;
                                    writelog.entry(log, entry);
                                    vi++;
                                }
                                vmc.Add("MAC Address:");
                                try
                                {
                                    vmc.Add("Not displayed in demo version");
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " VIF Collection Error " + vi;
                                    writelog.entry(log, entry);
                                    vi++;
                                }
                                vmc.Add("MTU:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vif.MTU));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " VIF Collection Error " + vi;
                                    writelog.entry(log, entry);
                                    vi++;
                                }
                                vmc.Add("Is VIF currently attached:");
                                try
                                {
                                    vmc.Add(Convert.ToString(vif.currently_attached));
                                }
                                catch
                                {
                                    vmc.Add("Data not available");
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " VIF Collection Error " + vi;
                                    writelog.entry(log, entry);
                                    vi++;
                                }
                            }
                            vi = 1;
                        }


                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " VIF Collection Finished ";
                        writelog.entry(log, entry);
                        
                        
                       
                        }
                    catch
                    {

                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " VIF Collection Failed ";
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
