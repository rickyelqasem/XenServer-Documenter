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
    class SRCollector
    {
        public object srcollect(Session session)
        {
            ArrayList vmc = new ArrayList();

            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string log;
            string entry;
            int sre = 1;


            try
            {
                List<XenRef<Host>> hostRefs = Host.get_all(session);
                foreach (XenRef<Host> hostRef in hostRefs)
                {
                    Host host = Host.get_record(session, hostRef);
                    vmc.Add("Xenserver Name:");
                    try
                    {
                        vmc.Add(Convert.ToString(host.name_label));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Error " + sre;
                        writelog.entry(log, entry);
                        sre++;
                    }

                    for (int i = 0; i <= host.PBDs.Count - 1; )
                    {
                        if (host.PBDs[i].ServerOpaqueRef == "OpaqueRef:NULL")
                        {
                        }
                        else
                        {
                            PBD pbd = PBD.get_record(session, host.PBDs[i].ServerOpaqueRef);


                            List<XenRef<SR>> srRefs = SR.get_all(session);
                            foreach (XenRef<SR> srRef in srRefs)
                            {
                                SR sr = SR.get_record(session, srRef);
                                for (int pi = 0; pi <= sr.PBDs.Count - 1; )
                                {
                                    if (pbd.opaque_ref == sr.PBDs[pi].ServerOpaqueRef)
                                    {
                                        vmc.Add("Storage Repository Name:");
                                        try
                                        {
                                            vmc.Add(Convert.ToString(sr.name_label));
                                        }
                                        catch
                                        {
                                            vmc.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Error " + sre;
                                            writelog.entry(log, entry);
                                            sre++;
                                        }
                                        vmc.Add("Storage Repository Description:");
                                        try
                                        {
                                            vmc.Add(Convert.ToString(sr.name_description));
                                        }
                                        catch
                                        {
                                            vmc.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Error " + sre;
                                            writelog.entry(log, entry);
                                            sre++;
                                        }
                                        vmc.Add("Storage Repository Usage:");
                                        try
                                        {
                                            double srutil = sr.physical_utilisation / 1e9;
                                            vmc.Add("Not displayed in demo version");
                                        }
                                        catch
                                        {
                                            vmc.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Error " + sre;
                                            writelog.entry(log, entry);
                                            sre++;
                                        }
                                        vmc.Add("Storage Repository Size:");
                                        try
                                        {
                                            double srsize = sr.physical_size / 1e9;
                                            vmc.Add("Not displayed in demo version");
                                        }
                                        catch
                                        {
                                            vmc.Add("Data not available");
                                            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                            entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Error " + sre;
                                            writelog.entry(log, entry);
                                            sre++;
                                        }
                                    }
                                    pi++;
                                }



                            }
                            i++;
                        }
                    }
                    sre = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Finished";
                writelog.entry(log, entry);
                
                
            }
            catch
            {

                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " SR Collection Failed";
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
