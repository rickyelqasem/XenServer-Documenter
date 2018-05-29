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
    class PoolCollector
    {
        public Object poolcollect(Session session)
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            int pe = 1;
            string log;
            string entry;

            try
            {

                List<XenRef<Pool>> poolRefs = Pool.get_all(session);
                List<XenRef<Host>> hostRefs = Host.get_all(session);


                foreach (XenRef<Pool> poolRef in poolRefs)
                {
                    Pool pool = Pool.get_record(session, poolRef);
                    vmc.Add("Pool Name:");
                    try
                    {
                        vmc.Add(Convert.ToString(pool.name_label));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Error " + pe;
                        writelog.entry(log, entry);
                        pe++;
                    }
                    vmc.Add("Pool Description:");
                    try
                    {
                        vmc.Add(Convert.ToString(pool.name_description));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Error " + pe;
                        writelog.entry(log, entry);
                        pe++;

                    }
                    vmc.Add("Is HA enabled:");
                    try
                    {
                        vmc.Add(Convert.ToString(pool.ha_enabled));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Error " + pe;
                        writelog.entry(log, entry);
                        pe++;
                    }
                    vmc.Add("Number of host failures tolerated:");
                    try
                    {
                        vmc.Add(Convert.ToString(pool.ha_host_failures_to_tolerate));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Error " + pe;
                        writelog.entry(log, entry);
                        pe++;
                    }
                    vmc.Add("Is HA overcommited: ");
                    try
                    {
                        vmc.Add(Convert.ToString(pool.ha_overcommitted));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Error " + pe;
                        writelog.entry(log, entry);
                        pe++;
                    }

                    
                        
                    
                    vmc.Add("Pool Master: ");
                    try
                    {
                        Host host = Host.get_record(session, pool.master.ServerOpaqueRef);
                        vmc.Add(Convert.ToString(host.name_label));
                    }
                    catch
                    {
                        vmc.Add("Data not available");
                        log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                        entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Error " + pe;
                        writelog.entry(log, entry);
                        pe++;
                    }
                    pe = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Finished ";
                writelog.entry(log, entry);
            }
            catch
            {
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Pool Collection Failed";
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
