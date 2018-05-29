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
    class BondCollector
    {
        public object bondcollect(Session session)
        {
            ArrayList vmc = new ArrayList();
            int be = 1;
            string log;
            string entry;
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            

            try
            {

            List<XenRef<Bond>> bondRefs = Bond.get_all(session);
            List<XenRef<PIF>> pifRefs = PIF.get_all(session);
            List<XenRef<Host>> hostRefs = Host.get_all(session);
            int bondnumber = 1;
            string hostname = "";
            string pifmac;
            
            
                foreach (XenRef<Bond> bondRef in bondRefs)
                {
                    Bond bond = Bond.get_record(session, bondRef);
                    vmc.Add("Bond Number:");


                    vmc.Add(Convert.ToString(bondnumber));
                    pifmac = "";
                    for (int i = 0; i <= bond.slaves.Count - 1; )
                    {

                        foreach (XenRef<PIF> pifRef in pifRefs)
                        {
                            PIF pif = PIF.get_record(session, pifRef);
                            if (bond.slaves[i].ServerOpaqueRef == pif.opaque_ref)
                            {
                                Host host = Host.get_record(session, pif.host.ServerOpaqueRef);

                                try
                                {
                                    hostname = Convert.ToString(host.name_label);
                                }
                                catch
                                {
                                    hostname = "Data not available";
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Bond Collection Error " + be;
                                    writelog.entry(log, entry);
                                    be++;
                                }
                                try
                                {
                                    pifmac = pifmac + " | " + Convert.ToString(pif.MAC);
                                }
                                catch
                                {
                                    pifmac = "Data not available";
                                    log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                                    entry = DateTime.Now.ToString("HH:mm:ss") + " Bond Collection Error " + be;
                                    writelog.entry(log, entry);
                                    be++;
                                }

                            }
                        }
                        i++;


                    }
                    bondnumber++;
                    vmc.Add("Bond Parent:");
                    vmc.Add(hostname);
                    vmc.Add("Bond members by MAC:");
                    vmc.Add(pifmac);
                    be = 1;
                }
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Bond Collection Finished";
                writelog.entry(log, entry);
                
            }
            catch
            {
                log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
                entry = DateTime.Now.ToString("HH:mm:ss") + " Bond Collection Failed";
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


