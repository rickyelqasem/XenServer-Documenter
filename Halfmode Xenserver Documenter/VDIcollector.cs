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
using System.IO;

namespace Halfmode_Xenserver_Documenter
{
    class VDIcollector
    {
        public object vdicollect(Session session)
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            StreamWriter SW;
            int vd = 1;
            try
            {

                List<XenRef<VM>> vmRefs = VM.get_all(session);
                List<XenRef<VBD>> vbdRefs = VBD.get_all(session);
                List<XenRef<VDI>> vdiRefs = VDI.get_all(session);



                foreach (XenRef<VM> vmRef in vmRefs)
                {

                    VM vm = VM.get_record(session, vmRef);
                    if (!vm.is_a_template && !vm.is_control_domain)
                    {


                        for (int i = 0; i <= vm.VBDs.Count - 1; )
                        {

                            if (vm.VBDs[i].ServerOpaqueRef == "OpaqueRef:NULL")
                            {
                            }
                            else
                            {
                                // get the info from VDB first.
                                VBD vbd = VBD.get_record(session, (String)vm.VBDs[i].ServerOpaqueRef);

                                if (vbd.VDI.ServerOpaqueRef == "OpaqueRef:NULL")
                                {
                                }
                                else
                                {
                                    vmc.Add("Virtual Machine Name:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vm.name_label));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }
                                    VDI vdi = VDI.get_record(session, vbd.VDI.ServerOpaqueRef);
                                    vmc.Add("Virtual Disk Name:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vdi.name_label));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }
                                    vmc.Add("Virtual Description:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vdi.name_description));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }
                                    vmc.Add("Virtual Disk Size:");
                                    try
                                    {
                                        vmc.Add("Not displayed in demo version");
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }
                                    vmc.Add("Physical Disk Usage:");
                                    try
                                    {
                                        vmc.Add("Not displayed in demo version");
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }

                                    SR sr = SR.get_record(session, vdi.SR.ServerOpaqueRef);
                                    vmc.Add("Parent Storage Repositry:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(sr.name_label));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }

                                    vmc.Add("Is VDI read-only:");
                                    try
                                    {
                                        vmc.Add(Convert.ToString(vdi.read_only));
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }
                                    vmc.Add("Is VDI Shareable:");
                                    try
                                    {
                                        vmc.Add("Not displayed in demo version");
                                    }
                                    catch
                                    {
                                        vmc.Add("Data not available");
                                        SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Error " + vd);
                                        SW.Close();
                                        vd++;
                                    }
                                }

                                i++;
                            }
                        }
                    }
                    vd = 1;
                }
                SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Finished");
                SW.Close();
                
            }
            catch
            {
                
                SW = File.AppendText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Collection Failed");
                SW.Close();
                
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
