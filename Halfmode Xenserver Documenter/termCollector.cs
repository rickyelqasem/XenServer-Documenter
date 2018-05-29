using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace Halfmode_Xenserver_Documenter
{
    class termCollector
    {
        public object termcollect()
        {
            ArrayList vmc = new ArrayList();
            string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string log;
            string entry;

            vmc.Add("API version");
            vmc.Add("Refers to the vendor of the API (Application Program Interface) used. With Citrix Xenserver 5.5 the vendor is Xensource");
            vmc.Add("Bond");
            vmc.Add("Improves host resiliency by using two physical NICs as one. If one NIC in a bond fails, then the network traffic automatically routes to the second NIC. Works in an active/passive mode,  with one physical NIC in use at a time. Is different than link aggregation, which acts in active/active mode. XenServer NIC bonds subsume the underlying physical interfaces");
            vmc.Add("CPU Core");
            vmc.Add("A single physical CPU may contain multiple CPU cores which are seen by the operating system as seperate CPUs.");
            vmc.Add("Host");
            vmc.Add("The physical hardware computer platform (server) in which Citrix Xenserver is installed.");
            vmc.Add("Hostname");
            vmc.Add("This is the Host's Fully Qualified Domain Name (FQDN) and often the same as the Xenserver name.");
            vmc.Add("Kernel version");
            vmc.Add("Xenserver software is built upon a linux operating system and this parameter is the version details of the controlling Linux platform.");
            vmc.Add("Network");
            vmc.Add("Represents a virtual Ethernet switch on a XenServer host. Network objects have a name, a UUID and the collection of VIFs and PIFs connected to the network");
            vmc.Add("PBD");
            vmc.Add("Represents the interface between a physical host and an attached storage repository. PBDs are connector objects that allow a given storage repository to be mapped to a host. PBDs store the device configuration fields that are used to connect to and interact with a given storage target. In the case of NFS, for instance, this device configuration includes the IP address of the NFS server and the associated mount path. PBD objects manage the runtime attachment of a given storage repository to a given host.");
            vmc.Add("PIF");
            vmc.Add("Represents a physical network interface on a XenServer host.");
            vmc.Add("Pool");
            vmc.Add("Resource pools allow multiple virtualization servers to be treated as a single entity from a management perspective, creating consolidated management of server resources. In addition Resource Pools allow all servers to share a common framework for network and storage, which facilitates features such as Automatic Placement of VMs and XenMotion.");
            vmc.Add("Snapshot");
            vmc.Add("A snapshot is a locally retained point-in-time image of data. These images are frequently used as user-recoverable backups.");
            vmc.Add("SR");
            vmc.Add("A storage repository (SR) allows pooling of disk devices and stores one or more VM virtual disk drives.");
            vmc.Add("Template");
            vmc.Add("A virtual machine (VM) template is a XenServer file that includes: • The VM environment • The settings for optimum storage, CPU and memory • The network configuration information.");
            vmc.Add("VBD");
            vmc.Add("A virtual block device (VBD) is a connector object that is similar to PBD and that allows mappings between VDIs and virtual machines. In addition to providing a mechanism to attach a VDI to a virtual machine, VBDs allow the fine-tuning of parameters regarding QoS, statistics and the bootability of a given VDI");
            vmc.Add("VCPU");
            vmc.Add("Defines a shared scheduled portion of a physical CPU to a virtual machine. This partitioned use of the CPU is transparent to the virtualized guest operating system.");
            vmc.Add("VDI");
            vmc.Add("An on-disk representation of a virtual disk provided to a guest VM. The VDI is the fundamental unit of virtualized storage in XenServer. The VDI is a virtual disk drive.");
            vmc.Add("VIF");
            vmc.Add("Represents a virtual interface on a virtual machine.");
            vmc.Add("VLAN");
            vmc.Add("VLAN's Separate segments of the network logically instead of physically. VLAN's Control broadcast traffic more efficiently and  subnetworks can be configured without routers.");
            vmc.Add("VM");
            vmc.Add("Virtua Machines (VM) refers to the guest operating system running in a virtualized environment  opposed to running on physical hardware.");
            vmc.Add("Xenserver Name");
            vmc.Add("This is the name label of the Host system used for Citrix Xenserver.");

            if ((vmc.Count & 2) == 0)
            {
            }
            else
            {
                vmc.Add(" ");
            }
            log = mydocs + "\\Halfmode\\HalfmodeConnection.log";
            entry = DateTime.Now.ToString("HH:mm:ss") + " Term Collection Finished ";
            writelog.entry(log, entry);
            return vmc;
        }
    }
}
