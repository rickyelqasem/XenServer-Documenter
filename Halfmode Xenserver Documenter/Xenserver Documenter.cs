using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using GraphicsHandler;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;
using System.Drawing.Text;
using XenAPI;
using System.Collections;
using Halfmode_Xenserver_Documenter;
using System.Threading;
using System.Xml;
using System.IO;

namespace Halfmode_Xenserver_Documenter
{
    
    public partial class Xenserver_Documenter : Form
    {
        int pagebordercolor;
        int titlebordercolor;
        string template;
        string stremail;
        string strhost;
        string stropt;
        int objcolbord = 8421440;
        int objcoltit = 8421440;
        bool hasconnected = false;
        bool logentry = false;
        int collectcont = 0;
        ArrayList vmlist = new ArrayList();
        ArrayList viflist = new ArrayList();
        ArrayList hostlist = new ArrayList();
        ArrayList piflist = new ArrayList();
        ArrayList srlist = new ArrayList();
        ArrayList reslist = new ArrayList();
        ArrayList snaplist = new ArrayList();
        ArrayList vbdlist = new ArrayList();
        ArrayList templatelist = new ArrayList();
        ArrayList poollist = new ArrayList();
        ArrayList bondlist = new ArrayList();
        ArrayList netlist = new ArrayList();
        ArrayList vdilist = new ArrayList();
        ArrayList termlist = new ArrayList();
        string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        

        
        private Thread trd;
        private Thread wtrd;
        private Thread trdvif;
        private Thread trdhost;
        private Thread trdnet;
        private Thread trdpool;
        private Thread trdresvm;
        private Thread trdsnap;
        private Thread trdsr;
        private Thread trdtemp;
        private Thread trdterm;
        private Thread trdvbd;
        //private Thread trdvm;
        private Thread trdbond;
        private Thread trdvdi;
        StreamWriter SW;

        public Xenserver_Documenter()
        {
            InitializeComponent();
            
            string newPath = System.IO.Path.Combine(mydocs, "Halfmode");
            System.IO.Directory.CreateDirectory(newPath);
            string mydochalfmode = mydocs + "\\Halfmode";
            Thread trd = new Thread(new ThreadStart(this.ThreadTask));
            //checklicense();
            connectionTab.BackColor = Color.SteelBlue;
            wordLayoutTab.BackColor = Color.SteelBlue;
            about.BackColor = Color.SteelBlue;
            //excelTab.BackColor = Color.SteelBlue;
            //pdfTab.BackColor = Color.SteelBlue;
            hostBox.Text = System.Net.Dns.GetHostName();
            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button3.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            btnBorderColor.BackColor = Color.SteelBlue;
            btnTitleBorder.BackColor = Color.SteelBlue;
            btnOpenTemplate.BackColor = Color.SteelBlue;
            btnWordGenerate.BackColor = Color.SteelBlue;
            textPageBorderCol.BackColor = Color.Teal;
            textTitleBorder.BackColor = Color.Teal;
            btnWordGenerate.Enabled = false;

          

            try
            {

                //get saved word settings
                XmlDocument xmlread = new XmlDocument();
                xmlread.Load(mydocs + "\\Halfmode\\HalfmodeWord.xml");
                XmlNode namexml = xmlread.SelectSingleNode(@"Word/Settings/Name");
                textBox14.Text = namexml.InnerText;
                XmlNode namexml2 = xmlread.SelectSingleNode(@"Word/Settings/Company");
                textBox15.Text = namexml2.InnerText;
                XmlNode namexml3 = xmlread.SelectSingleNode(@"Word/Settings/Address1");
                textBox16.Text = namexml3.InnerText;
                XmlNode namexml4 = xmlread.SelectSingleNode(@"Word/Settings/Address2");
                textBox17.Text = namexml4.InnerText;
                XmlNode namexml5 = xmlread.SelectSingleNode(@"Word/Settings/Town");
                textBox18.Text = namexml5.InnerText;
                XmlNode namexml6 = xmlread.SelectSingleNode(@"Word/Settings/County");
                textBox19.Text = namexml6.InnerText;
                XmlNode namexml7 = xmlread.SelectSingleNode(@"Word/Settings/Country");
                textBox20.Text = namexml7.InnerText;
                XmlNode namexml8 = xmlread.SelectSingleNode(@"Word/Settings/Post");
                textBox21.Text = namexml8.InnerText;
                XmlNode namexml9 = xmlread.SelectSingleNode(@"Word/Settings/Header");
                textBox11.Text = namexml9.InnerText;
                XmlNode namexml10 = xmlread.SelectSingleNode(@"Word/Settings/Title");
                textBox12.Text = namexml10.InnerText;
                XmlNode namexml11 = xmlread.SelectSingleNode(@"Word/Settings/Subtitle");
                textBox13.Text = namexml11.InnerText;
                xmlread.Save(mydocs + "\\Halfmode\\HalfmodeWord.xml");
                
                //get saved connection settings
                XmlDocument xmlread2 = new XmlDocument();
                xmlread2.Load(mydocs + "\\Halfmode\\HalfmodeConnect.xml");
                XmlNode namexml12 = xmlread2.SelectSingleNode(@"Connect/Settings/Server");
                textBox2.Text = namexml12.InnerText;
                XmlNode namexml13 = xmlread2.SelectSingleNode(@"Connect/Settings/User");
                textBox5.Text = namexml13.InnerText;
                xmlread2.Save(mydocs + "\\Halfmode\\HalfmodeConnect.xml");


            }
            catch
            {

            }


           


        }

        private void ThreadTask()
        {
            
            CollectBusy collectbusy = new CollectBusy();
            collectbusy.ShowDialog();
        }
        private void WordTread()
        {
            Wordbusy wordbusy = new Wordbusy();
            wordbusy.ShowDialog();
        }
        private void Threadvif()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);

            vifCollector vifc = new vifCollector();
            viflist = (ArrayList)vifc.vifcollect(session);
            collectcont++;
           
            
            

        }
        private void Threadnet()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);

            NetworkCollector netc = new NetworkCollector();
            netlist = (ArrayList)netc.netcollect(session);
            collectcont++;

        }
        private void Threadhost()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);

            Hostcollector hostc = new Hostcollector();
            hostlist = (ArrayList)hostc.hostcollect(session);
            collectcont++;

        }
        private void Threadsr()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);

            SRCollector SRc = new SRCollector();
            srlist = (ArrayList)SRc.srcollect(session);
            collectcont++;


        }
        private void Threadresvm()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);
            List<XenRef<VM>> vmRefs = VM.get_all(session);
            ResidentVMs RESc = new ResidentVMs();
            reslist = (ArrayList)RESc.rescollect(session, vmRefs);
            collectcont++;

        }
        private void Threadsnap()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);
            List<XenRef<VM>> vmRefs = VM.get_all(session);
            snapcollector snapc = new snapcollector();
            snaplist = (ArrayList)snapc.snapcollect(session, vmRefs);
            collectcont++;

        }

        private void Threadvbd()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);
            List<XenRef<VM>> vmRefs = VM.get_all(session);
            vbdCollector vbdc = new vbdCollector();
            vbdlist = (ArrayList)vbdc.vbdcollect(session, vmRefs);
            collectcont++;

        }
        private void Threadtemp()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);
            List<XenRef<VM>> vmRefs = VM.get_all(session);
            templateCollector tempc = new templateCollector();
            templatelist = (ArrayList)tempc.templatecollect(session, vmRefs);
            collectcont++;

        }
        private void Threadpool()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);

            PoolCollector poolc = new PoolCollector();
            poollist = (ArrayList)poolc.poolcollect(session);
            collectcont++;

        }
         private void Threadbond()
        {
            string server = textBox2.Text;
            string port = textBox4.Text;
            string username = textBox5.Text;
            string password = textBox6.Text;

            Session session = new Session(server, Convert.ToInt32(port));
            session.login_with_password(username, password);

            BondCollector bondc = new BondCollector();
            bondlist = (ArrayList)bondc.bondcollect(session);
            collectcont++;

        }
         private void Threadvdi()
         {
             string server = textBox2.Text;
             string port = textBox4.Text;
             string username = textBox5.Text;
             string password = textBox6.Text;

             Session session = new Session(server, Convert.ToInt32(port));
             session.login_with_password(username, password);

             VDIcollector vdic = new VDIcollector();
             vdilist = (ArrayList)vdic.vdicollect(session);
             collectcont++;

         }
         private void Threadterm()
         {
             string server = textBox2.Text;
             string port = textBox4.Text;
             string username = textBox5.Text;
             string password = textBox6.Text;

             Session session = new Session(server, Convert.ToInt32(port));
             session.login_with_password(username, password);

             termCollector termc = new termCollector();
             termlist = (ArrayList)termc.termcollect();
             collectcont++;

         }


        private void btnBorderColor_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            int R = colorDialog1.Color.R;
            int G = colorDialog1.Color.G;
            int B = colorDialog1.Color.B;
            pagebordercolor = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            textPageBorderCol.BackColor = Color.FromArgb(R, G, B);
            int cl = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            objcolbord = cl;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            int R = colorDialog1.Color.R;
            int G = colorDialog1.Color.G;
            int B = colorDialog1.Color.B;
            titlebordercolor = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            textTitleBorder.BackColor = Color.FromArgb(R, G, B);
            int cl = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            objcoltit = cl;
        } //titlie border

          

        private void btnOpenTemplate_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog tempfile = new OpenFileDialog();
            tempfile.Filter = "Word Templates (*.dot)|*.dot|All files (*.*)|*.*";
            tempfile.Title = "Open a Word Template";
            tempfile.ShowDialog();
            label11.Text = (System.IO.Path.GetFileNameWithoutExtension(tempfile.FileName));
            template = tempfile.FileName;
            if (template != "")
            {
                checkPageBorder1.Checked = false;
                pageNumber.Checked = false;
                textBox11.Text = "";
            }
            else
            {
                label11.Text = "N/A";
            }
        }

        private void button1_Click(object sender, EventArgs e)//register key
        {
            //create new config xml doc
            XmlTextWriter textwriter = new XmlTextWriter(mydocs + "\\Halfmode\\HalfmodeConfig.xml", null);
            textwriter.WriteStartDocument();
            textwriter.WriteStartElement("Config");
            textwriter.WriteEndElement();
            textwriter.WriteEndDocument();
            textwriter.Close();

            string hostname = hostBox.Text;
            string email = textBox8.Text;
            string hostemail = email + hostname;
            string cryptemail = StringCrypto.EncryptString(email);
            string cryptemailhost = StringCrypto.EncryptString(hostemail);
            bool hostok = false;
            
            if (textBox9.Text == cryptemailhost)
            {
               // masterkey.SetValue("email", StringCrypto.EncryptString(textBox8.Text));
               // masterkey.SetValue("host", textBox9.Text);
               // masterkey.SetValue("options", textBox10.Text);
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeConfig.xml");
                XmlElement el = xmldoc.CreateElement("License");
                string License = "<Email>" + cryptemail + "</Email>" +
                    "<Host>" + textBox9.Text + "</Host>" + "<Options>" + textBox10.Text + "</Options>";
                el.InnerXml = License;
                xmldoc.DocumentElement.AppendChild(el);
                xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeConfig.xml");

                hostok = true;
            }
            else
            {
                MessageBox.Show("Invalid Host license");
            }

            if (textBox10.Text == "XxjBF1SzAOQA")
            {
                //open word tab
                //tabControl1.TabPages.Add(wordLayoutTab);
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeConfig.xml");
                XmlElement el = xmldoc.CreateElement("License");
                string License = "<Email>" + cryptemail + "</Email>" +
                    "<Host>" + textBox9.Text + "</Host>" + "<Options>" + textBox10.Text + "</Options>";
                el.InnerXml = License;
                xmldoc.DocumentElement.AppendChild(el);
                xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeConfig.xml");

                hostok = true;
                if (hostok && hasconnected)
                {
                    button1.Enabled = false;
                    button2.Enabled = true;
                    textBox8.Visible = false;
                    textBox9.Visible = false;
                    textBox10.Visible = false;
                    hostBox.Visible = false;
                }
            }
            else if (textBox10.Text == "SOJYiDdLynKkqkxRovLIPgA=")
            {
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeConfig.xml");
                XmlElement el = xmldoc.CreateElement("License");
                string License = "<Email>" + cryptemail + "</Email>" +
                    "<Host>" + textBox9.Text + "</Host>" + "<Options>" + textBox10.Text + "</Options>";
                el.InnerXml = License;
                xmldoc.DocumentElement.AppendChild(el);
                xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeConfig.xml");

                hostok = true;
                if (hostok && hasconnected)
                {
                    button1.Enabled = false;
                    button2.Enabled = true;
                    textBox8.Visible = false;
                    textBox9.Visible = false;
                    textBox10.Visible = false;
                    hostBox.Visible = false;
                }
            }
            else if (textBox10.Text == "J5kd+LNBCpsA")
            {
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeConfig.xml");
                XmlElement el = xmldoc.CreateElement("License");
                string License = "<Email>" + cryptemail + "</Email>" +
                    "<Host>" + textBox9.Text + "</Host>" + "<Options>" + textBox10.Text + "</Options>";
                el.InnerXml = License;
                xmldoc.DocumentElement.AppendChild(el);
                xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeConfig.xml");

                hostok = true;
                if (hostok && hasconnected)
                {
                    button1.Enabled = false;
                    button2.Enabled = true;
                    textBox8.Visible = false;
                    textBox9.Visible = false;
                    textBox10.Visible = false;
                    hostBox.Visible = false;
                }
            }
            else if (textBox10.Text == "SOJYiDdLynI1Kh37TIAh+gA=")
            {
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeConfig.xml");
                XmlElement el = xmldoc.CreateElement("License");
                string License = "<Email>" + cryptemail + "</Email>" +
                    "<Host>" + textBox9.Text + "</Host>" + "<Options>" + textBox10.Text + "</Options>";
                el.InnerXml = License;
                xmldoc.DocumentElement.AppendChild(el);
                xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeConfig.xml");

                hostok = true;
                if (hostok && hasconnected)
                {
                    button1.Enabled = false;
                    button2.Enabled = true;
                    textBox8.Visible = false;
                    textBox9.Visible = false;
                    textBox10.Visible = false;
                    hostBox.Visible = false;
                }
            }
            else
            {
                MessageBox.Show("Invalid Options license");
            }
            //checklicense();

        }

        public void checklicense()
        {
            bool hostok = false;
            
            try
            {
                
                XmlDocument xmlread3 = new XmlDocument();
                xmlread3.Load(mydocs + "\\Halfmode\\HalfmodeConfig.xml");
                XmlNode conxml = xmlread3.SelectSingleNode(@"Config/License/Email");
                stremail = conxml.InnerText;
                XmlNode conxml2 = xmlread3.SelectSingleNode(@"Config/License/Host");
                strhost = conxml2.InnerText;
                XmlNode conxml3 = xmlread3.SelectSingleNode(@"Config/License/Options");
                stropt = conxml3.InnerText;

                xmlread3.Save(mydocs + "\\Halfmode\\HalfmodeConfig.xml");

                textBox8.Text = StringCrypto.DecryptString(stremail);
                textBox9.Text = strhost;
                textBox10.Text = stropt;
            }
            catch
            {
                MessageBox.Show("Please review License");
                return;
            }
            string namehost = System.Net.Dns.GetHostName();
            string deemail = StringCrypto.DecryptString(stremail);
            deemail = deemail + namehost;
            string crypthostname = StringCrypto.EncryptString(deemail);
            if (strhost == crypthostname)
            {
                hostok = true;
                // hostname is correct then enable generate buttons;
                //btnWordGenerate.Enabled = true;
                //btnWordGenerate.Text = "Generate Word Document";
            }
            if (stropt == "XxjBF1SzAOQA")
            {
                
                button1.Enabled = false;
                button2.Enabled = true;
                textBox8.Enabled = false;
                textBox9.Enabled = false;
                textBox10.Enabled = false;
                textBox8.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                hostBox.Visible = false;
                label5.Text = "Registered";
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;

                if (hostok && hasconnected)
                {
                    btnWordGenerate.Enabled = true;
                    btnWordGenerate.Text = "Generate Word Document";
                }
                return;
            }
            if (stropt == "SOJYiDdLynKkqkxRovLIPgA=")
            {
                //open word + excel tab
                //tabControl1.TabPages.Add(wordLayoutTab);
                //tabControl1.TabPages.Add(excelTab);
                button1.Enabled = false;
                button2.Enabled = true;
                textBox8.Enabled = false;
                textBox9.Enabled = false;
                textBox10.Enabled = false;
                textBox8.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                hostBox.Visible = false;
                label5.Text = "Registered";
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;

                if (hostok && hasconnected)
                {
                    btnWordGenerate.Enabled = true;
                    btnWordGenerate.Text = "Generate Word Document";
                }
                return;
            }
            if (stropt == "J5kd+LNBCpsA")
            {
                //open word + pdf tab
                //tabControl1.TabPages.Add(wordLayoutTab);
                //tabControl1.TabPages.Add(pdfTab);
                button1.Enabled = false;
                button2.Enabled = true;
                textBox8.Enabled = false;
                textBox9.Enabled = false;
                textBox10.Enabled = false;
                label5.Text = "Registered";
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                if (hostok && hasconnected)
                {
                    btnWordGenerate.Enabled = true;
                    btnWordGenerate.Text = "Generate Word Document";
                }
                return;
            }
            if (stropt == "SOJYiDdLynI1Kh37TIAh+gA=")
            {
                //open word + excel + pdf tab
                //tabControl1.TabPages.Add(wordLayoutTab);
                //tabControl1.TabPages.Add(excelTab);
                //tabControl1.TabPages.Add(pdfTab);
                button1.Enabled = false;
                button2.Enabled = true;
                textBox8.Enabled = false;
                textBox9.Enabled = false;
                textBox10.Enabled = false;
                textBox8.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                hostBox.Visible = false;
                label5.Text = "Registered";
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;

                if (hostok && hasconnected)
                {
                    btnWordGenerate.Enabled = true;
                    btnWordGenerate.Text = "Generate Word Document";
                }
                return;
            }
            
            
            return; 
}

        private void button2_Click(object sender, EventArgs e)//new reg key
        {
            button1.Enabled = true;
            button2.Enabled = false;
            textBox8.Visible = true;
            textBox9.Visible = true;
            textBox10.Visible = true;
            hostBox.Visible = true;
            textBox8.Enabled= true;
            textBox9.Enabled = true;
            textBox10.Enabled= true;
            label5.Text = "Email Address";
            label6.Visible = true;
            label7.Visible = true;
            label8.Visible = true;
        }

        private void btnWordGenerate_Click(object sender, EventArgs e)
        {

            this.Hide();
            //create wordsettings xml

            XmlTextWriter textwriter = new XmlTextWriter(mydocs + "\\Halfmode\\HalfmodeWord.xml", null);
            textwriter.WriteStartDocument();
            textwriter.WriteStartElement("Word");
            textwriter.WriteEndElement();
            textwriter.WriteEndDocument();
            textwriter.Close();



            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeWord.xml");
            XmlElement el = xmldoc.CreateElement("Settings");
            string Settings = "<Name>" + textBox14.Text + "</Name>" +
                "<Company>" + textBox15.Text + "</Company>" + "<Address1>" + textBox16.Text + "</Address1>"
                + "<Address2>" + textBox17.Text + "</Address2>"
                + "<Town>" + textBox18.Text + "</Town>"
                + "<County>" + textBox19.Text + "</County>"
                + "<Country>" + textBox20.Text + "</Country>"
                + "<Post>" + textBox21.Text + "</Post>"
                + "<Header>" + textBox11.Text + "</Header>"
                + "<Title>" + textBox12.Text + "</Title>"
                + "<Subtitle>" + textBox13.Text + "</Subtitle>";
            el.InnerXml = Settings;
            xmldoc.DocumentElement.AppendChild(el);
            xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeWord.xml");

            logentry = false;
            while (!logentry)
            {
                try
                {
                    SW = File.CreateText(mydocs + "\\Halfmode\\WordBuild.log");
                    SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Generating Word Report");
                    SW.Close();
                    logentry = true;
                }
                catch
                {
                }
            }


            btnWordGenerate.Text = "Creating Word Report";


            btnWordGenerate.Enabled = false;
            Thread wordtrd = new Thread(new ThreadStart(this.WordTread));
            this.wtrd = wordtrd;
            wtrd.IsBackground = true;
            wtrd.Start();

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            object oTrue = true;
            object oFalse = false;
            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;

            oWord = new Word.Application();
            //oWord.Visible = true

            //##################################






            #region Setup Document

            //use word template or not
            if (label11.Text != "N/A")
            {
                object oTemplate = template;
                oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
                ref oMissing, ref oMissing);
                oDoc.ShowSpellingErrors = false;
                oDoc.ShowGrammaticalErrors = false;

            }

            else
            {
                oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);
                oDoc.ShowSpellingErrors = false;
                oDoc.ShowGrammaticalErrors = false;

                //Header
                oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
                oWord.Selection.TypeParagraph();
                oWord.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                oWord.ActiveWindow.Selection.Font.Name = fontDialog1.Font.Name;
                oWord.ActiveWindow.Selection.TypeText(textBox11.Text);
                //move back to main document
                oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
            }
            //#####################################################

            //set doc border at top of the page

            if (checktitleborder.Checked)
            {
                oDoc.ActiveWindow.Selection.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oDoc.ActiveWindow.Selection.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth300pt;
                oDoc.ActiveWindow.Selection.Borders[Word.WdBorderType.wdBorderTop].Color = (Word.WdColor)objcoltit; //objcoltit is defined by color picker
            }
            //set page border

            if (checkPageBorder1.Checked)
            {
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderLeft].Color = (Word.WdColor)objcolbord;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderRight].Color = (Word.WdColor)objcolbord;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderTop].Color = (Word.WdColor)objcolbord;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oWord.ActiveDocument.Sections[1].Borders[Word.WdBorderType.wdBorderBottom].Color = (Word.WdColor)objcolbord;
            }

            //footer page number
            if (pageNumber.Checked)
            {
                oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                oWord.Selection.TypeParagraph();
                oWord.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                oWord.ActiveWindow.Selection.TypeText("Page ");
                Object CurrentPage = Word.WdFieldType.wdFieldPage;
                oWord.ActiveWindow.Selection.Fields.Add(oWord.Selection.Range, ref CurrentPage, ref oMissing, ref oMissing);
            }
            //move back to main document
            oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            //commit formatting
            oWord.ActiveDocument.Sections[1].Borders.ApplyPageBordersToAllSections();

            //drop 3 lines
            oWord.Selection.TypeParagraph();
            //oWord.Selection.TypeParagraph();
            //oWord.Selection.TypeParagraph();

            // Center text then print the text after it
            oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oWord.Selection.Font.Size = 34;
            oWord.Selection.Font.Name = textBox12.Font.Name;
            oWord.Selection.TypeText(textBox12.Text);
            oWord.Selection.TypeParagraph();
            oWord.Selection.Font.Size = 14;
            oWord.Selection.TypeParagraph();
            oWord.Selection.Font.Name = textBox13.Font.Name;
            oWord.Selection.TypeText(textBox13.Text);
            oWord.Selection.TypeParagraph();


            //make text left
            oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();


            //list name, company and address on first page
            object styleNoSpacing = Word.WdBuiltinStyle.wdStylePlainText;
            oWord.Selection.Font.Size = 10;
            oWord.Selection.set_Style(ref styleNoSpacing);
            oWord.Selection.Font.Name = "Tahoma";
            oWord.Selection.TypeText("Prepared by");
            oWord.Selection.Font.Bold = 1;
            oWord.Selection.TypeText("\n" + textBox14.Text);
            oWord.Selection.TypeText("\n" + textBox15.Text);
            oWord.Selection.TypeText("\n" + textBox16.Text);
            if (textBox17.Text != "")
            {
                oWord.Selection.TypeText("\n" + textBox17.Text);
            }
            oWord.Selection.TypeText("\n" + textBox18.Text);
            oWord.Selection.TypeText("\n" + textBox19.Text);
            oWord.Selection.TypeText("\n" + textBox20.Text);
            oWord.Selection.TypeText("\n" + textBox21.Text);


            // move to next page:
            object breakPage = Word.WdBreakType.wdSectionBreakNextPage;
            oWord.Selection.InsertBreak(ref breakPage);

            oWord.Selection.Font.Size = 12;
            oWord.Selection.TypeText("Table of Contents");
            oWord.Selection.TypeParagraph();
            oWord.Selection.TypeParagraph();

            //insert TOC

            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "3";
            Object oTOCTableID = "TableOfContents";
            Word.Range rngTOC = oWord.Selection.Range;
            oDoc.TablesOfContents.Add(rngTOC, ref oTrue, ref oUpperHeadingLevel,
                ref oLowerHeadingLevel, ref oMissing, ref oTOCTableID, ref oTrue,
                ref oTrue, ref oMissing, ref oTrue, ref oTrue, ref oTrue);

            // move to next page:
            oWord.Selection.InsertBreak(ref breakPage);

            logentry = false;
            while (!logentry)
            {
                try
                {
                    SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                    SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Word Formatting Complete");
                    SW.Close();
                    logentry = true;
                }
                catch
                {
                }
            }


            #endregion

            #region first section VMS
            // Set first heading
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            Object styleHeading1 = "Heading 1"; //defines object to define heading level


            oPara1.Range.Text = "Virtual Machines";
            oPara1.Range.Font.Name = "Tahoma";
            oPara1.Range.set_Style(ref styleHeading1); // sets the text as a heading

            //this line ends the para - needed for formating
            oPara1.Range.InsertParagraphAfter();

            //check if 2003 or 2007

            RegistryKey RK = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Office\\12.0");
            object breakline = Word.WdBreakType.wdLineBreak;
            oPara1.Range.InsertBreak(ref breakline);

            if (RK != null)
            {
                //not 2007
                oPara1.Range.set_Style(ref styleNoSpacing);
            }
            else
            {

                // is 2007

            }
            if (vmlist.Count >= 2)
            {

                Word.Table otable1;
                Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                otable1 = oDoc.Tables.Add(wrdRng, Convert.ToInt32(vmlist.Count / 2), 2, ref oMissing, ref oMissing);
                otable1.Range.Paragraphs.SpaceAfter = 6;
                otable1.Range.Font.Bold = 0;
                otable1.Range.Font.Name = "Ariel";
                int listcount = 0;
                for (int count = 1; count <= Convert.ToInt32(vmlist.Count / 2); )
                {

                    otable1.Cell(count, 1).Range.Text = (string)vmlist[listcount];
                    otable1.Cell(count, 2).Range.Text = (string)vmlist[listcount + 1];
                    count++;
                    listcount = listcount + 2;



                }

                //make table heading bolds
                string setbold;
                for (int tc = 1; tc <= Convert.ToInt32(vmlist.Count / 2); tc++)
                {
                    setbold = otable1.Cell(tc, 1).Range.Text;
                    if (setbold == "Virtual Machine Name:\r\a")
                    {
                        otable1.Cell(tc, 1).Range.Font.Bold = 1;
                    }

                }


                otable1.Borders.Enable = 1;

                logentry = false;
                while (!logentry)
                {
                    try
                    {
                        SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Virtual Machine Table Complete");
                        SW.Close();
                        logentry = true;
                    }
                    catch
                    {
                    }
                }
            }
            else
            {
                Word.Table otable1;
                Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                otable1 = oDoc.Tables.Add(wrdRng, 1, 2, ref oMissing, ref oMissing);
                otable1.Range.Paragraphs.SpaceAfter = 6;
                otable1.Range.Font.Bold = 0;
                otable1.Range.Font.Name = "Ariel";
                otable1.Cell(1, 1).Range.Text = "No Information Collected";
                otable1.Cell(1, 2).Range.Text = "Table Not Processed";

                logentry = false;
                while (!logentry)
                {
                    try
                    {
                        SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                        SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Virtual Machine Table Not Processed");
                        SW.Close();
                        logentry = true;
                    }
                    catch
                    {
                    }

                }
            }
            #endregion

                #region second section VM VIFs
                Word.Paragraph oPara2;
                oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara2.Range.Text = "Virtual Machine Virtual NICs (VIF)";
                oPara2.Range.Font.Name = "Tahoma";
                oPara2.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara2.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara2.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara2.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara2.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }

                //################ next table ##############################

                if (viflist.Count >= 2)
                {
                    Word.Table otable2;
                    Word.Range wrdRng2 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable2 = oDoc.Tables.Add(wrdRng2, Convert.ToInt32(viflist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable2.Range.Paragraphs.SpaceAfter = 6;
                    otable2.Range.Font.Bold = 0;
                    otable2.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(viflist.Count / 2); )
                    {

                        otable2.Cell(count, 1).Range.Text = (string)viflist[listcount];
                        otable2.Cell(count, 2).Range.Text = (string)viflist[listcount + 1];
                        count++;
                        listcount = listcount + 2;



                    }
                    //make table heading bolds

                    string vifsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(viflist.Count / 2); tc++)
                    {
                        vifsetbold = otable2.Cell(tc, 1).Range.Text;
                        if (vifsetbold == "Virtual Machine Name:\r\a")
                        {
                            otable2.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    otable2.Borders.Enable = 1;

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VIF Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable2;
                    Word.Range wrdRng2 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable2 = oDoc.Tables.Add(wrdRng2, 1, 2, ref oMissing, ref oMissing);
                    otable2.Range.Paragraphs.SpaceAfter = 6;
                    otable2.Range.Font.Bold = 0;
                    otable2.Range.Font.Name = "Ariel";
                    otable2.Cell(1, 1).Range.Text = "No Information Collected";
                    otable2.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VIF Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region Host Info
                //section 3 host info

                Word.Paragraph oPara3;
                oPara3 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara3.Range.Text = "Host Configuration";
                oPara3.Range.Font.Name = "Tahoma";
                oPara3.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara3.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara3.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara3.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara3.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }

                if (hostlist.Count >= 2)
                {
                    Word.Table otable3;
                    Word.Range wrdRng3 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable3 = oDoc.Tables.Add(wrdRng3, Convert.ToInt32(hostlist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable3.Range.Paragraphs.SpaceAfter = 6;
                    otable3.Range.Font.Bold = 0;
                    otable3.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(hostlist.Count / 2); )
                    {

                        otable3.Cell(count, 1).Range.Text = (string)hostlist[listcount];
                        otable3.Cell(count, 2).Range.Text = (string)hostlist[listcount + 1];
                        count++;
                        listcount = listcount + 2;




                    }
                    otable3.Borders.Enable = 1;

                    //make table heading bolds
                    string hostsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(hostlist.Count / 2); tc++)
                    {
                        hostsetbold = otable3.Cell(tc, 1).Range.Text;
                        if (hostsetbold == "Xenserver Name:\r\a")
                        {
                            otable3.Cell(tc, 1).Range.Font.Bold = 1;
                        }
                        else if (hostsetbold == "Number of Physical Block Devices (PBD):\r\a")
                        {
                            otable3.Cell(tc, 1).Range.Font.Bold = 1;
                        }
                        else if (hostsetbold == "Number of Physical NICs (PIF):\r\a")
                        {
                            otable3.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Host Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable3;
                    Word.Range wrdRng3 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable3 = oDoc.Tables.Add(wrdRng3, 1, 2, ref oMissing, ref oMissing);
                    otable3.Range.Paragraphs.SpaceAfter = 6;
                    otable3.Range.Font.Bold = 0;
                    otable3.Range.Font.Name = "Ariel";
                    otable3.Cell(1, 1).Range.Text = "No Information Collected";
                    otable3.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Host Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region BOND info
                Word.Paragraph oPara3a;
                oPara3a = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara3a.Range.Text = "NIC Bond Configuration";
                oPara3a.Range.Font.Name = "Tahoma";
                oPara3a.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara3a.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara3a.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara3a.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara3a.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }

                if (bondlist.Count >= 2)
                {
                    Word.Table otable3a;
                    Word.Range wrdRng3a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable3a = oDoc.Tables.Add(wrdRng3a, Convert.ToInt32(bondlist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable3a.Range.Paragraphs.SpaceAfter = 6;
                    otable3a.Range.Font.Bold = 0;
                    otable3a.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(bondlist.Count / 2); )
                    {

                        otable3a.Cell(count, 1).Range.Text = (string)bondlist[listcount];
                        otable3a.Cell(count, 2).Range.Text = (string)bondlist[listcount + 1];
                        count++;
                        listcount = listcount + 2;




                    }
                    otable3a.Borders.Enable = 1;

                    //make table heading bolds
                    string bondsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(bondlist.Count / 2); tc++)
                    {
                        bondsetbold = otable3a.Cell(tc, 1).Range.Text;
                        if (bondsetbold == "Bond Number:\r\a")
                        {
                            otable3a.Cell(tc, 1).Range.Font.Bold = 1;
                        }


                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Bond Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable3a;
                    Word.Range wrdRng3a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable3a = oDoc.Tables.Add(wrdRng3a, 1, 2, ref oMissing, ref oMissing);
                    otable3a.Range.Paragraphs.SpaceAfter = 6;
                    otable3a.Range.Font.Bold = 0;
                    otable3a.Range.Font.Name = "Ariel";
                    otable3a.Cell(1, 1).Range.Text = "No Information Collected";
                    otable3a.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Bond Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion


                #region Network info
                Word.Paragraph oPara3b;
                oPara3b = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara3b.Range.Text = "Network Configuration";
                oPara3b.Range.Font.Name = "Tahoma";
                oPara3b.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara3b.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara3b.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara3b.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara3b.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (netlist.Count >= 2)
                {
                    Word.Table otable3b;
                    Word.Range wrdRng3b = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable3b = oDoc.Tables.Add(wrdRng3b, Convert.ToInt32(netlist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable3b.Range.Paragraphs.SpaceAfter = 6;
                    otable3b.Range.Font.Bold = 0;
                    otable3b.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(netlist.Count / 2); )
                    {

                        otable3b.Cell(count, 1).Range.Text = (string)netlist[listcount];
                        otable3b.Cell(count, 2).Range.Text = (string)netlist[listcount + 1];
                        count++;
                        listcount = listcount + 2;




                    }
                    otable3b.Borders.Enable = 1;

                    //make table heading bolds
                    string netsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(netlist.Count / 2); tc++)
                    {
                        netsetbold = otable3b.Cell(tc, 1).Range.Text;
                        if (netsetbold == "Network Name:\r\a")
                        {
                            otable3b.Cell(tc, 1).Range.Font.Bold = 1;
                        }


                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Network Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable3b;
                    Word.Range wrdRng3b = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable3b = oDoc.Tables.Add(wrdRng3b, 1, 2, ref oMissing, ref oMissing);
                    otable3b.Range.Paragraphs.SpaceAfter = 6;
                    otable3b.Range.Font.Bold = 0;
                    otable3b.Range.Font.Name = "Ariel";
                    otable3b.Cell(1, 1).Range.Text = "No Information Collected";
                    otable3b.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Network Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region SR info
                //section 4 SR info

                Word.Paragraph oPara4;
                oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara4.Range.Text = "Storage Repository";
                oPara4.Range.Font.Name = "Tahoma";
                oPara4.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara4.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara4.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara4.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara4.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (srlist.Count >= 2)
                {
                    Word.Table otable4;
                    Word.Range wrdRng4 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable4 = oDoc.Tables.Add(wrdRng4, Convert.ToInt32(srlist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable4.Range.Paragraphs.SpaceAfter = 6;
                    otable4.Range.Font.Bold = 0;
                    otable4.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(srlist.Count / 2); )
                    {

                        otable4.Cell(count, 1).Range.Text = (string)srlist[listcount];
                        otable4.Cell(count, 2).Range.Text = (string)srlist[listcount + 1];
                        count++;
                        listcount = listcount + 2;




                    }

                    otable4.Borders.Enable = 1;
                    //make table heading bolds
                    string SRsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(srlist.Count / 2); tc++)
                    {
                        SRsetbold = otable4.Cell(tc, 1).Range.Text;
                        if (SRsetbold == "Storage Repository Name:\r\a")
                        {
                            otable4.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                        if (SRsetbold == "Xenserver Name:\r\a")
                        {
                            otable4.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " SR Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable4;
                    Word.Range wrdRng4 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable4 = oDoc.Tables.Add(wrdRng4, 1, 2, ref oMissing, ref oMissing);
                    otable4.Range.Paragraphs.SpaceAfter = 6;
                    otable4.Range.Font.Bold = 0;
                    otable4.Range.Font.Name = "Ariel";
                    otable4.Cell(1, 1).Range.Text = "No Information Collected";
                    otable4.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " SR Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region Resident VMs
                // section 5 resident vms
                Word.Paragraph oPara5;
                oPara5 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara5.Range.Text = "Hosts with Running VMs";
                oPara5.Range.Font.Name = "Tahoma";
                oPara5.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara5.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara5.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara5.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara5.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (reslist.Count >= 2)
                {
                    Word.Table otable5;
                    Word.Range wrdRng5 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable5 = oDoc.Tables.Add(wrdRng5, Convert.ToInt32(reslist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable5.Range.Paragraphs.SpaceAfter = 6;
                    otable5.Range.Font.Bold = 0;
                    otable5.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(reslist.Count / 2); )
                    {

                        otable5.Cell(count, 1).Range.Text = (string)reslist[listcount];
                        otable5.Cell(count, 2).Range.Text = (string)reslist[listcount + 1];
                        count++;
                        listcount = listcount + 2;





                    }
                    otable5.Borders.Enable = 1;
                    string ressetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(reslist.Count / 2); tc++)
                    {
                        ressetbold = otable5.Cell(tc, 1).Range.Text;
                        if (ressetbold == "XenServer Name:\r\a")
                        {
                            otable5.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Resident VM Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable5;
                    Word.Range wrdRng5 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable5 = oDoc.Tables.Add(wrdRng5, 1, 2, ref oMissing, ref oMissing);
                    otable5.Range.Paragraphs.SpaceAfter = 6;
                    otable5.Range.Font.Bold = 0;
                    otable5.Range.Font.Name = "Ariel";
                    otable5.Cell(1, 1).Range.Text = "No Information Collected";
                    otable5.Cell(1, 2).Range.Text = "Table Not Processed";
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Resident VM Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region snapshot parent

                Word.Paragraph oPara6;
                oPara6 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara6.Range.Text = "VM with Snapshots";
                oPara6.Range.Font.Name = "Tahoma";
                oPara6.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara6.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara6.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara6.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara6.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (snaplist.Count >= 2)
                {
                    Word.Table otable6;
                    Word.Range wrdRng6 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable6 = oDoc.Tables.Add(wrdRng6, Convert.ToInt32(snaplist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable6.Range.Paragraphs.SpaceAfter = 6;
                    otable6.Range.Font.Bold = 0;
                    otable6.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(snaplist.Count / 2); )
                    {

                        otable6.Cell(count, 1).Range.Text = (string)snaplist[listcount];
                        otable6.Cell(count, 2).Range.Text = (string)snaplist[listcount + 1];
                        count++;
                        listcount = listcount + 2;


                    }
                    otable6.Borders.Enable = 1;
                    string snapsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(snaplist.Count / 2); tc++)
                    {
                        snapsetbold = otable6.Cell(tc, 1).Range.Text;
                        if (snapsetbold == "Virtual Machine Name:\r\a")
                        {
                            otable6.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Snapshot Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable6;
                    Word.Range wrdRng6 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable6 = oDoc.Tables.Add(wrdRng6, 1, 2, ref oMissing, ref oMissing);
                    otable6.Range.Paragraphs.SpaceAfter = 6;
                    otable6.Range.Font.Bold = 0;
                    otable6.Range.Font.Name = "Ariel";
                    otable6.Cell(1, 1).Range.Text = "No Information Collected";
                    otable6.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Snapshot Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }

                #endregion

                #region VDI

                Word.Paragraph oPara6a;
                oPara6a = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara6a.Range.Text = "VDI Virtual Disk Images";
                oPara6a.Range.Font.Name = "Tahoma";
                oPara6a.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara6a.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara6a.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara6a.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara6a.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (vdilist.Count >= 2)
                {
                    Word.Table otable6a;
                    Word.Range wrdRng6a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable6a = oDoc.Tables.Add(wrdRng6a, Convert.ToInt32(vdilist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable6a.Range.Paragraphs.SpaceAfter = 6;
                    otable6a.Range.Font.Bold = 0;
                    otable6a.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(vdilist.Count / 2); )
                    {

                        otable6a.Cell(count, 1).Range.Text = (string)vdilist[listcount];
                        otable6a.Cell(count, 2).Range.Text = (string)vdilist[listcount + 1];
                        count++;
                        listcount = listcount + 2;


                    }
                    otable6a.Borders.Enable = 1;
                    string vdisetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(vdilist.Count / 2); tc++)
                    {
                        vdisetbold = otable6a.Cell(tc, 1).Range.Text;
                        if (vdisetbold == "Virtual Machine Name:\r\a")
                        {
                            otable6a.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }


                }
                else
                {
                    Word.Table otable6a;
                    Word.Range wrdRng6a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable6a = oDoc.Tables.Add(wrdRng6a, 1, 2, ref oMissing, ref oMissing);
                    otable6a.Range.Paragraphs.SpaceAfter = 6;
                    otable6a.Range.Font.Bold = 0;
                    otable6a.Range.Font.Name = "Ariel";
                    otable6a.Cell(1, 1).Range.Text = "No Information Collected";
                    otable6a.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VDI Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region VM VBDs

                Word.Paragraph oPara7;
                oPara7 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara7.Range.Text = "VM with Virtual Block Devices";
                oPara7.Range.Font.Name = "Tahoma";
                oPara7.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara7.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara7.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara7.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara7.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (vdilist.Count >= 2)
                {
                    Word.Table otable7;
                    Word.Range wrdRng7 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable7 = oDoc.Tables.Add(wrdRng7, Convert.ToInt32(vbdlist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable7.Range.Paragraphs.SpaceAfter = 6;
                    otable7.Range.Font.Bold = 0;
                    otable7.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(vbdlist.Count / 2); )
                    {

                        otable7.Cell(count, 1).Range.Text = (string)vbdlist[listcount];
                        otable7.Cell(count, 2).Range.Text = (string)vbdlist[listcount + 1];
                        count++;
                        listcount = listcount + 2;


                    }
                    otable7.Borders.Enable = 1;
                    string vbdsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(vbdlist.Count / 2); tc++)
                    {
                        vbdsetbold = otable7.Cell(tc, 1).Range.Text;
                        if (vbdsetbold == "Virtual Machine Name:\r\a")
                        {
                            otable7.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VBD Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable7;
                    Word.Range wrdRng7 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable7 = oDoc.Tables.Add(wrdRng7, 1, 2, ref oMissing, ref oMissing);
                    otable7.Range.Paragraphs.SpaceAfter = 6;
                    otable7.Range.Font.Bold = 0;
                    otable7.Range.Font.Name = "Ariel";
                    otable7.Cell(1, 1).Range.Text = "No Information Collected";
                    otable7.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " VBD Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region pool list

                Word.Paragraph oPara7a;
                oPara7a = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara7a.Range.Text = "Pool Information";
                oPara7a.Range.Font.Name = "Tahoma";
                oPara7a.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara7a.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara7a.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara7a.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara7a.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (poollist.Count >= 2)
                {
                    Word.Table otable7a;
                    Word.Range wrdRng7a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable7a = oDoc.Tables.Add(wrdRng7a, Convert.ToInt32(poollist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable7a.Range.Paragraphs.SpaceAfter = 6;
                    otable7a.Range.Font.Bold = 0;
                    otable7a.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(poollist.Count / 2); )
                    {

                        otable7a.Cell(count, 1).Range.Text = (string)poollist[listcount];
                        otable7a.Cell(count, 2).Range.Text = (string)poollist[listcount + 1];
                        count++;
                        listcount = listcount + 2;


                    }
                    otable7a.Borders.Enable = 1;
                    string poolsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(poollist.Count / 2); tc++)
                    {
                        poolsetbold = otable7a.Cell(tc, 1).Range.Text;
                        if (poolsetbold == "Pool Name:\r\a")
                        {
                            otable7a.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Pool Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable7a;
                    Word.Range wrdRng7a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable7a = oDoc.Tables.Add(wrdRng7a, 1, 2, ref oMissing, ref oMissing);
                    otable7a.Range.Paragraphs.SpaceAfter = 6;
                    otable7a.Range.Font.Bold = 0;
                    otable7a.Range.Font.Name = "Ariel";
                    otable7a.Cell(1, 1).Range.Text = "No Information Collected";
                    otable7a.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Pool Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region Template list

                Word.Paragraph oPara8;
                oPara8 = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara8.Range.Text = "Template List";
                oPara8.Range.Font.Name = "Tahoma";
                oPara8.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara8.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara8.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara8.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara8.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (templatelist.Count >= 2)
                {
                    Word.Table otable8;
                    Word.Range wrdRng8 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable8 = oDoc.Tables.Add(wrdRng8, Convert.ToInt32(templatelist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable8.Range.Paragraphs.SpaceAfter = 6;
                    otable8.Range.Font.Bold = 0;
                    otable8.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(templatelist.Count / 2); )
                    {

                        otable8.Cell(count, 1).Range.Text = (string)templatelist[listcount];
                        otable8.Cell(count, 2).Range.Text = (string)templatelist[listcount + 1];
                        count++;
                        listcount = listcount + 2;


                    }
                    otable8.Borders.Enable = 1;
                    string tempsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(templatelist.Count / 2); tc++)
                    {
                        tempsetbold = otable8.Cell(tc, 1).Range.Text;
                        if (tempsetbold == "Template Name:\r\a")
                        {
                            otable8.Cell(tc, 1).Range.Font.Bold = 1;
                        }

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Template Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable8;
                    Word.Range wrdRng8 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable8 = oDoc.Tables.Add(wrdRng8, 1, 2, ref oMissing, ref oMissing);
                    otable8.Range.Paragraphs.SpaceAfter = 6;
                    otable8.Range.Font.Bold = 0;
                    otable8.Range.Font.Name = "Ariel";
                    otable8.Cell(1, 1).Range.Text = "No Information Collected";
                    otable8.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Template Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion

                #region terms

                Word.Paragraph oPara8a;
                oPara8a = oDoc.Content.Paragraphs.Add(ref oMissing);

                oPara8a.Range.Text = "Document Terms";
                oPara8a.Range.Font.Name = "Tahoma";
                oPara8a.Range.set_Style(ref styleHeading1);


                // move to next page:
                oPara8a.Range.InsertBreak(ref breakPage);

                //this line ends the para - needed for formating
                oPara8a.Range.InsertParagraphAfter();

                //this line puts a line between heading and table
                oPara8a.Range.InsertBreak(ref breakline);

                if (RK != null)
                {
                    //not 2007
                    oPara8a.Range.set_Style(ref styleNoSpacing);
                }
                else
                {

                    // is 2007

                }
                if (termlist.Count >= 2)
                {
                    Word.Table otable8a;
                    Word.Range wrdRng8a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable8a = oDoc.Tables.Add(wrdRng8a, Convert.ToInt32(termlist.Count / 2), 2, ref oMissing, ref oMissing);
                    otable8a.Range.Paragraphs.SpaceAfter = 6;
                    otable8a.Range.Font.Bold = 0;
                    otable8a.Range.Font.Name = "Ariel";

                    int listcount = 0;
                    for (int count = 1; count <= Convert.ToInt32(termlist.Count / 2); )
                    {

                        otable8a.Cell(count, 1).Range.Text = (string)termlist[listcount];
                        otable8a.Cell(count, 2).Range.Text = (string)termlist[listcount + 1];
                        count++;
                        listcount = listcount + 2;


                    }
                    otable8a.Borders.Enable = 1;
                    string termsetbold;
                    for (int tc = 1; tc <= Convert.ToInt32(termlist.Count / 2); tc++)
                    {
                        termsetbold = otable8a.Cell(tc, 1).Range.Text;

                        otable8a.Cell(tc, 1).Range.Font.Bold = 1;

                    }
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Term Table Complete");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    Word.Table otable8a;
                    Word.Range wrdRng8a = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    otable8a = oDoc.Tables.Add(wrdRng8a, 1, 2, ref oMissing, ref oMissing);
                    otable8a.Range.Paragraphs.SpaceAfter = 6;
                    otable8a.Range.Font.Bold = 0;
                    otable8a.Range.Font.Name = "Ariel";
                    otable8a.Cell(1, 1).Range.Text = "No Information Collected";
                    otable8a.Cell(1, 2).Range.Text = "Table Not Processed";

                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW = File.AppendText(mydocs + "\\Halfmode\\WordBuild.log");
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Term Table Not Processed");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                #endregion


                //UPDATING THE TABLE OF CONTENTS
                oDoc.TablesOfContents[1].Update();

                //UPDATING THE TABLE OF CONTENTS
                oDoc.TablesOfContents[1].UpdatePageNumbers();


                //kill off the busy graphic thread.
                wtrd.Abort();
                this.Show();
                //minimize main app and bring to life word.
                this.WindowState = FormWindowState.Minimized;
                oWord.Visible = true;
                //goto first page 
                object item = Word.WdGoToItem.wdGoToLine;
                object whichItem = Word.WdGoToDirection.wdGoToFirst;
                oWord.Selection.GoTo(ref item, ref whichItem, ref oMissing, ref oMissing);
                //allow generate again
                btnWordGenerate.Enabled = true;
                btnWordGenerate.Text = "Generate Word report again";

            }
        

        private void button3_Click_1(object sender, EventArgs e)//first exit button
        {
            Application.Exit();
        }
      

        private void button4_Click(object sender, EventArgs e) //connect and collect
        {
            
            if (textBox2.Text == "")
            {
                MessageBox.Show("Please input a server name", "Server Missing");
            }
            else if (textBox4.Text == "")
            {
                MessageBox.Show("Please input a port number", "Port Missing");
            }
            else if (textBox5.Text == "")
            {
                MessageBox.Show("Please input a username", "User Missing");
            }
            else if (textBox6.Text == "")
            {
                MessageBox.Show("Please input a password", "Password Missing");
            }
            else
            {
                this.Hide();

                XmlTextWriter textwriter = new XmlTextWriter(mydocs + "\\Halfmode\\HalfmodeConnect.xml", null);
                textwriter.WriteStartDocument();
                textwriter.WriteStartElement("Connect");
                textwriter.WriteEndElement();
                textwriter.WriteEndDocument();
                textwriter.Close();

                string server = textBox2.Text;
                string port = textBox4.Text;
                string username = textBox5.Text;
                string password = textBox6.Text;

                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(mydocs + "\\Halfmode\\HalfmodeConnect.xml");
                XmlElement el = xmldoc.CreateElement("Settings");
                string Settings = "<Server>" + textBox2.Text + "</Server>" +
                    "<User>" + textBox5.Text + "</User>";
                el.InnerXml = Settings;
                xmldoc.DocumentElement.AppendChild(el);
                xmldoc.Save(mydocs + "\\Halfmode\\HalfmodeConnect.xml");

                bool connectok = false;
               
                
                
                        SW = File.CreateText(mydocs + "\\Halfmode\\HalfmodeConnection.log");
                        logentry = true;
                    
                

                button4.Enabled = false;
                try
                {
                    Session testsession = new Session(server, Convert.ToInt32(port));
                    testsession.login_with_password(username, password);
                    testsession.logout();
                    connectok = true;
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                    SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Connected OK");
                    SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Problem with Connection\n\nCheck your Settings", "Connection Error");
                    Application.Exit();
                    connectok = false;
                    logentry = false;
                    while (!logentry)
                    {
                        try
                        {
                            SW.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " Connection Error");
                            SW.Close();
                            logentry = true;
                        }
                        catch
                        {
                        }
                    }
                }

                if (connectok)
                {
                    Thread collecttrd = new Thread(new ThreadStart(this.ThreadTask));
                    this.trd = collecttrd;
                    trd.IsBackground = true;
                    trd.Start();

                    
                    

                    Session session = new Session(server, Convert.ToInt32(port));
                    session.login_with_password(username, password);

                    List<XenRef<VM>> vmRefs = VM.get_all(session);

                    VMCollector vmc = new VMCollector();
                    vmlist = (ArrayList)vmc.vmcollect(session, vmRefs);
                    collectcont++;

                    if (Convert.ToString(vmlist[0]) != "0")
                    {
                        // run other collections here.
                        hasconnected = true;
                        btnWordGenerate.Enabled = true;
                        btnWordGenerate.Text = "Generate Word Report";
                        try
                        {

                            Thread vifthread = new Thread(new ThreadStart(this.Threadvif));
                            this.trdvif = vifthread;
                            trdvif.IsBackground = true;
                            trdvif.Start();

                            
                           
                        }
                        catch
                        {
                            MessageBox.Show("Problem with VIF Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread hostthread = new Thread(new ThreadStart(this.Threadhost));
                            this.trdhost = hostthread;
                            trdhost.IsBackground = true;
                            trdhost.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Host Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread srthread = new Thread(new ThreadStart(this.Threadsr));
                            this.trdsr = srthread;
                            trdsr.IsBackground = true;
                            trdsr.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with SR Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread resthread = new Thread(new ThreadStart(this.Threadresvm));
                            this.trdresvm = resthread;
                            trdresvm.IsBackground = true;
                            trdresvm.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Resident VM Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread snapthread = new Thread(new ThreadStart(this.Threadsnap));
                            this.trdsnap = snapthread;
                            trdsnap.IsBackground = true;
                            trdsnap.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Snapshot Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread vbdthread = new Thread(new ThreadStart(this.Threadvbd));
                            this.trdvbd = vbdthread;
                            trdvbd.IsBackground = true;
                            trdvbd.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with VDB Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread tempthread = new Thread(new ThreadStart(this.Threadtemp));
                            this.trdtemp = tempthread;
                            trdtemp.IsBackground = true;
                            trdtemp.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Template Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread poolthread = new Thread(new ThreadStart(this.Threadpool));
                            this.trdpool = poolthread;
                            trdpool.IsBackground = true;
                            trdpool.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Pool Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            Thread bondthread = new Thread(new ThreadStart(this.Threadbond));
                            this.trdbond = bondthread;
                            trdbond.IsBackground = true;
                            trdbond.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Bond Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                            
                            Thread netthread = new Thread(new ThreadStart(this.Threadnet));
                            this.trdnet = netthread;
                            trdnet.IsBackground = true;
                            trdnet.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Network Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                             Thread vdithread = new Thread(new ThreadStart(this.Threadvdi));
                            this.trdvdi = vdithread;
                            trdvdi.IsBackground = true;
                            trdvdi.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with VDI Collecter", "Collector Error");
                            Application.Exit();
                        }
                        try
                        {
                             Thread termthread = new Thread(new ThreadStart(this.Threadterm));
                            this.trdterm = termthread;
                            trdterm.IsBackground = true;
                            trdterm.Start();
                        }
                        catch
                        {
                            MessageBox.Show("Problem with Term Collecter", "Collector Error");
                            Application.Exit();
                        }
                    }
                    else
                    {
                        //if here then connection settings wrong.
                        hasconnected = false;
                    }
                    timer1.Enabled = true;

                     
                }
            }
        }

        private void textPageBorderCol_TextChanged(object sender, EventArgs e)//not used
        {
            
        }

        private void textTitleBorder_TextChanged(object sender, EventArgs e)//not used
        {
            
        }

        private void textPageBorderCol_DoubleClick(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            int R = colorDialog1.Color.R;
            int G = colorDialog1.Color.G;
            int B = colorDialog1.Color.B;
            pagebordercolor = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            textPageBorderCol.BackColor = Color.FromArgb(R, G, B);
            int cl = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            objcolbord = cl;
        }

        private void textTitleBorder_DoubleClick(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            int R = colorDialog1.Color.R;
            int G = colorDialog1.Color.G;
            int B = colorDialog1.Color.B;
            titlebordercolor = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            textTitleBorder.BackColor = Color.FromArgb(R, G, B);
            int cl = Convert.ToInt32(ColorTranslator.ToOle(Color.FromArgb(R, G, B)).ToString());
            objcoltit = cl;
        }

        private void button5_Click(object sender, EventArgs e)//word exit button
        {
            Application.Exit();
        }

        private void button6_Click(object sender, EventArgs e)//header font
        {
            float currentsize;
            currentsize = textBox11.Font.Size;
            fontDialog1.ShowDialog();
            textBox11.Font = fontDialog1.Font;
            textBox11.Font = new Font(textBox11.Font.Name, currentsize);
        }

        private void button7_Click(object sender, EventArgs e)//title font
        {
            float currentsize;
            currentsize = textBox12.Font.Size;
            fontDialog2.ShowDialog();
            textBox12.Font = fontDialog2.Font;
            textBox12.Font = new Font(textBox12.Font.Name, currentsize);
        }

        private void button8_Click(object sender, EventArgs e)//sub title font
        {
            float currentsize;
            currentsize = textBox13.Font.Size;
            fontDialog3.ShowDialog();
            textBox13.Font = fontDialog3.Font;
            textBox13.Font = new Font(textBox13.Font.Name, currentsize);
        }

        private void Xenserver_Documenter_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo sinfo = new ProcessStartInfo("http://store.halfmode.com");

            System.Diagnostics.Process.Start(sinfo);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (collectcont >=13)
            {
                this.button4.Enabled = true;


                trd.Abort();
                this.Show();
                tabControl1.SelectedIndex = 1;
                timer1.Enabled = false;
            }
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }



            
        }

       
    }
