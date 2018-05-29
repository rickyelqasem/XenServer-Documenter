using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Halfmode_Xenserver_Documenter
{
    public partial class Wordbusy : Form
    {
        string mydocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string fileName = "\\Halfmode\\WordBuild.log";
        int count = 0;
        public Wordbusy()
        {
            InitializeComponent();
        }

        private void Wordbusy_Load(object sender, EventArgs e)
        {
            //Form XDform = Halfmode_Xenserver_Documenter.Xenserver_Documenter.ActiveForm;
            //Point parent = XDform.Location;
            //this.DesktopLocation = new Point(parent.X + 27, parent.Y + 473);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                richTextBox1.Text = System.IO.File.ReadAllText(mydocs + fileName, System.Text.Encoding.Default);
            }
            catch
            {
            }
            richTextBox1.Select(richTextBox1.Text.Length - 1, richTextBox1.Text.Length - 1);
            richTextBox1.ScrollToCaret();

            progressBar1.Value = count;
            if (count == 100)
            {
                count = 0;
            }
            else
            {
                count++;
            }
        }
    }
}