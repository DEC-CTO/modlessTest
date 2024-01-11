using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using ACadSharp.IO;
using ACadSharp;
using ACadSharp.Attributes;
using ACadSharp.IO.Templates;
using ACadSharp.Tables;
using ACadSharp.Entities;
using ACadSharp.Tables.Collections;

namespace modlessTest
{
    public partial class Main : Form
    {
        private RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;

        public Main(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();
            m_Handler = handler;
            m_ExEvent = exEvent;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            m_ExEvent.Dispose();
            m_ExEvent = null;
            m_Handler = null;

            base.OnFormClosed(e);
        }

        private void EnableCommands(bool status)
        {
            foreach (Control ctrl in this.Controls)
            {
                ctrl.Enabled = status;
            }
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
            Main main = new Main();
            main.DozeOff();
        }

        private void DozeOff()
        {
            EnableCommands(false);
        }
        public void WakeUp()
        {
            EnableCommands(true);
        }

        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.Test);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.gangpin);
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.readCAD);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.CreateBeams);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.deckslab);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.deckSlab2);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.deckSlab2);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.gang);
        }
    }
}
