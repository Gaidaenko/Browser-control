using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.IO;
using System.Text;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using System.Drawing;
using CsQuery.ExtensionMethods.Internal;

namespace control
{
    public partial class Form1 : Form
    {       
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFile open = new OpenFile();
            open.openFile();           
        }
    }
}
