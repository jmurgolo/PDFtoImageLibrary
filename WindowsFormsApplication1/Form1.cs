using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GhostscriptSharp;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        //private readonly string TEST_FILE_LOCATION = "C:\\Users\\jmm\\Desktop\\98SEP28pg11_trimmed.pdf";
        //private readonly string SINGLE_FILE_LOCATION = "C:\\Users\\jmm\\Desktop\\output.jpg";
        //private readonly string MULTIPLE_FILE_LOCATION = "C:\\Users\\jmm\\Desktop\\output%d.jpg";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //pictureBox1.Image = GhostscriptWrapper.GeneratePageThumb(textBox1.Text, 1, 100, 100, 612, 792);
            //Assert.IsTrue(File.Exists(SINGLE_FILE_LOCATION));
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
