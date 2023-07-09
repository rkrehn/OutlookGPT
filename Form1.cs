using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookGPT
{
    public partial class Form1 : Form
    {
        public string ResponseText {  get; set; }
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = ResponseText;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
