using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookGPT
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (!File.Exists(System.Windows.Forms.Application.UserAppDataPath + "openaikey.dat")) return;
            using (StreamReader sr = new StreamReader(System.Windows.Forms.Application.UserAppDataPath + "openaikey.dat"))
            {
                textBox1.Text = sr.ReadToEnd();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if(textBox1.Text.Length < 5)
            {
                MessageBox.Show("I need an API key to work.", "Missing API Key", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else
            {
                using(StreamWriter sw = new StreamWriter(Application.UserAppDataPath + "openaikey.dat"))
                {
                    sw.WriteLine(textBox1.Text);
                }

                Form frm1 = new Form1();
                frm1.ShowDialog();
                this.Hide();
            }

        }
    }
}
