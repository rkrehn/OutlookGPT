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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.OpenAPI;
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

                MessageBox.Show("For this update to take place, please close the new mail window and try again.", "OutlookGPT Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Hide();
            }

        }
    }
}
