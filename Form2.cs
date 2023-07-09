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
                Properties.Settings.Default.OpenAPI = textBox1.Text;
                Properties.Settings.Default.Save();
                this.Hide();
            }

        }
    }
}
