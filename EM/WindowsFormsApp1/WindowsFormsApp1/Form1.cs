using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            guna2CircleProgressBar1.Value = 1; 
        }

        private void button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            guna2CircleProgressBar1.Value += 1;
            if(guna2CircleProgressBar1.Value == 100)
            {
                timer1.Enabled = false;

                //Opening second login form and close this form
                new Login().Show();
                this.Hide();
            }
        }
    }
}
