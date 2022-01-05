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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            
        }

        private void button_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(txtUserName.Text == "Admin" && txtUserPass.Text == "Marshall1301")
            {
                new MainMenu().Show();
                this.Close();
            }
            else
            {
                MessageBox.Show("The User name or password you entered is incorrect, try again");
                txtUserName.Clear();
                txtUserPass.Clear();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            txtUserName.Clear();
            txtUserPass.Clear();
        }

        private void txtUserName_TextChanged(object sender, EventArgs e)
        {
        }

        private void txtUserPass_TextChanged(object sender, EventArgs e)
        {
        }
    }
}
