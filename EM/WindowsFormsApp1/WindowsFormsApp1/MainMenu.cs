using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    public partial class MainMenu : Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn
   (
       int nLeftRect,
       int nTopRect,
       int nRightRect,
       int nBottomRect,
       int nWidthEllipse,
       int nHeightEllipse
   );
        public MainMenu()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
        }

        private void button_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {

        }

        private void btnDashboard_Leave(object sender, EventArgs e)
        {
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void button2_Leave(object sender, EventArgs e)
        {
            btnEmployee.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnView_Leave(object sender, EventArgs e)
        {
            btnView.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnSalary_Leave(object sender, EventArgs e)
        {
            btnSalary.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnLogout_Leave(object sender, EventArgs e)
        {
            btnLogout.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {

        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnEmployee_Click(object sender, EventArgs e)
        {
            new ManageEmployee().Show();
            this.Close();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            new View().Show();
            this.Close();
        }

        private void btnSalary_Click(object sender, EventArgs e)
        {
            new Salary().Show();
            this.Close();
        }
    }
}
