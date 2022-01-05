using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    public partial class ManageEmployee : Form
    {
        public ManageEmployee()
        {
            InitializeComponent();
        }

        SqlConnection Con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ashur\OneDrive\Рабочий стол\WindowsFormsApp1\WindowsFormsApp1\MyEmployeeDb.mdf;Integrated Security=True;Connect Timeout=30");


        private void button_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            if (EmpIdTb.Text == "" || EmpNameTb.Text == "" || EmpAddTb.ToString() == "" || EmpPhoneTb.ToString() == "")
            {
                MessageBox.Show("Missing Information");
            }
            else
            {
                try
                {
                    Con.Open();
                    string query = "insert into EmployeeTbl values('" + EmpIdTb.Text + "','" + EmpNameTb.Text + "','" + EmpAddTb.Text + "','" + EmpPosCb.SelectedItem.ToString() + "','" + EmpDOB.Value.Date + "','" + EmpPhoneTb.Text + "','" + EmpEduCb.SelectedItem.ToString() + "','" + EmpGenCb.SelectedItem.ToString() + "')";
                    SqlCommand cmd= new SqlCommand(query, Con);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Employee Succesfully Added");
                    Con.Close();
                    populate();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void populate()
        {
            Con.Open();
            string query = "select * from EmployeeTbl";
            SqlDataAdapter sda = new SqlDataAdapter(query, Con);
            SqlCommandBuilder builder = new SqlCommandBuilder(sda);
            var ds = new DataSet();
            sda.Fill(ds);
            EmpDGV.DataSource = ds.Tables[0];
            Con.Close();
        }

        private void ManageEmployee_Load(object sender, EventArgs e)
        {
            populate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            new MainMenu().Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(EmpIdTb.Text =="")
            {
                MessageBox.Show("Enter The Employee Id");
            }
            else
            {
                try
                {
                    Con.Open();
                    string query = "delete from EmployeeTbl where Empid='" + EmpIdTb.Text + "';";
                    SqlCommand cmd = new SqlCommand(query, Con);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Employee Deleted Succesfully");
                    Con.Close();
                    populate();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void EmpDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            EmpIdTb.Text = EmpDGV.SelectedRows[0].Cells[0].Value.ToString();
            EmpNameTb.Text = EmpDGV.SelectedRows[0].Cells[1].Value.ToString();
            EmpAddTb.Text = EmpDGV.SelectedRows[0].Cells[2].Value.ToString();
            EmpGenCb.Text = EmpDGV.SelectedRows[0].Cells[7].Value.ToString();
            EmpDOB.Text = EmpDGV.SelectedRows[0].Cells[4].Value.ToString();
            EmpPosCb.Text = EmpDGV.SelectedRows[0].Cells[3].Value.ToString();
            EmpPhoneTb.Text = EmpDGV.SelectedRows[0].Cells[5].Value.ToString();
            EmpEduCb.Text = EmpDGV.SelectedRows[0].Cells[6].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (EmpIdTb.Text == "" || EmpNameTb.Text == "" || EmpAddTb.ToString() == "" || EmpPhoneTb.ToString() == "")
            {
                MessageBox.Show("Missing Information");
            }
            else
            {
                try
                {
                    Con.Open();
                    string query = "update EmployeeTbl set EmpName='" + EmpNameTb.Text + "',EmpAdd='" + EmpAddTb.Text + "',EmpPos='" + EmpPosCb.SelectedItem.ToString() + "',EmpDOB='" + EmpDOB.Value.Date + "',Empphone='" + EmpPhoneTb.Text + "',EmpEdu='" + EmpEduCb.SelectedItem.ToString() + "', EmpGen='" + EmpGenCb.SelectedItem.ToString() + "' where EmpId='" + EmpIdTb.Text + "';";
                    SqlCommand cmd = new SqlCommand(query, Con);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Employee Updated Successfully");
                    Con.Close();
                    populate();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
