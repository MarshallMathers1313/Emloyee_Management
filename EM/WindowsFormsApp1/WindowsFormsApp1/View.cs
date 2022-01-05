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
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace WindowsFormsApp1
{
    public partial class View : Form
    {
        public View()
        {
            InitializeComponent();
        }

        SqlConnection Con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\ashur\OneDrive\Рабочий стол\WindowsFormsApp1\WindowsFormsApp1\MyEmployeeDb.mdf;Integrated Security=True;Connect Timeout=30");

        private void fetchempdate()
        {
            Con.Open();
            string query = "select * from EmployeeTbl where EmpId='" + Search.Text + "'";
            SqlCommand cmd = new SqlCommand(query, Con);
            DataTable dt = new DataTable();
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt);
            foreach(DataRow dr in dt.Rows)
            {
                EmpIdTb.Text = dr["EmpId"].ToString();
                EmpNameTb.Text = dr["EmpName"].ToString();
                EmpAddTb.Text = dr["EmpAdd"].ToString();
                EmpPosTb.Text = dr["EmpPos"].ToString();
                EmpDOBTb.Text = dr["EmpDOB"].ToString();
                EmpPhoneTb.Text = dr["EmpPhone"].ToString();
                EmpEduTb.Text = dr["EmpEdu"].ToString();
                EmpGenTb.Text = dr["EmpGen"].ToString();
            }
            Con.Close();
        }

        private void View_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            fetchempdate();
        }

        private void button_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            new MainMenu().Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            create_document();
        }

        public void create_document()
        {
            //Creating a new word document
            WordDocument document = new WordDocument();
            //Adding a new section to the document
            WSection section = document.AddSection() as WSection;
            //Set margin for document
            section.PageSetup.Margins.All = 72;
            //Set page size of the section
            section.PageSetup.PageSize = new SizeF(612, 792);

            //Create Paragraph styles
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 11f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 8;
            style.ParagraphFormat.LineSpacing = 13.8f;

            style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
            style.ApplyBaseStyle("Normal");
            style.CharacterFormat.FontName = "Calibri Light";
            style.CharacterFormat.FontSize = 16f;
            style.CharacterFormat.TextColor = Color.FromArgb(46, 116, 181);
            style.ParagraphFormat.BeforeSpacing = 12;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.Keep = true;
            style.ParagraphFormat.KeepFollow = true;
            style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

            //Append paragraph
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Heading 1");
            paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            WTextRange textRange = paragraph.AppendText("EMPLOYEE SUMMARY\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 18f;
            textRange.CharacterFormat.FontName = "Calibri";

            //Append paragraph
            paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.FirstLineIndent = 36;
            paragraph.BreakCharacterFormat.FontSize = 14f;

            paragraph.AppendText("Employee ID: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpIdTb.Text + "\n");

            paragraph.AppendText("Employee Name: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpNameTb.Text + "\n");

            paragraph.AppendText("Employee Address: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpAddTb.Text + "\n");

            paragraph.AppendText("Employee Gender: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpGenTb.Text + "\n");

            paragraph.AppendText("Employee Position: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpPosTb.Text + "\n");

            paragraph.AppendText("Employee Phone: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpPhoneTb.Text + "\n");

            paragraph.AppendText("Employee DOB: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpDOBTb.Text + "\n");

            paragraph.AppendText("Employee Education: ").CharacterFormat.Bold = true;
            paragraph.AppendText(EmpEduTb.Text);

            //Save word document
            document.Save("example.docx");
            //Open word document
            System.Diagnostics.Process.Start("example.docx");
            document.Close();
        }
    }
}
