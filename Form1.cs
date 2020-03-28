using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFileRead
{
    public partial class Form1 : Form
    {
        string filepath;
        string EXt;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button2.Enabled = false;
            label1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                filepath = file.FileName;
                label1.Text = filepath;
                EXt = Path.GetExtension(filepath);
                if (EXt.CompareTo(".xls") == 0 || EXt.CompareTo(".xlsx") == 0)
                {
                    label1.Visible = true;
                    button2.Enabled = true;
                }

                else {
                    MessageBox.Show("Please Select .xls or .xlsx File only", "Warning", MessageBoxButtons.OKCancel,MessageBoxIcon.Warning);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable datafile = new DataTable();
            datafile = ReadFile(filepath, EXt);
            dataGridView1.DataSource = datafile;
        }

        public DataTable ReadFile(string FilePath, string Fileext)
        {
            string conn = "";
            DataTable dt = new DataTable();
            if (Fileext.CompareTo(".xls") == 0)
            {
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FilePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";
            }
            else
            {
                conn= @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties='Excel 12.0;HDR=NO';";
            }

            using (OleDbConnection ol = new OleDbConnection(conn))
            {
                OleDbDataAdapter oldata = new OleDbDataAdapter("Select * from [Sheet1$]", ol);
                oldata.Fill(dt);

            }
            return dt;


        }
    }
}
