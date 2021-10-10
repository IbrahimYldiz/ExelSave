using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ExelSave
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }//‪C:\Users\fenny\Desktop\deneme.xlsx
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\fenny\Desktop\deneme.xlsx;Extended Properties='Excel 12.0 Xml; HDR=YES';");
        void list()
        {
            OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Sayfa1$]", connection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            list();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("insert into [Sayfa1$] (Plaka,Adı) values (@p,@p2)", connection);
            command.Parameters.AddWithValue("@p1", textBox2.Text);
            command.Parameters.AddWithValue("@p2", textBox1.Text);
            command.ExecuteNonQuery();
            connection.Close();
            MessageBox.Show("Kayıt Başarılı");
            list(); 
        }
    }
}
