using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Abiturient
{
    public partial class Vhod : Form
    {
        public Vhod()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");

                string Sql = "Select * from Login where Login='" + LoginTextBox.Text.Trim() + "'" +
                    " and Password='" + passwordTextBox.Text.Trim() + "'";

                SqlDataAdapter sda = new SqlDataAdapter(Sql, conn);
                DataTable dt = new DataTable();
                DataSet ds = new DataSet();
                sda.Fill(dt);

                if (dt.Rows.Count == 1)
                {
                    this.Hide();
                    Glav glav = new Glav();
                    glav.Show();
                }
                else
                {
                    MessageBox.Show("Введите правильно логин и пароль.");
                }
            }
            catch
            {
                MessageBox.Show("Заполните все поля!");
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
