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
    public partial class EditAbiturient : Form
    {
        public int id_test;
        public EditAbiturient(int id)
        {
            InitializeComponent();
            id_test = id;
        }

        void Get_list()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from Specialnost", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Specialnost", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Specialnost");
                    comboBox1.DataSource = null;
                    comboBox1.DataSource = ds.Tables["Specialnost"];
                    comboBox1.DisplayMember = "Name_Specialnost";
                    comboBox1.ValueMember = "Id";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        void Get_list2()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from Forma_obuchenia", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Forma_obuchenia", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Forma_obuchenia");
                    comboBox2.DataSource = null;
                    comboBox2.DataSource = ds.Tables["Forma_obuchenia"];
                    comboBox2.DisplayMember = "Name_Forma_obuchenia";
                    comboBox2.ValueMember = "Id";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
        void Get_list3()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from [National]", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM [National]", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "[National]");
                    comboBox4.DataSource = null;
                    comboBox4.DataSource = ds.Tables["[National]"];
                    comboBox4.DisplayMember = "Name_Nation";
                    comboBox4.ValueMember = "Id";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        void Get_list4()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from Ulitsa", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Ulitsa", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Ulitsa");
                    comboBox9.DataSource = null;
                    comboBox9.DataSource = ds.Tables["Ulitsa"];
                    comboBox9.DisplayMember = "Name_Ulitsa";
                    comboBox9.ValueMember = "Id";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        void Get_list5()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from School", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM School", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "School");
                    comboBox5.DataSource = null;
                    comboBox5.DataSource = ds.Tables["School"];
                    comboBox5.DisplayMember = "Name_School";
                    comboBox5.ValueMember = "Id";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox5.Text) || string.IsNullOrWhiteSpace(textBox6.Text)|| string.IsNullOrWhiteSpace(textBox8.Text)|| string.IsNullOrWhiteSpace(textBox9.Text)|| string.IsNullOrWhiteSpace(maskedTextBox4.Text)|| string.IsNullOrWhiteSpace(comboBox1.Text) || string.IsNullOrWhiteSpace(comboBox2.Text) || string.IsNullOrWhiteSpace(comboBox3.Text) || string.IsNullOrWhiteSpace(comboBox4.Text) || string.IsNullOrWhiteSpace(comboBox5.Text) || string.IsNullOrWhiteSpace(comboBox7.Text) || string.IsNullOrWhiteSpace(comboBox8.Text) || string.IsNullOrWhiteSpace(maskedTextBox1.Text) || string.IsNullOrWhiteSpace(maskedTextBox2.Text) || string.IsNullOrWhiteSpace(maskedTextBox3.Text) || string.IsNullOrWhiteSpace(dateTimePicker1.Text) || string.IsNullOrWhiteSpace(dateTimePicker2.Text)|| string.IsNullOrWhiteSpace(dateTimePicker3.Text))
            {
                MessageBox.Show("Заполните все поля!");
            }
            else
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("UPDATE [Abiturient] SET FIO = @FIO, Id_specialnost = @Id_specialnost, Id_Forma_obuchenia = @Id_Forma_obuchenia, Gender = @Gender, Birthday = @Birthday, Mesto_Rojdenia = @Mesto_Rojdenia, Id_National = @Id_National, INN = @INN, Strahovoy_Nomer = @Strahovoy_Nomer, Id_Ulitsa =@Id_Ulitsa, Dom = @Dom, Passport_Seria =@Passport_Seria, Passport_Nomer = @Passport_Nomer, Passport_Kogda_Vidan = @Passport_Kogda_Vidan, Passport_Kem_Vidan = @Passport_Kem_Vidan, Id_School = @Id_School, Lgota = @Lgota, Mother_FIO = @Mother_FIO, Father_FIO = @Father_FIO, Avg_attestat = @Avg_attestat, Data_zayavleniya = @Data_zayavleniya, Nomer_zayavleniya = @Nomer_zayavleniya, Ekzamen1 = @Ekzamen1, Ekzamen2 = @Ekzamen2 WHERE Id = @id", con);
                command.Parameters.AddWithValue("@Id", id_test);
                command.Parameters.AddWithValue("@FIO", textBox1.Text.Trim());
                command.Parameters.AddWithValue("@Id_specialnost", comboBox1.SelectedValue);
                command.Parameters.AddWithValue("@Id_Forma_obuchenia", comboBox2.SelectedValue);
                command.Parameters.AddWithValue("@Gender", comboBox3.Text);
                command.Parameters.AddWithValue("@Birthday", dateTimePicker1.Text);
                command.Parameters.AddWithValue("@Mesto_Rojdenia", textBox2.Text.Trim());
                command.Parameters.AddWithValue("@Id_National", comboBox4.SelectedValue);
                command.Parameters.AddWithValue("@INN", maskedTextBox1.Text.Trim());
                command.Parameters.AddWithValue("@Strahovoy_nomer", maskedTextBox2.Text.Trim());
                command.Parameters.AddWithValue("@Id_Ulitsa", comboBox4.SelectedValue);
                command.Parameters.AddWithValue("@Dom", textBox3.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Seria", textBox4.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Nomer", textBox5.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Kogda_Vidan", textBox6.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Kem_Vidan", dateTimePicker2.Text);
                command.Parameters.AddWithValue("@Id_School", comboBox5.SelectedValue);
                command.Parameters.AddWithValue("@Lgota", comboBox6.Text);
                command.Parameters.AddWithValue("@Mother_FIO", textBox8.Text.Trim());
                command.Parameters.AddWithValue("@Father_FIO", textBox9.Text.Trim());
                command.Parameters.AddWithValue("@Avg_attestat", maskedTextBox3.Text.ToString());
                command.Parameters.AddWithValue("@Data_zayavleniya", dateTimePicker3.Text);
                command.Parameters.AddWithValue("@Nomer_zayavleniya", maskedTextBox4.Text.Trim());
                command.Parameters.AddWithValue("@Ekzamen1", comboBox7.Text);
                command.Parameters.AddWithValue("@Ekzamen2", comboBox8.Text);

                command.ExecuteNonQuery();
                con.Close();
                this.Hide();
            }

        }

        private void EditAbiturient_Load(object sender, EventArgs e)
        {

            SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
            con.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM Abiturient WHERE Id = @id", con);
            command.Parameters.AddWithValue("@id", id_test);
            SqlDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                textBox1.Text = reader["FIO"].ToString();
                comboBox1.SelectedValue = reader["Id_specialnost"];
                comboBox2.SelectedValue = reader["Id_Forma_obuchenia"];
                comboBox3.Text = reader["Gender"].ToString();
                dateTimePicker1.Text = reader["Birthday"].ToString();
                textBox2.Text = reader["Mesto_Rojdenia"].ToString();
                comboBox4.SelectedValue = reader["Id_National"];
                maskedTextBox1.Text = reader["INN"].ToString();
                maskedTextBox2.Text = reader["Strahovoy_Nomer"].ToString();
                comboBox4.SelectedValue = reader["Id_Ulitsa"];
                textBox3.Text = reader["Dom"].ToString();
                textBox4.Text = reader["Passport_Seria"].ToString();
                textBox5.Text = reader["Passport_Nomer"].ToString();
                textBox6.Text = reader["Passport_Kogda_Vidan"].ToString();
                dateTimePicker2.Text = reader["Passport_Kem_Vidan"].ToString();
                comboBox5.SelectedValue = reader["Id_School"];
                comboBox6.Text = reader["Lgota"].ToString();
                textBox8.Text = reader["Mother_FIO"].ToString();
                textBox9.Text = reader["Father_FIO"].ToString();
                maskedTextBox3.Text = reader["Avg_attestat"].ToString();
                dateTimePicker3.Text = reader["Data_zayavleniya"].ToString();
                maskedTextBox4.Text = reader["Nomer_zayavleniya"].ToString();
                comboBox7.Text = reader["Ekzamen1"].ToString();
                comboBox8.Text = reader["Ekzamen2"].ToString();
            }
            con.Close();

            Get_list();
            Get_list2();
            Get_list3();
            Get_list4();
            Get_list5();
        }
    }
}
