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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Abiturient
{
    public partial class AddAbiturient : Form
    {
        public AddAbiturient()
        {
            InitializeComponent();
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

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox5.Text) || string.IsNullOrWhiteSpace(textBox6.Text)|| string.IsNullOrWhiteSpace(textBox8.Text)|| string.IsNullOrWhiteSpace(textBox9.Text)|| string.IsNullOrWhiteSpace(maskedTextBox4.Text)|| string.IsNullOrWhiteSpace(comboBox1.Text) || string.IsNullOrWhiteSpace(comboBox2.Text) || string.IsNullOrWhiteSpace(comboBox3.Text) || string.IsNullOrWhiteSpace(comboBox4.Text) || string.IsNullOrWhiteSpace(comboBox5.Text) || string.IsNullOrWhiteSpace(comboBox7.Text) || string.IsNullOrWhiteSpace(comboBox8.Text) || string.IsNullOrWhiteSpace(maskedTextBox1.Text) || string.IsNullOrWhiteSpace(maskedTextBox2.Text) || string.IsNullOrWhiteSpace(maskedTextBox3.Text) || string.IsNullOrWhiteSpace(dateTimePicker1.Text) || string.IsNullOrWhiteSpace(dateTimePicker2.Text)|| string.IsNullOrWhiteSpace(dateTimePicker3.Text))
            {
                MessageBox.Show("Заполните все поля!");
            }
            else {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("INSERT INTO [Abiturient] (FIO, Id_specialnost, Id_Forma_obuchenia, Gender, Birthday, Mesto_Rojdenia, Id_National, INN, Strahovoy_Nomer, Id_Ulitsa, Dom, Passport_Seria, Passport_Nomer, Passport_Kogda_Vidan, Passport_Kem_Vidan, Id_School, Lgota, Mother_FIO, Father_FIO, Avg_attestat, Data_zayavleniya, Nomer_zayavleniya, Ekzamen1, Ekzamen2) VALUES (@FIO, @Id_specialnost, @Id_Forma_obuchenia, @Gender, @Birthday, @Mesto_Rojdenia, @Id_National, @INN, @Strahovoy_Nomer, @Id_Ulitsa, @Dom, @Passport_Seria, @Passport_Nomer, @Passport_Kogda_Vidan, @Passport_Kem_Vidan, @Id_School, @Lgota, @Mother_FIO, @Father_FIO, @Avg_attestat, @Data_zayavleniya, @Nomer_zayavleniya, @Ekzamen1, @Ekzamen2)", con);
                command.Parameters.AddWithValue("@FIO", textBox1.Text.Trim());
                command.Parameters.AddWithValue("@Id_specialnost", comboBox1.SelectedValue);
                command.Parameters.AddWithValue("@Id_Forma_obuchenia", comboBox2.SelectedValue);
                command.Parameters.AddWithValue("@Gender", comboBox3.Text);
                command.Parameters.AddWithValue("@Birthday", dateTimePicker1.Text);
                command.Parameters.AddWithValue("@Mesto_Rojdenia", textBox2.Text.Trim());
                command.Parameters.AddWithValue("@Id_National", comboBox4.SelectedValue);
                command.Parameters.AddWithValue("@INN", maskedTextBox1.Text.Trim());
                command.Parameters.AddWithValue("@Strahovoy_nomer", maskedTextBox2.Text.Trim());
                command.Parameters.AddWithValue("@Id_Ulitsa", comboBox9.SelectedValue);
                command.Parameters.AddWithValue("@Dom", textBox3.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Seria", textBox4.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Nomer", textBox5.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Kogda_Vidan", textBox6.Text.Trim());
                command.Parameters.AddWithValue("@Passport_Kem_Vidan", dateTimePicker3.Text);
                command.Parameters.AddWithValue("@Id_School", comboBox5.SelectedValue);
                command.Parameters.AddWithValue("@Lgota", comboBox6.Text);
                command.Parameters.AddWithValue("@Mother_FIO", textBox8.Text.Trim());
                command.Parameters.AddWithValue("@Father_FIO", textBox9.Text.Trim());
                command.Parameters.AddWithValue("@Avg_attestat", maskedTextBox3.Text.ToString());
                command.Parameters.AddWithValue("@Data_zayavleniya", dateTimePicker2.Value.Date);
                command.Parameters.AddWithValue("@Nomer_zayavleniya", maskedTextBox4.Text.Trim());
                command.Parameters.AddWithValue("@Ekzamen1", comboBox7.Text);
                command.Parameters.AddWithValue("@Ekzamen2", comboBox8.Text);


                command.ExecuteNonQuery();
                con.Close();
                this.Hide();
            }
           
        }

        private void AddAbiturient_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet24.School". При необходимости она может быть перемещена или удалена.
            this.schoolTableAdapter.Fill(this.abiturientDataSet24.School);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet23.Ulitsa". При необходимости она может быть перемещена или удалена.
            this.ulitsaTableAdapter.Fill(this.abiturientDataSet23.Ulitsa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet22.National". При необходимости она может быть перемещена или удалена.
            this.nationalTableAdapter.Fill(this.abiturientDataSet22.National);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet21.Specialnost". При необходимости она может быть перемещена или удалена.
            this.specialnostTableAdapter.Fill(this.abiturientDataSet21.Specialnost);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet20.Forma_obuchenia". При необходимости она может быть перемещена или удалена.
            this.forma_obucheniaTableAdapter.Fill(this.abiturientDataSet20.Forma_obuchenia);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet18.School_RU". При необходимости она может быть перемещена или удалена.
            this.school_RUTableAdapter.Fill(this.abiturientDataSet18.School_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet17.Ulitsa_RU". При необходимости она может быть перемещена или удалена.
            this.ulitsa_RUTableAdapter.Fill(this.abiturientDataSet17.Ulitsa_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet16.National_RU". При необходимости она может быть перемещена или удалена.
            this.national_RUTableAdapter.Fill(this.abiturientDataSet16.National_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet15.Forma_obuchenia_RU". При необходимости она может быть перемещена или удалена.
            this.forma_obuchenia_RUTableAdapter.Fill(this.abiturientDataSet15.Forma_obuchenia_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet14.Specialnost_RU". При необходимости она может быть перемещена или удалена.
            this.specialnost_RUTableAdapter.Fill(this.abiturientDataSet14.Specialnost_RU);


            Get_list();
            Get_list2();
            Get_list3();
            Get_list4();
            Get_list5();
        }
    }
}
