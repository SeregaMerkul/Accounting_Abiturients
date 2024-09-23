using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Abiturient
{
    public partial class Glav : Form
    {
        public Glav()
        {
            InitializeComponent();
        }


        void Get_list10()
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
                    comboBox3.DataSource = null;
                    comboBox3.DataSource = ds.Tables["Specialnost"];
                    comboBox3.DisplayMember = "Name_Specialnost";
                    comboBox3.ValueMember = "Id";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox8.Text))
            {

                MessageBox.Show("Заполните все поля!");
            }
            else
            {


                bool found = false;
                foreach (DataGridViewRow row in dataGridView5.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[2].Value.ToString() == textBox8.Text.Trim() && row.Cells[1].Value.ToString() == comboBox2.Text.ToString())
                    {
                        MessageBox.Show("Такая улица уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("Update [Ulitsa] SET Name_ulitsa = @Name_ulitsa, Id_Nas_punkt = @Id_Nas_punkt WHERE Id= @id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView5[0, dataGridView5.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_ulitsa", textBox8.Text.Trim());
                    command.Parameters.AddWithValue("@Id_Nas_punkt", comboBox2.SelectedValue);
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list5();
                }
            }
        }

        

        

        private void Glav_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet19.Abiturient_RU". При необходимости она может быть перемещена или удалена.
            this.abiturient_RUTableAdapter2.Fill(this.abiturientDataSet19.Abiturient_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet13.Oblast_RU". При необходимости она может быть перемещена или удалена.
            this.oblast_RUTableAdapter1.Fill(this.abiturientDataSet13.Oblast_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet12.Nas_punkt". При необходимости она может быть перемещена или удалена.
            this.nas_punktTableAdapter.Fill(this.abiturientDataSet12.Nas_punkt);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet11.Oblast". При необходимости она может быть перемещена или удалена.
            this.oblastTableAdapter.Fill(this.abiturientDataSet11.Oblast);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet10.Specialnost_RU". При необходимости она может быть перемещена или удалена.
            this.specialnost_RUTableAdapter.Fill(this.abiturientDataSet10.Specialnost_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet9.Ulitsa_RU". При необходимости она может быть перемещена или удалена.
            this.ulitsa_RUTableAdapter.Fill(this.abiturientDataSet9.Ulitsa_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet8.Nas_punkt_RU". При необходимости она может быть перемещена или удалена.
            this.nas_punkt_RUTableAdapter.Fill(this.abiturientDataSet8.Nas_punkt_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet7.Oblast_RU". При необходимости она может быть перемещена или удалена.
            this.oblast_RUTableAdapter.Fill(this.abiturientDataSet7.Oblast_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet6.School_RU". При необходимости она может быть перемещена или удалена.
            this.school_RUTableAdapter.Fill(this.abiturientDataSet6.School_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet5.National_RU". При необходимости она может быть перемещена или удалена.
            this.national_RUTableAdapter.Fill(this.abiturientDataSet5.National_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet3.Forma_obuchenia_RU". При необходимости она может быть перемещена или удалена.
            this.forma_obuchenia_RUTableAdapter.Fill(this.abiturientDataSet3.Forma_obuchenia_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet2.Abiturient_RU". При необходимости она может быть перемещена или удалена.
            // this.abiturient_RUTableAdapter1.Fill(this.abiturientDataSet2.Abiturient_RU);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "abiturientDataSet1.Abiturient_RU". При необходимости она может быть перемещена или удалена.
            //this.abiturient_RUTableAdapter.Fill(this.abiturientDataSet1.Abiturient_RU);

            Get_list10();
            SqlConnection conn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
            
            string query = "SELECT * FROM Login";
            using (SqlCommand command = new SqlCommand(query, conn))
            {
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);

                    // Теперь у вас есть DataSet с данными из базы данных.
                }
            }
        }

        void Get_list()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from National_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM National_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "National_RU");
                    dataGridView1.DataSource = ds.Tables["National_RU"];
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
                SqlDataAdapter da = new SqlDataAdapter("select * from School_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM School_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "School_RU");
                    dataGridView2.DataSource = ds.Tables["School_RU"];
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
                SqlDataAdapter da = new SqlDataAdapter("select * from Oblast_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Oblast_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Oblast_RU");
                    dataGridView3.DataSource = ds.Tables["Oblast_RU"];
                    comboBox1.DataSource = null; 
                    comboBox1.DataSource = ds.Tables["Oblast_RU"]; 
                    comboBox1.DisplayMember = "Область"; 
                    comboBox1.ValueMember = "#п\\п"; 
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
                SqlDataAdapter da = new SqlDataAdapter("select * from Nas_punkt_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Nas_punkt_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Nas_punkt_RU");
                    dataGridView4.DataSource = ds.Tables["Nas_punkt_RU"];
                    comboBox2.DataSource = null;
                    comboBox2.DataSource = ds.Tables["Nas_punkt_RU"];
                    comboBox2.DisplayMember = "Название_населенного_пункта"; 
                    comboBox2.ValueMember = "#п\\п"; 
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
                SqlDataAdapter da = new SqlDataAdapter("select * from Ulitsa_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Ulitsa_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Ulitsa_RU");
                    dataGridView5.DataSource = ds.Tables["Ulitsa_RU"];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        void Get_list6()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from Specialnost_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Specialnost_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Specialnost_RU");
                    dataGridView6.DataSource = ds.Tables["Specialnost_RU"];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        void Get_list8()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from Forma_Obuchenia_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Forma_Obuchenia_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Forma_Obuchenia_RU");
                    dataGridView8.DataSource = ds.Tables["Forma_Obuchenia_RU"];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        void Get_list9()
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                SqlDataAdapter da = new SqlDataAdapter("select * from Abiturient_RU", con);
                DataSet ds = new DataSet();
                con.Open();
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM Abiturient_RU", con);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable table = new DataTable();
                    adapter.Fill(ds, "Abiturient_RU");
                    dataGridView7.DataSource = ds.Tables["Abiturient_RU"];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            { MessageBox.Show("Заполните все поля!"); 
            
                
            }
            else
            {
               bool found = false;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox1.Text.Trim())
                    {
                        MessageBox.Show("Такая нацональность уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [National] (Name_Nation) VALUES (@Name_Nation)", con);
                    command.Parameters.AddWithValue("@Name_Nation", textBox1.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list();
                } 
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Заполните все поля!");


            }
            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox2.Text.Trim())
                    {
                        MessageBox.Show("Такая школа уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [School] (Name_School) VALUES (@Name_School)", con);
                    command.Parameters.AddWithValue("@Name_School", textBox2.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list2();
                }
            }
        }
       
        private void button6_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox3.Text))
            {

                MessageBox.Show("Заполните все поля!");
            }

            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox3.Text.Trim())
                    {
                        MessageBox.Show("Такая область уже есть");
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [Oblast] (Name_Oblast) VALUES (@Name_Oblast)", con);
                    command.Parameters.AddWithValue("@Name_Oblast", textBox3.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list3();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
           
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [Oblast] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView3[0, dataGridView3.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list3();
            }
            catch
            {
                
                    MessageBox.Show("Данная область связана с населённым пунктом \n Пожалуйста удалите связанные населенные пункты");
                
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
          
            
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [National] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list();
            }
            catch
            {
                
                MessageBox.Show("Данная национальность связана с абитуриентами \n Пожалуйста удалите связанные абитуриенты");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [School] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list2();
            }
            catch
            {
                MessageBox.Show("Данная школа связана с абитуриентами \n Пожалуйста удалите связанные абитуриенты");
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox5.Text))
           
                {
                    MessageBox.Show("Заполните все поля!");
                }
            else
                {
                    bool found = false;
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox5.Text.Trim())
                    {
                        MessageBox.Show("Такая специальность уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [Specialnost] (Name_Specialnost) VALUES (@Name_Specialnost)", con);
                    command.Parameters.AddWithValue("@Name_Specialnost", textBox5.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list6();
                    Get_list10();
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [Specialnost] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView6[0, dataGridView6.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list6();
                Get_list10();
            }
            catch
            {
                MessageBox.Show("Данная специальность связана с абитуриентами \n Пожалуйста удалите связанные абитуриенты");
            }
        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox6.Text))
           
                {
                    MessageBox.Show("Заполните все поля!");


                }
            else
                {
                    bool found = false;

                foreach (DataGridViewRow row in dataGridView8.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox6.Text.Trim())
                    {
                        MessageBox.Show("Такая форма обучения уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [Forma_Obuchenia] (Name_Forma_Obuchenia) VALUES (@Forma_Obuchenia)", con);
                    command.Parameters.AddWithValue("@Forma_Obuchenia", textBox6.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list8();
                }
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [Forma_Obuchenia] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView8[0, dataGridView8.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list8();
            }
            catch
            {
                MessageBox.Show("Данная форма обучения связана с абитуриентами \n Пожалуйста удалите связанные абитуриенты");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox7.Text))
            {
                    MessageBox.Show("Заполните все поля!");
            }

                else
                {
                    bool found = false;
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox4.Text.Trim() && row.Cells[3].Value.ToString() == comboBox1.Text.ToString())
                    {
                        MessageBox.Show("Такой населённый пункт уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [Nas_punkt] (Name_nas_punkt, Vid_nas_punkt, Id_oblast) VALUES (@Name_nas_punkt, @Vid_nas_punkt, @Id_oblast)", con);
                    command.Parameters.AddWithValue("@Name_nas_punkt", textBox4.Text.Trim());
                    command.Parameters.AddWithValue("@Vid_nas_punkt", textBox7.Text.Trim());
                    command.Parameters.AddWithValue("@Id_oblast", comboBox1.SelectedValue);
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list4();
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox8.Text))
            {

                MessageBox.Show("Заполните все поля!");
            }
            else
            {


                bool found = false;
                foreach (DataGridViewRow row in dataGridView5.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[2].Value.ToString() == textBox8.Text.Trim() && row.Cells[1].Value.ToString() == comboBox2.Text.ToString())
                    {
                        MessageBox.Show("Такая улица уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("INSERT INTO [Ulitsa] (Name_ulitsa, Id_Nas_punkt) VALUES (@Name_ulitsa, @Id_Nas_punkt)", con);
                    command.Parameters.AddWithValue("@Name_ulitsa", textBox8.Text.Trim());
                    command.Parameters.AddWithValue("@Id_Nas_punkt", comboBox2.SelectedValue);
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list5();
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
            con.Open();
            SqlCommand command = new SqlCommand("delete from [Ulitsa] where Id = @Id", con);
            command.Parameters.AddWithValue("@Id", dataGridView5[0, dataGridView5.CurrentRow.Index].Value.ToString());
            command.ExecuteNonQuery();
            con.Close();
            Get_list5();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [Nas_punkt] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView4[0, dataGridView4.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list4();
            }
            catch
            {
                MessageBox.Show("Данный населенный пункт связан с улицами \n Пожалуйста удалите связанные улицы");

            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            AddAbiturient addAbiturient = new AddAbiturient();
            addAbiturient.ShowDialog();
            if (DialogResult == DialogResult.OK || DialogResult == DialogResult.None)
            {
                Get_list9();
            }
        } 

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                con.Open();
                SqlCommand command = new SqlCommand("delete from [Abiturient] where Id = @Id", con);
                command.Parameters.AddWithValue("@Id", dataGridView7[0, dataGridView7.CurrentRow.Index].Value.ToString());
                command.ExecuteNonQuery();
                con.Close();
                Get_list9();
            }
            catch
            {
                MessageBox.Show("OK");

            }
        }

        private void ChangeButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Заполните все поля!");


            }
            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox1.Text.Trim())
                    {
                        MessageBox.Show("Такая нацональность уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("UPDATE [National] SET Name_Nation = @Name_Nation WHERE Id = @id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_Nation", textBox1.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Заполните все поля!");


            }
            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox2.Text.Trim())
                    {
                        MessageBox.Show("Такая школа уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("UPDATE [School] SET Name_School = @Name_School WHERE Id = @id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_School", textBox2.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list2();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Заполните все поля!");


            }
            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox3.Text.Trim())
                    {
                        MessageBox.Show("Такая область уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("UPDATE [Oblast] SET Name_Oblast = @Name_Oblast WHERE Id = @id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView3[0, dataGridView3.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_Oblast", textBox3.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list3();
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox6.Text))
            {
                MessageBox.Show("Заполните все поля!");


            }
            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView8.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox6.Text.Trim())
                    {
                        MessageBox.Show("Такая форма обучения уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("UPDATE [Forma_Obuchenia] SET Name_Forma_Obuchenia = @Name_Forma_Obuchenia WHERE Id = @id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView8[0, dataGridView8.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_Forma_Obuchenia", textBox6.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list8();
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox5.Text))
            {
                MessageBox.Show("Заполните все поля!");


            }
            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox5.Text.Trim())
                    {
                        MessageBox.Show("Такая специальность уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("UPDATE [Specialnost] SET Name_Specialnost = @Name_Specialnost WHERE Id = @id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView6[0, dataGridView6.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_Specialnost", textBox5.Text.Trim());
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list6();
                    Get_list10();
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox7.Text))
            {
                MessageBox.Show("Заполните все поля!");
            }

            else
            {
                bool found = false;
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    // Проверяем значение в конкретной ячейке (предположим, что оно находится в первой ячейке)
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == textBox4.Text.Trim() && row.Cells[3].Value.ToString() == comboBox1.Text.ToString())
                    {
                        MessageBox.Show("Такой населённый пункт уже есть");
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
                    con.Open();
                    SqlCommand command = new SqlCommand("UPDATE [Nas_punkt] SET Name_nas_punkt = @Name_nas_punkt, Vid_nas_punkt = @Vid_nas_punkt, Id_oblast = @Id_oblast WHERE Id=@id", con);
                    command.Parameters.AddWithValue("@Id", dataGridView4[0, dataGridView4.CurrentRow.Index].Value.ToString());
                    command.Parameters.AddWithValue("@Name_nas_punkt", textBox4.Text.Trim());
                    command.Parameters.AddWithValue("@Vid_nas_punkt", textBox7.Text.Trim());
                    command.Parameters.AddWithValue("@Id_oblast", comboBox1.SelectedValue);
                    command.ExecuteNonQuery();
                    con.Close();
                    Get_list4();
                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            int id_test = Convert.ToInt32(dataGridView7[0, dataGridView7.CurrentRow.Index].Value.ToString());
            EditAbiturient editAbiturient = new EditAbiturient(id_test);
            editAbiturient.ShowDialog();
            if (DialogResult == DialogResult.OK || DialogResult == DialogResult.None)
            {
                Get_list9();
            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button22_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            Microsoft.Office.Interop.Excel.Application exApp = new Excel.Application();
            exApp.Application.Workbooks.Add(Type.Missing);
            exApp.Application.Columns.ColumnWidth = 20;

            Excel.Range _excelCells = (Excel.Range)exApp.get_Range("A1", "Y1").Cells;
            _excelCells.Merge(Type.Missing);
            exApp.Cells[1, 1].Value = "Отчёт поданных заявлений с " + dateTimePicker1.Text + " по " + dateTimePicker2.Text + " на специальность: " + comboBox3.Text;
            exApp.Cells[1, 1].Font.Size = 14;
            exApp.Cells[1, 1].Font.Italic = true;
            exApp.Cells[1, 1].Font.Bold = true;
            exApp.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            exApp.Cells[3, 1] = "№п\\п";
            exApp.Cells[3, 1].Font.Bold = true;
            exApp.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            exApp.Cells[3, 2] = "ФИО абитуриента";
            exApp.Cells[3, 2].Font.Bold = true;
            exApp.Cells[3, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            exApp.Cells[3, 3] = "Название_Специальности";
            exApp.Cells[3, 3].Font.Bold = true;
            exApp.Cells[3, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 3].columnwidth = 30;

            exApp.Cells[3, 4] = "Форма_обучения";
            exApp.Cells[3, 4].Font.Bold = true;
            exApp.Cells[3, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            exApp.Cells[3, 5] = "Пол";
            exApp.Cells[3, 5].Font.Bold = true;
            exApp.Cells[3, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 5].columnwidth = 30;

            exApp.Cells[3, 6] = "День рождения";
            exApp.Cells[3, 6].Font.Bold = true;
            exApp.Cells[3, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 6].columnwidth = 30;

            exApp.Cells[3, 7] = "Место рождения";
            exApp.Cells[3, 7].Font.Bold = true;
            exApp.Cells[3, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 7].columnwidth = 30;

            exApp.Cells[3, 8] = "НАциональность";
            exApp.Cells[3, 8].Font.Bold = true;
            exApp.Cells[3, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 8].columnwidth = 30;

            exApp.Cells[3, 9] = "ИНН";
            exApp.Cells[3, 9].Font.Bold = true;
            exApp.Cells[3, 9].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 9].columnwidth = 30;

            exApp.Cells[3, 10] = "Страховой полис";
            exApp.Cells[3, 10].Font.Bold = true;
            exApp.Cells[3, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 10].columnwidth = 30;

            exApp.Cells[3, 11] = "Адрес";
            exApp.Cells[3, 11].Font.Bold = true;
            exApp.Cells[3, 11].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 11].columnwidth = 30;

            exApp.Cells[3, 12] = "Номер дома";
            exApp.Cells[3, 12].Font.Bold = true;
            exApp.Cells[3, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 12].columnwidth = 30;

            exApp.Cells[3, 13] = "Паспорт Серия";
            exApp.Cells[3, 13].Font.Bold = true;
            exApp.Cells[3, 13].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 13].columnwidth = 30;

            exApp.Cells[3, 14] = "Паспорт Номер";
            exApp.Cells[3, 14].Font.Bold = true;
            exApp.Cells[3, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 14].columnwidth = 30;

            exApp.Cells[3, 15] = "Паспорт Кем выдан";
            exApp.Cells[3, 15].Font.Bold = true;
            exApp.Cells[3, 15].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 15].columnwidth = 30;

            exApp.Cells[3, 16] = "Паспорт когда выдан";
            exApp.Cells[3, 16].Font.Bold = true;
            exApp.Cells[3, 16].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 16].columnwidth = 30;

            exApp.Cells[3, 17] = "Школа";
            exApp.Cells[3, 17].Font.Bold = true;
            exApp.Cells[3, 17].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 17].columnwidth = 30;

            exApp.Cells[3, 18] = "Льгота";
            exApp.Cells[3, 18].Font.Bold = true;
            exApp.Cells[3, 18].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 18].columnwidth = 30;

            exApp.Cells[3, 19] = "ФИО матери";
            exApp.Cells[3, 19].Font.Bold = true;
            exApp.Cells[3, 19].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 19].columnwidth = 30;

            exApp.Cells[3, 20] = "ФИО отца";
            exApp.Cells[3, 20].Font.Bold = true;
            exApp.Cells[3, 20].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 20].columnwidth = 30;
            
            exApp.Cells[3, 21] = "Средний балл аттестата";
            exApp.Cells[3, 21].Font.Bold = true;
            exApp.Cells[3, 21].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 21].columnwidth = 30;

            exApp.Cells[3, 22] = "Дата подачи заявления";
            exApp.Cells[3, 22].Font.Bold = true;
            exApp.Cells[3, 22].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 22].columnwidth = 30;

            exApp.Cells[3, 23] = "Номер заявления";
            exApp.Cells[3, 23].Font.Bold = true;
            exApp.Cells[3, 23].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 23].columnwidth = 30;

            exApp.Cells[3, 24] = "Экзамен1";
            exApp.Cells[3, 24].Font.Bold = true;
            exApp.Cells[3, 24].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 24].columnwidth = 30;

            exApp.Cells[3, 25] = "Экзамен2";
            exApp.Cells[3, 25].Font.Bold = true;
            exApp.Cells[3, 25].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            exApp.Cells[3, 25].columnwidth = 30;

            string sql = "SELECT * FROM Abiturient_RU WHERE [Дата_подачи_заявления] between @SelectedDate and @SelectedDate2 and [Название_Специальности] = @SelectSpecialnost";
            SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\\abiturient.mdf;Integrated Security=True;Connect Timeout=30");
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@SelectedDate", dateTimePicker1.Value.Date);
            cmd.Parameters.AddWithValue("@SelectedDate2", dateTimePicker2.Value.Date);
            cmd.Parameters.AddWithValue("@SelectSpecialnost", comboBox3.Text);
            SqlDataReader dr = cmd.ExecuteReader();
            int i = 5;
            while (dr.Read())
            {
                exApp.Cells[i, 1].Value = (String.Format("{0}", dr["#п\\п"]));
                exApp.Cells[i, 2].Value = (String.Format("{0}", dr["ФИО_абитуриента"]));
                exApp.Cells[i, 3].Value = (String.Format("{0}", dr["Название_Специальности"]));
                exApp.Cells[i, 4].Value = (String.Format("{0}", dr["Форма_Обучения"]));
                exApp.Cells[i, 5].Value = (String.Format("{0}", dr["Пол"]));

                exApp.Cells[i, 6].Value = (String.Format("{0}", dr["День_рождения"]));
                exApp.Cells[i, 7].Value = (String.Format("{0}", dr["Место_рождения"]));
                exApp.Cells[i, 8].Value = (String.Format("{0}", dr["Национальность"]));
                exApp.Cells[i, 9].Value = (String.Format("{0}", dr["ИНН"]));
                exApp.Cells[i, 10].Value = (String.Format("{0}", dr["Страховой_полис"]));

                exApp.Cells[i, 11].Value = (String.Format("{0}", dr["Адрес"]));
                exApp.Cells[i, 12].Value = (String.Format("{0}", dr["Номер_дома"]));
                exApp.Cells[i, 13].Value = (String.Format("{0}", dr["Паспорт_Серия"]));
                exApp.Cells[i, 14].Value = (String.Format("{0}", dr["Паспорт_Номер"]));
                exApp.Cells[i, 15].Value = (String.Format("{0}", dr["Паспорт_Кем_выдан"]));

                exApp.Cells[i, 16].Value = (String.Format("{0}", dr["Паспорт_Когда_выдан"]));
                exApp.Cells[i, 17].Value = (String.Format("{0}", dr["Школа"]));
                exApp.Cells[i, 18].Value = (String.Format("{0}", dr["Льгота"]));
                exApp.Cells[i, 19].Value = (String.Format("{0}", dr["ФИО_Матери"]));
                exApp.Cells[i, 20].Value = (String.Format("{0}", dr["ФИО_Отца"]));

                exApp.Cells[i, 21].Value = (String.Format("{0}", dr["Средний_балл_аттестата"]));
            
                exApp.Cells[i, 22].Value = (String.Format("{0}", dr["Дата_подачи_заявления"]));
                exApp.Cells[i, 23].Value = (String.Format("{0}", dr["Номер_заявления"]));
                exApp.Cells[i, 24].Value = (String.Format("{0}", dr["Экзамен1"]));
                exApp.Cells[i, 25].Value = (String.Format("{0}", dr["Экзамен2"]));

                i++;
            }
            dr.Close();
            con.Close();

            exApp.Visible = true;
        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void tabPage9_Click(object sender, EventArgs e)
        {

        }
    }
}
