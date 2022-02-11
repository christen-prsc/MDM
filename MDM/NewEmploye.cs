using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDM
{
    public partial class NewEmploye : Form
    {

        private SqlConnection sqlConnection = null;
        public NewEmploye()
        {
            InitializeComponent();

            comboBox1.Items.Add ("Женский");
            comboBox1.Items.Add("Мужской");

            comboBox2.Items.Add("В браке");
            comboBox2.Items.Add("Не в браке");

            comboBox5.Items.Add("Работает");
            comboBox5.Items.Add("Уволен");

            comboBox6.Items.Add("Нет");
            comboBox6.Items.Add("Есть");


            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");
            sqlConnection.Open();
            SqlDataReader dr;
            SqlDataReader dr2;

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = sqlConnection;

            cmd.CommandText = "SELECT * FROM PositionAtWork";
            dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                comboBox3.Items.Add(dr["Position"]);
            }
            dr.Close();

            SqlCommand cmd2 = new SqlCommand();
            cmd2.Connection = sqlConnection;

            cmd2.CommandText = "SELECT * FROM Department";
            dr2 = cmd2.ExecuteReader();

            while (dr2.Read())
            {
                comboBox4.Items.Add(dr2["Department"]);
            }
            dr2.Close();


        }

        

        private void button4_Click(object sender, EventArgs e)
        {
            maskedTextBox2.Clear();
            textBox2.Clear();
            maskedTextBox3.Clear();
            maskedTextBox4.Clear();
            maskedTextBox5.Clear();
            maskedTextBox1.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox5.Text = "";
            comboBox6.Text = "";


        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Image Files(*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG| All files (*.*)|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    pictureBox1.Image = new Bitmap(ofd.FileName);
                }
                catch 
                {
                    MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = Properties.Resources.img1;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            
            int year = DateTime.Now.Year;

            SqlCommand command = new SqlCommand($"INSERT INTO [Employe] (PersonalNum, FIO, INN, Pasport,InnsuedBy,DataOfInssued,Snils,PhoneNum,Mail,Address,Position,Department,BithDay,EmployementData,Gender,MaritalStatus,Education,Status,Expirience) VALUES (@id, @name, @inn, @pas, @insby, @datains, @snils, @phone, @mail, @address, @pos, @depar, @bday, @empldata, @gender, @marit, @educ, @stat, @exp)", sqlConnection);

            if (maskedTextBox1.Text.Length < 1 || maskedTextBox2.Text.Length < 1 || textBox2.Text.Length < 1 || maskedTextBox4.Text.Length < 1 || textBox6.Text.Length < 1 || comboBox3.Text.Length < 1 || comboBox4.Text.Length < 1 || comboBox1.Text.Length < 1 || comboBox2.Text.Length < 1 || comboBox5.Text.Length < 1 || comboBox6.Text.Length < 1)
            {
                MessageBox.Show("Заполните поля!");
                
            } else if (year - dateTimePicker1.Value.Year < 16)
            {
                MessageBox.Show("Возраст меньше 16 лет");
            }
            else
            {
                command.Parameters.AddWithValue("id", Convert.ToString(maskedTextBox2.Text));
                command.Parameters.AddWithValue("name", textBox2.Text);
                command.Parameters.AddWithValue("inn", Convert.ToString(maskedTextBox3.Text));
                command.Parameters.AddWithValue("pas", Convert.ToString(maskedTextBox4.Text));
                command.Parameters.AddWithValue("insby", textBox9.Text);
                command.Parameters.AddWithValue("datains", dateTimePicker2.Value);
                command.Parameters.AddWithValue("snils", Convert.ToString(maskedTextBox5.Text));
                command.Parameters.AddWithValue("phone", Convert.ToString(maskedTextBox1.Text));
                command.Parameters.AddWithValue("mail", textBox6.Text);
                //command.Parameters.AddWithValue("photo",pictureBox1 );
                command.Parameters.AddWithValue("address", textBox7.Text);
                command.Parameters.AddWithValue("pos", Convert.ToString(comboBox3.SelectedItem));
                command.Parameters.AddWithValue("depar", Convert.ToString(comboBox4.SelectedItem));
                command.Parameters.AddWithValue("bday", dateTimePicker1.Value);
                command.Parameters.AddWithValue("empldata", dateTimePicker3.Value);
                command.Parameters.AddWithValue("gender", Convert.ToString(comboBox1.SelectedItem));
                command.Parameters.AddWithValue("marit", Convert.ToString(comboBox2.SelectedItem));
                command.Parameters.AddWithValue("educ", textBox8.Text);
                command.Parameters.AddWithValue("stat", Convert.ToString(comboBox5.SelectedItem));
                command.Parameters.AddWithValue("exp", Convert.ToString(comboBox6.SelectedItem));
                command.ExecuteNonQuery();
                MessageBox.Show("Сотрудник добавлен");

                this.Close();
                
            }
        }

        private void maskedTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void maskedTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            string sym = e.KeyChar.ToString();
            char number = e.KeyChar;

            if ((!Regex.Match(sym, @"[а-яА-Я]|[a-zA-Z]").Success) && (number != 8))
            {
                e.Handled = true;
            }
        }

        private void maskedTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }

        }

        private void maskedTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            string sym = e.KeyChar.ToString();
            char number = e.KeyChar;

            if ((!Regex.Match(sym, @"[а-яА-Я]|[a-zA-Z]").Success) && (number != 8))
            {
                e.Handled = true;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            string sym = e.KeyChar.ToString();
            char number = e.KeyChar;

            if ((!Regex.Match(sym, @"[а-яА-Я]|[a-zA-Z]").Success) && (number != 8))
            {
                e.Handled = true;
            }
        }
    }
}
