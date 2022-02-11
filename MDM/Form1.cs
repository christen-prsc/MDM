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

namespace MDM
{
    public partial class Form1 : Form
    {

        private SqlConnection sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string loginUser = textBox1.Text;
            string pasUser = textBox2.Text;
           

            string query = "SELECT Login FROM Passwords WHERE Login = @ul AND Password = @uP ";
            string returnValue = "";

            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");
            sqlConnection.Open();




            using (SqlCommand auto = new SqlCommand(query, sqlConnection))
            {
                auto.Parameters.Add("uL", SqlDbType.VarChar).Value = loginUser;
                auto.Parameters.Add("uP", SqlDbType.VarChar).Value = pasUser;
                returnValue = (string)auto.ExecuteScalar();
            }


            if (String.IsNullOrEmpty(returnValue))
            {
                MessageBox.Show("Заполните поля корректными данными");
                return;
            }
            returnValue = returnValue.Trim();

            if (returnValue == "expert")
            {
                Admin f1 = new Admin();
                f1.ShowDialog();

            }
            else if (returnValue == "cadr")
            {
                Kadrovik f2 = new Kadrovik();
                f2.ShowDialog();
            }
          
            textBox1.Clear();
            textBox2.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
