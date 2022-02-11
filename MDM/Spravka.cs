using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDM
{
    public partial class Spravka : Form
    {


        private SqlConnection sqlConnection = null;
        
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;

        private SqlDataAdapter sqlDataAdapter2 = null;
        private DataSet dataSet2 = null;
        public Spravka()
        {
            InitializeComponent();
        }


        private void LoadData()
        {
            try
            {
                sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

                sqlConnection.Open();

                SqlCommand cmd = new SqlCommand(@"SELECT * FROM Employe ", sqlConnection);

                sqlDataAdapter = new SqlDataAdapter(cmd);

                dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet, "Employe");
                dataGridView1.DataSource = dataSet.Tables["Employe"];

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadData2()
        {
            try
            {
                sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

                sqlConnection.Open();

                SqlCommand cmd = new SqlCommand(@"SELECT * FROM PositionAtWork ", sqlConnection);

                sqlDataAdapter2 = new SqlDataAdapter(cmd);

                dataSet2 = new DataSet();
                sqlDataAdapter2.Fill(dataSet2, "PositionAtWork");
                dataGridView2.DataSource = dataSet2.Tables["PositionAtWork"];

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Spravka_Load(object sender, EventArgs e)
        {
            LoadData();
            LoadData2();

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            label2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            label3.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            label4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            label5.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            label1.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();


        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            label6.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int width = panel1.Width;
            int height = panel1.Height;

            Bitmap bmp = new Bitmap(width, height);
            panel1.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, width, height));

            bmp.Save(@"C:\Users\Кристина\OneDrive\Рабочий стол\C#\MDM\Справка.png", ImageFormat.Png);

            MessageBox.Show("Справка сохранена");
        }
    }
}
