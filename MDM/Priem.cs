using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDM
{
    public partial class Priem : Form
    {

        private SqlConnection sqlConnection = null;

        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;

        private SqlDataAdapter sqlDataAdapter2 = null;
        private DataSet dataSet2 = null;


        Bitmap MemoryImage;
        private PrintDocument printDocument1 = new PrintDocument();
        private PrintPreviewDialog previewdlg = new PrintPreviewDialog();

        private PrintDocument printDocument2 = new PrintDocument();
        private PrintPreviewDialog previewdlg2 = new PrintPreviewDialog();
        public Priem()
        {
            InitializeComponent();
            printDocument1.PrintPage += new PrintPageEventHandler(printdoc1_PrintPage);
            printDocument1.PrintPage += new PrintPageEventHandler(printdoc2_PrintPage);

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

        private void Priem_Load(object sender, EventArgs e)
        {
            LoadData();
            LoadData2();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            label1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            label2.Text = Convert.ToString( DateTime.Now) ;
            label3.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            label4.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            label5.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            label6.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            ////увольнение
            ///
            label7.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            label9.Text = Convert.ToString(DateTime.Now);
            label13.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            label10.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            label12.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            label11.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();

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

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            label8.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Print(this.panel1);
        }

        public void GetPrintArea(Panel pnl)
        {
            MemoryImage = new Bitmap(pnl.Width, pnl.Height);
            pnl.DrawToBitmap(MemoryImage, new Rectangle(0, 0, pnl.Width, pnl.Height));
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (MemoryImage != null)
            {
                e.Graphics.DrawImage(MemoryImage, 0, 0);
                base.OnPaint(e);
            }
        }
        void printdoc1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Rectangle pagearea = e.PageBounds;
            e.Graphics.DrawImage(MemoryImage, (pagearea.Width / 2) - (this.panel1.Width / 2), this.panel1.Location.Y);
        }

        void printdoc2_PrintPage(object sender, PrintPageEventArgs e)
        {
            Rectangle pagearea = e.PageBounds;
            e.Graphics.DrawImage(MemoryImage, (pagearea.Width / 2) - (this.panel2.Width / 2), this.panel2.Location.Y);
        }
        public void Print(Panel pnl)
        {
            Panel pannel = pnl;
            GetPrintArea(pnl);
            previewdlg.Document = printDocument1;
            previewdlg.ShowDialog();
        }

        public void Print2(Panel pnl)
        {
            Panel pannel = pnl;
            GetPrintArea(pnl);
            previewdlg2.Document = printDocument1;
            previewdlg2.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Print2(this.panel2);
        }
    }

}
