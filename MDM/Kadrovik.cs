using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace MDM
{
    public partial class Kadrovik : Form
    {
        public Kadrovik()
        {
            InitializeComponent();

            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            //SqlDataReader dr;
            SqlDataReader dr2;

            //SqlCommand cmd = new SqlCommand();
            //cmd.Connection = sqlConnection;

            //cmd.CommandText = "SELECT * FROM Employe";
            //dr = cmd.ExecuteReader();

            //while (dr.Read())
            //{
            //    comboBox1.Items.Add(dr["PersonalNum"]);
            //}
            //dr.Close();

            SqlCommand cmd2 = new SqlCommand();
            cmd2.Connection = sqlConnection;

            cmd2.CommandText = "SELECT * FROM Department";
            dr2 = cmd2.ExecuteReader();

            while (dr2.Read())
            {
                comboBox3.Items.Add(dr2["Department"]);
            }
            dr2.Close();

            comboBox2.Items.Add("Работает");
            comboBox2.Items.Add("Уволен");

            comboBox4.Items.Add("Мужской");
            comboBox4.Items.Add("Женский");
        }


        private SqlConnection sqlConnection = null;
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;
        private bool newRowAdding = false;

        private SqlCommandBuilder sqlBuilder2 = null;
        private SqlDataAdapter sqlDataAdapter2 = null;
        private DataSet dataSet2 = null;
        private bool newRowAdding2 = false;

        private void LoadData()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT *, 'Delete' AS [Command] FROM Employe", sqlConnection);

                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);

                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();

                dataSet = new DataSet();

                sqlDataAdapter.Fill(dataSet, "Employe");

                dataGridView1.DataSource = dataSet.Tables["Employe"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[20, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Kadrovik_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            LoadData();
            LoadData2();
        }


        private void ReloadData()
        {
            try
            {
                dataSet.Tables["Employe"].Clear();

                sqlDataAdapter.Fill(dataSet, "Employe");

                dataGridView1.DataSource = dataSet.Tables["Employe"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[20, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 20)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[20].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView1.Rows.RemoveAt(rowIndex);

                            dataSet.Tables["Employe"].Rows[rowIndex].Delete();

                            sqlDataAdapter.Update(dataSet, "Employe");
                        }
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet.Tables["Employe"].Rows[r]["PersonalNum"] = dataGridView1.Rows[r].Cells["PersonalNum"].Value;
                        dataSet.Tables["Employe"].Rows[r]["FIO"] = dataGridView1.Rows[r].Cells["FIO"].Value;
                        dataSet.Tables["Employe"].Rows[r]["INN"] = dataGridView1.Rows[r].Cells["INN"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Pasport"] = dataGridView1.Rows[r].Cells["Pasport"].Value;
                        dataSet.Tables["Employe"].Rows[r]["InnsuedBy"] = dataGridView1.Rows[r].Cells["InnsuedBy"].Value;
                        dataSet.Tables["Employe"].Rows[r]["DataOfInssued"] = dataGridView1.Rows[r].Cells["DataOfInssued"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Snils"] = dataGridView1.Rows[r].Cells["Snils"].Value;
                        dataSet.Tables["Employe"].Rows[r]["PhoneNum"] = dataGridView1.Rows[r].Cells["PhoneNum"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Mail"] = dataGridView1.Rows[r].Cells["Mail"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Photo"] = dataGridView1.Rows[r].Cells["Photo"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Address"] = dataGridView1.Rows[r].Cells["Address"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Position"] = dataGridView1.Rows[r].Cells["Position"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Department"] = dataGridView1.Rows[r].Cells["Department"].Value;
                        dataSet.Tables["Employe"].Rows[r]["BithDay"] = dataGridView1.Rows[r].Cells["BithDay"].Value;
                        dataSet.Tables["Employe"].Rows[r]["EmployementData"] = dataGridView1.Rows[r].Cells["EmployementData"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Gender"] = dataGridView1.Rows[r].Cells["Gender"].Value;
                        dataSet.Tables["Employe"].Rows[r]["MaritalStatus"] = dataGridView1.Rows[r].Cells["MaritalStatus"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Education"] = dataGridView1.Rows[r].Cells["Education"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Status"] = dataGridView1.Rows[r].Cells["Status"].Value;
                        dataSet.Tables["Employe"].Rows[r]["Expirience"] = dataGridView1.Rows[r].Cells["Expirience"].Value;
                       
                        sqlDataAdapter.Update(dataSet, "Employe");

                        dataGridView1.Rows[e.RowIndex].Cells[20].Value = "Delete";
                    }

                    ReloadData();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dataGridView1.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[20, rowIndex] = linkCell;

                    editingRow.Cells["Command"].Value = "Update";

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /////отпуска
        ///
        private void LoadData2()
        {
            try
            {
                sqlDataAdapter2 = new SqlDataAdapter("SELECT *, 'Delete' AS [Command] FROM Otpysk", sqlConnection);

                sqlBuilder2 = new SqlCommandBuilder(sqlDataAdapter2);

                sqlBuilder2.GetInsertCommand();
                sqlBuilder2.GetUpdateCommand();
                sqlBuilder2.GetDeleteCommand();

                dataSet2 = new DataSet();

                sqlDataAdapter2.Fill(dataSet2, "Otpysk");

                dataGridView2.DataSource = dataSet2.Tables["Otpysk"];

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[4, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReloadData2()
        {
            try
            {
                dataSet2.Tables["Otpysk"].Clear();

                sqlDataAdapter2.Fill(dataSet2, "Otpysk");

                dataGridView2.DataSource = dataSet2.Tables["Otpysk"];

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[4, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding2 == false)
                {
                    int rowIndex = dataGridView2.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dataGridView2.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[4, rowIndex] = linkCell;

                    editingRow.Cells["Command"].Value = "Update";

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 4)
                {
                    string task = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView2.Rows.RemoveAt(rowIndex);

                            dataSet2.Tables["Otpysk"].Rows[rowIndex].Delete();

                            sqlDataAdapter2.Update(dataSet2, "Otpysk");
                        }
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet2.Tables["Otpysk"].Rows[r]["ID_Otpysk"] = dataGridView1.Rows[r].Cells["ID_Otpysk"].Value;
                        dataSet2.Tables["Otpysk"].Rows[r]["PersonalNum"] = dataGridView1.Rows[r].Cells["PersonalNum"].Value;
                        dataSet2.Tables["Otpysk"].Rows[r]["DataStart"] = dataGridView1.Rows[r].Cells["DataStart"].Value;
                        dataSet2.Tables["Otpysk"].Rows[r]["DataFinish"] = dataGridView1.Rows[r].Cells["DataFinish"].Value;
                       
                        sqlDataAdapter2.Update(dataSet2, "Otpysk");

                        dataGridView2.Rows[e.RowIndex].Cells[4].Value = "Delete";
                    }

                    ReloadData2();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            NewEmploye f1 = new NewEmploye();
            f1.ShowDialog();

        }

        private void button4_Click(object sender, EventArgs e)
        {
           LoadData();
        }

        private void Kadrovik_Activated(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void btClear1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            comboBox1.Text = "";

        }

        private void btAdd1_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            SqlCommand command = new SqlCommand($"INSERT INTO [Otpysk] (ID_Otpysk, PersonalNum, DataStart, DataFinish) VALUES (@id, @name, @start, @finish)", sqlConnection);

            if (comboBox1.Text.Length < 1 || textBox1.Text.Length < 1 )
            {
                MessageBox.Show("Зполните поля!");
            }
            else
            {
                command.Parameters.AddWithValue("id", textBox1.Text);
                command.Parameters.AddWithValue("name", comboBox1.SelectedItem);
                command.Parameters.AddWithValue("start", dateTimePicker1.Value);
                command.Parameters.AddWithValue("finish", dateTimePicker2.Value);
                command.ExecuteNonQuery();
                MessageBox.Show("Информация добавлена");
                ReloadData2();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataSet.Tables["Employe"].Clear();

            SqlCommand cmd = new SqlCommand(@"SELECT * FROM Employe WHERE Status = @stat", sqlConnection);

            cmd.Parameters.AddWithValue("@stat", comboBox2.SelectedItem);

            sqlDataAdapter = new SqlDataAdapter(cmd);

            dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet, "Employe");
            dataGridView1.DataSource = dataSet.Tables["Employe"];
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataSet.Tables["Employe"].Clear();

            SqlCommand cmd = new SqlCommand(@"SELECT * FROM Employe WHERE Department = @dep", sqlConnection);

            cmd.Parameters.AddWithValue("@dep", comboBox3.SelectedItem);

            sqlDataAdapter = new SqlDataAdapter(cmd);

            dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet, "Employe");
            dataGridView1.DataSource = dataSet.Tables["Employe"];
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataSet.Tables["Employe"].Clear();

            SqlCommand cmd = new SqlCommand(@"SELECT * FROM Employe WHERE Gender = @gen", sqlConnection);

            cmd.Parameters.AddWithValue("@gen", comboBox4.SelectedItem);

            sqlDataAdapter = new SqlDataAdapter(cmd);

            dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet, "Employe");
            dataGridView1.DataSource = dataSet.Tables["Employe"];
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Spravka sp = new Spravka();
            sp.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Priem pr = new Priem();
            pr.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "PDF (*.pdf)|*.pdf";
                sfd.FileName = "Отпуска.pdf";
                bool fileError = false;


                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            PdfPTable pdfTable = new PdfPTable(dataGridView2.Columns.Count);

                            BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1250, BaseFont.EMBEDDED);

                            pdfTable.DefaultCell.Padding = 3;
                            pdfTable.WidthPercentage = 100;
                            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                            pdfTable.SpacingAfter = 20;
                            pdfTable.SpacingBefore = 20;


                            string FONT_LOCATION = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "times.ttf");
                            BaseFont baseFont = BaseFont.CreateFont(FONT_LOCATION, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font text = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.NORMAL);

                            iTextSharp.text.Font text1 = new iTextSharp.text.Font(baseFont, 20, iTextSharp.text.Font.NORMAL);

                            foreach (DataGridViewColumn column in dataGridView2.Columns)
                            {
                                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText, text));
                                pdfTable.AddCell(cell);
                            }

                            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                            {
                                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                                {
                                    pdfTable.AddCell(new Phrase(dataGridView2.Rows[i].Cells[j].Value.ToString(), text));
                                }
                            }


                            using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                            {
                                Document pdfDoc = new Document(PageSize.A4, 10f, 20f, 20f, 10f);



                                PdfWriter.GetInstance(pdfDoc, stream);
                                pdfDoc.Open();

                                Paragraph para1 = new Paragraph("Отпуска", text1);
                                para1.Alignment = Element.ALIGN_CENTER;
                                pdfDoc.Add(para1);
                                pdfDoc.Add(pdfTable);

                                pdfDoc.Close();
                                stream.Close();
                            }

                            MessageBox.Show("Отчет успешно создан", "Info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error :" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            SqlDataReader dr;

            comboBox1.Items.Clear();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = sqlConnection;

            cmd.CommandText = "SELECT * FROM Employe";
            dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                comboBox1.Items.Add(dr["PersonalNum"]);
            }
            dr.Close();
        }
    }
}
