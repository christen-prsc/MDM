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
    public partial class Admin : Form
    {

        private SqlConnection sqlConnection = null;
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;
        private bool newRowAdding = false;

        private SqlCommandBuilder sqlBuilder2 = null;
        private SqlDataAdapter sqlDataAdapter2 = null;
        private DataSet dataSet2 = null;
        private bool newRowAdding2 = false;

        private SqlCommandBuilder sqlBuilder3 = null;
        private SqlDataAdapter sqlDataAdapter3 = null;
        private DataSet dataSet3 = null;
        private bool newRowAdding3 = false;


        public Admin()
        {
            InitializeComponent();

            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            SqlDataReader dr;
           

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


        private void LoadData()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT *, 'Delete' AS [Command] FROM Passwords", sqlConnection);

                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);

                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();

                dataSet = new DataSet();

                sqlDataAdapter.Fill(dataSet, "Passwords");

                dataGridView1.DataSource = dataSet.Tables["Passwords"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Admin_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            LoadData();
            LoadData2();
            LoadData3();
        }


        private void ReloadData()
        {
            try
            {
                dataSet.Tables["Passwords"].Clear();

                sqlDataAdapter.Fill(dataSet, "Passwords");

                dataGridView1.DataSource = dataSet.Tables["Passwords"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, i] = linkCell;
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
                if (e.ColumnIndex == 4)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView1.Rows.RemoveAt(rowIndex);

                            dataSet.Tables["Passwords"].Rows[rowIndex].Delete();

                            sqlDataAdapter.Update(dataSet, "Passwords");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView1.Rows.Count - 2;

                        DataRow row = dataSet.Tables["Passwords"].NewRow();

                        row["ID_Password"] = dataGridView1.Rows[rowIndex].Cells["ID_Password"].Value;
                        row["PersonalNum"] = dataGridView1.Rows[rowIndex].Cells["PersonalNum"].Value;
                        row["Login"] = dataGridView1.Rows[rowIndex].Cells["Login"].Value;
                        row["Password"] = dataGridView1.Rows[rowIndex].Cells["Password"].Value;
                       

                        dataSet.Tables["Passwords"].Rows.Add(row);

                        dataSet.Tables["Passwords"].Rows.RemoveAt(dataSet.Tables["Passwords"].Rows.Count - 1);

                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);

                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "Delete";

                        sqlDataAdapter.Update(dataSet, "Passwords");

                        newRowAdding = false;
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet.Tables["Passwords"].Rows[r]["ID_Password"] = dataGridView1.Rows[r].Cells["ID_Password"].Value;
                        dataSet.Tables["Passwords"].Rows[r]["PersonalNum"] = dataGridView1.Rows[r].Cells["PersonalNum"].Value;
                        dataSet.Tables["Passwords"].Rows[r]["Login"] = dataGridView1.Rows[r].Cells["Login"].Value;
                        dataSet.Tables["Passwords"].Rows[r]["Password"] = dataGridView1.Rows[r].Cells["Password"].Value;
                       

                        sqlDataAdapter.Update(dataSet, "Passwords");

                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "Delete";
                    }

                    ReloadData();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    newRowAdding = true;

                    int lastRow = dataGridView1.Rows.Count - 2;

                    DataGridViewRow row = dataGridView1.Rows[lastRow];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, lastRow] = linkCell;

                    row.Cells["Command"].Value = "Insert";
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

                    dataGridView1[4, rowIndex] = linkCell;

                    editingRow.Cells["Command"].Value = "Update";

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        //////////////отделы
        ///
        private void LoadData2()
        {
            try
            {
                sqlDataAdapter2 = new SqlDataAdapter("SELECT *, 'Delete' AS [Command] FROM Department", sqlConnection);

                sqlBuilder2 = new SqlCommandBuilder(sqlDataAdapter2);

                sqlBuilder2.GetInsertCommand();
                sqlBuilder2.GetUpdateCommand();
                sqlBuilder2.GetDeleteCommand();

                dataSet2 = new DataSet();

                sqlDataAdapter2.Fill(dataSet2, "Department");

                dataGridView2.DataSource = dataSet2.Tables["Department"];

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[3, i] = linkCell;
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
                dataSet2.Tables["Department"].Clear();

                sqlDataAdapter2.Fill(dataSet2, "Department");

                dataGridView2.DataSource = dataSet2.Tables["Department"];

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[3, i] = linkCell;
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
                if (e.ColumnIndex == 3)
                {
                    string task = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView2.Rows.RemoveAt(rowIndex);

                            dataSet2.Tables["Department"].Rows[rowIndex].Delete();

                            sqlDataAdapter2.Update(dataSet2, "Department");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView2.Rows.Count - 2;

                        DataRow row = dataSet2.Tables["Department"].NewRow();

                        row["Department"] = dataGridView2.Rows[rowIndex].Cells["Department"].Value;
                        row["NumCabinet"] = dataGridView2.Rows[rowIndex].Cells["NumCabinet"].Value;
                        row["PhoneNum"] = dataGridView2.Rows[rowIndex].Cells["PhoneNum"].Value;
                       


                        dataSet2.Tables["Department"].Rows.Add(row);

                        dataSet2.Tables["Department"].Rows.RemoveAt(dataSet2.Tables["Department"].Rows.Count - 1);

                        dataGridView2.Rows.RemoveAt(dataGridView2.Rows.Count - 2);

                        dataGridView2.Rows[e.RowIndex].Cells[3].Value = "Delete";

                        sqlDataAdapter2.Update(dataSet2, "Department");

                        newRowAdding2 = false;
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet2.Tables["Department"].Rows[r]["Department"] = dataGridView2.Rows[r].Cells["Department"].Value;
                        dataSet2.Tables["Department"].Rows[r]["NumCabinet"] = dataGridView2.Rows[r].Cells["NumCabinet"].Value;
                        dataSet2.Tables["Department"].Rows[r]["PhoneNum"] = dataGridView2.Rows[r].Cells["PhoneNum"].Value;

                        sqlDataAdapter2.Update(dataSet2, "Department");

                        dataGridView2.Rows[e.RowIndex].Cells[3].Value = "Delete";
                    }

                    ReloadData2();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridView2_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding2 == false)
                {
                    newRowAdding2 = true;

                    int lastRow = dataGridView2.Rows.Count - 2;

                    DataGridViewRow row = dataGridView2.Rows[lastRow];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[3, lastRow] = linkCell;

                    row.Cells["Command"].Value = "Insert";
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

                    DataGridViewRow editingRow2 = dataGridView2.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView2[3, rowIndex] = linkCell;

                    editingRow2.Cells["Command"].Value = "Update";

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        //////должности
        ///
        private void LoadData3()
        {
            try
            {
                sqlDataAdapter3 = new SqlDataAdapter("SELECT *, 'Delete' AS [Command] FROM PositionAtWork", sqlConnection);

                sqlBuilder3 = new SqlCommandBuilder(sqlDataAdapter3);

                sqlBuilder3.GetInsertCommand();
                sqlBuilder3.GetUpdateCommand();
                sqlBuilder3.GetDeleteCommand();

                dataSet3 = new DataSet();

                sqlDataAdapter3.Fill(dataSet3, "PositionAtWork");

                dataGridView3.DataSource = dataSet3.Tables["PositionAtWork"];

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView3[2, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReloadData3()
        {
            try
            {
                dataSet3.Tables["PositionAtWork"].Clear();

                sqlDataAdapter3.Fill(dataSet3, "PositionAtWork");

                dataGridView3.DataSource = dataSet3.Tables["PositionAtWork"];

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView3[2, i] = linkCell;
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 2)
                {
                    string task = dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView3.Rows.RemoveAt(rowIndex);

                            dataSet3.Tables["PositionAtWork"].Rows[rowIndex].Delete();

                            sqlDataAdapter3.Update(dataSet3, "PositionAtWork");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView3.Rows.Count - 2;

                        DataRow row = dataSet3.Tables["PositionAtWork"].NewRow();

                        row["Position"] = dataGridView3.Rows[rowIndex].Cells["Position"].Value;
                        row["Salary"] = dataGridView3.Rows[rowIndex].Cells["Salary"].Value;

                        dataSet3.Tables["PositionAtWork"].Rows.Add(row);

                        dataSet3.Tables["PositionAtWork"].Rows.RemoveAt(dataSet3.Tables["PositionAtWork"].Rows.Count - 1);

                        dataGridView3.Rows.RemoveAt(dataGridView3.Rows.Count - 2);

                        dataGridView3.Rows[e.RowIndex].Cells[2].Value = "Delete";

                        sqlDataAdapter3.Update(dataSet3, "PositionAtWork");

                        newRowAdding3 = false;
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet3.Tables["PositionAtWork"].Rows[r]["Position"] = dataGridView3.Rows[r].Cells["Position"].Value;
                        dataSet3.Tables["PositionAtWork"].Rows[r]["Salary"] = dataGridView3.Rows[r].Cells["Salary"].Value;

                        sqlDataAdapter3.Update(dataSet3, "PositionAtWork");

                        dataGridView3.Rows[e.RowIndex].Cells[2].Value = "Delete";
                    }

                    ReloadData3();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

      

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding3 == false)
                {
                    int rowIndex = dataGridView3.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow3 = dataGridView3.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView3[2, rowIndex] = linkCell;

                    editingRow3.Cells["Command"].Value = "Update";

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btAdd1_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            SqlCommand command = new SqlCommand($"INSERT INTO [Passwords] (ID_Password, PersonalNum, Login, Password) VALUES (@id, @name, @login, @pas)", sqlConnection);

           if (comboBox1.Text.Length < 1 || textBox2.Text.Length < 1 || textBox3.Text.Length < 1 || textBox1.Text.Length < 1)
            {
                MessageBox.Show("Зполните поля!");
            }
            else
            {
                command.Parameters.AddWithValue("id", textBox1.Text);
                command.Parameters.AddWithValue("name", comboBox1.SelectedItem);
                command.Parameters.AddWithValue("login", textBox2.Text);
                command.Parameters.AddWithValue("pas", textBox3.Text);
                command.ExecuteNonQuery();
                MessageBox.Show("Информация добавлена");
                ReloadData();
            }
        }

        private void btClear1_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            textBox2.Clear();
            textBox3.Clear();
            textBox1.Clear();
        }

        private void btClear2_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
            textBox6.Clear();
            maskedTextBox1.Clear();
        }

        private void btAdd2_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            SqlCommand command = new SqlCommand($"INSERT INTO [Department] (Department, NumCabinet, PhoneNum) VALUES (@name, @number, @phone)", sqlConnection);

            if (textBox4.Text.Length < 1 || textBox6.Text.Length < 1 || maskedTextBox1.Text.Length < 1 )
            {
                MessageBox.Show("Зполните поля!");
            }
            else
            {
                command.Parameters.AddWithValue("name", textBox4.Text);
                command.Parameters.AddWithValue("number", textBox6.Text);
                command.Parameters.AddWithValue("phone", Convert.ToString(maskedTextBox1.Text));
                command.ExecuteNonQuery();
                MessageBox.Show("Отдел добавлен");
                ReloadData2();
            }
        }

        private void btClear3_Click(object sender, EventArgs e)
        {
            textBox5.Clear();
            textBox7.Clear();
        }

        private void btAdd3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-VIGAS6C\SQLEXPRESS;Initial Catalog=Cadr;Integrated Security=True");

            sqlConnection.Open();

            SqlCommand command = new SqlCommand($"INSERT INTO [PositionAtWork] (Position, Salary) VALUES (@name, @sal)", sqlConnection);

            if (textBox5.Text.Length < 1 || textBox7.Text.Length < 1 )
            {
                MessageBox.Show("Зполните поля!");
            }
            else
            {
                command.Parameters.AddWithValue("name", textBox5.Text);
                command.Parameters.AddWithValue("sal", textBox7.Text);
                command.ExecuteNonQuery();
                MessageBox.Show("Должность добавлена");
                ReloadData3();
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            string sym = e.KeyChar.ToString();

            if (!Char.IsDigit(number) && number != 8 && (!Regex.Match(sym, @"[а-яА-Я]|[a-zA-Z]").Success))
            {
                e.Handled = true;
            }

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            string sym = e.KeyChar.ToString();

            if (!Char.IsDigit(number) && number != 8 &&  (!Regex.Match(sym, @"[а-яА-Я]|[a-zA-Z]").Success)) 
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            string sym = e.KeyChar.ToString();
            char number = e.KeyChar;

            if ((!Regex.Match(sym, @"[а-яА-Я]|[a-zA-Z]").Success) && (number != 8))
            {
                e.Handled = true;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
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
