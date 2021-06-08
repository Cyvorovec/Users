using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using ClosedXML.Excel;

namespace Test
{
    public partial class Form1 : Form
    {
        string path = "Users.db";

        public Form1()
        {
            InitializeComponent();
        }

        private void data_show()
        {
            SQLiteConnection DB = new SQLiteConnection("Data Source=Users.db ; Version = 3");
            DB.Open();
            dataGridView1.Rows.Clear();
            string com = "select ID, FullName, Login, Date from Users where DeleteFlag = 0";
            SQLiteCommand cmd = new SQLiteCommand(com, DB);
            SQLiteDataReader DR = cmd.ExecuteReader();

            while (DR.Read())
            {
                dataGridView1.Rows.Insert(0, DR.GetInt32(0), DR.GetString(1), DR.GetString(2), DR.GetString(3));
            }
            DB.Close();
        }

        private void Create_DB()
        {
            if (!System.IO.File.Exists(path))
            {
                SQLiteConnection.CreateFile(path);
                using (var sqlite = new SQLiteConnection(@"Data Source = " + path))
                {
                    sqlite.Open();
                    string com = "CREATE TABLE Users(ID INTEGER NOT NULL UNIQUE,FullName TEXT,Login TEXT UNIQUE,Date REAL,DeleteFlag  INTEGER,PRIMARY KEY (ID AUTOINCREMENT))";
                    SQLiteCommand cmd = new SQLiteCommand(com, sqlite);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            Create_DB();
            data_show();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }


        private void Add_Click(object sender, EventArgs e)
        {
            if (textBoxFioAdd.Text != "" && textBoxLogAdd.Text != "")
            {
                SQLiteConnection DB = new SQLiteConnection("Data Source=Users.db ; Version = 3");
                DB.Open();
                DateTime rowdate = DateTime.Now;
                string date = rowdate.ToString("yyyy-MM-dd");
                SQLiteCommand command = DB.CreateCommand();
                command.CommandText = "INSERT INTO Users(FullName, Login, Date, DeleteFlag) VALUES (@FullName, @Login, '" + date + "', 0)";
                command.Parameters.Add("@FullName", System.Data.DbType.String).Value = textBoxFioAdd.Text.ToUpper();
                command.Parameters.Add("@Login", System.Data.DbType.String).Value = textBoxLogAdd.Text.ToUpper();
                command.ExecuteNonQuery();

                data_show();


                DB.Close();

            }
            else
            {
                MessageBox.Show("Введите ФИО и Логин!");
            }
        }


        private void Load_Click(object sender, EventArgs e)
        {
            SQLiteConnection DB = new SQLiteConnection("Data Source=Users.db ; Version = 3");
            DB.Open();
            SQLiteDataAdapter DA = new SQLiteDataAdapter("select *from Users", DB);
            DataTable DT = new DataTable();
            DA.Fill(DT);
            using (SaveFileDialog SFD = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (SFD.ShowDialog() == DialogResult.OK)
                {
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        workbook.Worksheets.Add(DT,"Users");
                        workbook.SaveAs(SFD.FileName);
                    }
                }
            }
            DB.Close();
        }

        private void Del_Click(object sender, EventArgs e)
        {
            if (textBoxFioDel.Text != "" && textBoxLogDel.Text != "")
            {
                SQLiteConnection DB = new SQLiteConnection("Data Source=Users.db ; Version = 3");
                DB.Open();
                SQLiteCommand command = DB.CreateCommand();
                command.CommandText = "UPDATE Users Set DeleteFlag = 1 where Login =@Login and FullName=@FullName";
                command.Parameters.Add("@FullName", System.Data.DbType.String).Value = textBoxFioDel.Text.ToUpper();
                command.Parameters.Add("@Login", System.Data.DbType.String).Value = textBoxLogDel.Text.ToUpper();
                command.ExecuteNonQuery();


                data_show();

                DB.Close();
            }
            else
            {
                MessageBox.Show("Введите ФИО и Логин!");
            }
        }

        private void Upd_Click(object sender, EventArgs e)
        {
            if (textBoxIDSearch.Text != "")
            {
                SQLiteConnection DB = new SQLiteConnection("Data Source=Users.db ; Version = 3");
                DB.Open();
                SQLiteCommand com = DB.CreateCommand();
                if (textBoxLogUpd.Text != "")
                {
                    com.CommandText = "UPDATE Users Set Login = @Login where Id =@Id";
                    com.Parameters.Add("@Login", System.Data.DbType.String).Value = textBoxLogUpd.Text.ToUpper();
                    com.Parameters.Add("@Id", System.Data.DbType.Int32).Value = textBoxIDSearch.Text;
                    com.ExecuteNonQuery();

                }
                if (textBoxFioUpd.Text != "")
                {
                    com.CommandText = "UPDATE Users Set FullName = @FullName where Id =@Id";
                    com.Parameters.Add("@FullName", System.Data.DbType.String).Value = textBoxFioUpd.Text.ToUpper();
                    com.Parameters.Add("@Id", System.Data.DbType.Int32).Value = textBoxIDSearch.Text;
                    com.ExecuteNonQuery();
                }

                if (string.IsNullOrEmpty(textBoxFioUpd.Text) && string.IsNullOrEmpty(textBoxLogUpd.Text))
                {
                    MessageBox.Show("Введите данные на обновление!");
                }

                data_show();
                DB.Close();
            }
        }
    }
}
