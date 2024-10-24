using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Text;
using System.ComponentModel;
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        private SqlConnection sqlConnection = null;
        private DataTable table = null;
        private SqlDataAdapter adapter = null;
        public string Table_name = null;
        public string Columns_name = "*";

        public Form1()
        {
            

            InitializeComponent();
            

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=DESKTOP-DBKJNF6;Initial Catalog=БдСтраховаяКомпания; Integrated Security=True");
            sqlConnection.Open();
            DataTable allTables = sqlConnection.GetSchema("Tables");
            foreach (DataRow row in allTables.Rows)
            {
                string table_name = row["TABLE_NAME"].ToString();
                comboBox1.Items.Add(table_name);
            }
            comboBox1.SelectedIndex = 0;
            LoadDataForSelectedTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(adapter);
            adapter.Update(table);
            table.AcceptChanges();
        }

        private void LoadDataForSelectedTable()
        {
            comboBox2.ResetText();
            comboBox2.Items.Clear();
            Table_name = comboBox1.SelectedItem.ToString();
            DataTable Columns = sqlConnection.GetSchema("Columns", new string[] { null, null, Table_name });
            foreach (DataRow row in Columns.Rows)
            {
                string columnName = row["COLUMN_NAME"].ToString();
                comboBox2.Items.Add(columnName);
            }

            adapter = new SqlDataAdapter(@"SELECT " + Columns_name + " FROM " + Table_name, sqlConnection);
            view();
        }

        private void LoadDataForSelectedColumns()
        {
            Columns_name = comboBox2.SelectedItem.ToString();
            adapter = new SqlDataAdapter(@"SELECT " + Columns_name + " FROM " + Table_name, sqlConnection);
            view();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Columns_name = "*";
            LoadDataForSelectedTable();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadDataForSelectedColumns();
        }
        private void view()
        {
            try
            {
                table = new DataTable();

                adapter.Fill(table);

                dataGridView1.DataSource = table;

            }
            catch (SqlException err)
            {
                string error = err.Message;
                MessageBox.Show(error);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {

            dataGridView1.ClearSelection();
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Некорректный ввод поиска");
                return;
            }
            var value = textBox1.Text.Trim();
            int cnt = 0, numRow = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                var row = dataGridView1.Rows[i];
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (row.Cells[j].Value.ToString().Contains(value))
                    {
                        row.Selected = true;
                        cnt++; numRow = i; break;
                    }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            adapter = new SqlDataAdapter(@"SELECT * FROM " + Table_name, sqlConnection);
            comboBox2.ResetText();
            textBox1.Clear();
            view();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            int rowIndex = dataGridView1.CurrentCell.RowIndex;
            Console.WriteLine(rowIndex);
            dataGridView1.Rows.RemoveAt(rowIndex);
        }


        public void WordExport(DataGridView Grid, string path, string columns)
        {
            if (Grid == null)
            {
                MessageBox.Show("Ошибка экспорта");
                return;
            }
            var wordApp = new word.Application();
            wordApp.Visible = true;
            var wordDoc = wordApp.Documents.Add(path);
            if (wordDoc.ProtectionType != word.WdProtectionType.wdNoProtection)
            {
                wordDoc.Unprotect();
            }
            wordDoc.RemoveDocumentInformation(word.WdRemoveDocInfoType.wdRDIDocumentProperties);
            Replace_Time(wordDoc, "time", DateTime.Now);
            int rowCount = dataGridView1.Rows.Count;
            string min = dataGridView1[columns, 0].Value.ToString();
            string max = dataGridView1[columns, rowCount - 2 ].Value.ToString();
            Replace(wordDoc, "min", min);
            Replace(wordDoc, "max", max);
            InsertDataword(wordDoc, Grid);


        }
        private void InsertDataword(word.Document Doc, DataGridView Grid)
        {
            word.Range range = Doc.Content;
            range.Find.Execute(FindText: $"{{{"Table"}}}", Forward: true);

            var table = Doc.Tables.Add(range, Grid.Rows.Count, Grid.Columns.Count);
            for (int j = 0; j < Grid.Columns.Count; j++)
            {
                table.Rows[1].Cells[j + 1].Range.Text = Grid.Columns[j].HeaderText;
            }
            for (int i = 0; i < Grid.Rows.Count; i++)
            {
                for (int j = 0; j < Grid.Columns.Count; j++)
                {
                    if (Grid[j, i].Value != null)
                    {
                        table.Rows[i + 2].Cells[j + 1].Range.Text = Grid[j, i].Value.ToString();
                    }
                }
            }
            table.Borders.Enable = 1;

        }
        static void Replace_Time(word.Document Doc, string variable, DateTime value)
        {
            word.Range range = Doc.Content;
            range.Find.Execute(FindText: $"{{{variable}}}", Replace: word.WdReplace.wdReplaceAll, ReplaceWith: value);
        }
        static void Replace(word.Document Doc, string variable, string value)
        {
            word.Range range = Doc.Content;
            range.Find.Execute(FindText: $"{{{variable}}}", Replace: word.WdReplace.wdReplaceAll, ReplaceWith: value);
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
      
        public void temp_(int index, string path_word, string columns)
        {
            string fileSql = "C:\\Users\\angerquell\\Desktop\\Новая папка\\SQLQuery" + index +  ".sql";
            string sqlQuery = File.ReadAllText(fileSql, Encoding.Default);
            Console.Write(sqlQuery);
            adapter = new SqlDataAdapter(sqlQuery, sqlConnection);
            view();
            dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Descending);
            int rowCount = dataGridView1.Rows.Count;
            Console.Write(dataGridView1[columns,0].Value);
            Console.Write(dataGridView1[columns, rowCount - 2].Value);
            WordExport(dataGridView1, path: path_word, columns);
        }
     

        private void button5_Click(object sender, EventArgs e)
        {
            Form3 newForm = new Form3(this);
            newForm.Show();
        }
    }
}