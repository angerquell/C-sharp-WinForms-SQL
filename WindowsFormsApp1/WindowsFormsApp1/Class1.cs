using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
namespace WindowsFormsApp1
{
    class DataBase
    {
        const string Connect_Basa = @"Data Source=DESKTOP-DBKJNF6;Initial Catalog=2-43 Variant 6B KK; Integrated Security=True";
        SqlConnection sqlConnection = new SqlConnection(Connect_Basa);
        SqlDataAdapter phoneAdapter;
        SqlCommandBuilder phoneAdapterBuilder;
        DataSet dataSet = new DataSet();

        BindingSource BindingId = new BindingSource();
        DataGridView DataGridviewId = new DataGridView();
        BindingNavigator bindingNavigatorId = new BindingNavigator();
        public void Connect()
        {
            if (sqlConnection.State == System.Data.ConnectionState.Closed)
            {
                sqlConnection.Open();
            }
            else
            {
                sqlConnection.Close();
            }
        }
        public SqlConnection GetConnection()
        {
            return sqlConnection;
        }
        public void DataApater_Fill()
        {
            phoneAdapter = new SqlDataAdapter(@"SELECT * FROM Phone_conversation", sqlConnection);
            phoneAdapter.Fill(dataSet, "Phone_conversation");
            phoneAdapterBuilder = new SqlCommandBuilder(phoneAdapter);
        }
        public void DataRelation()
        {
            DataRelation dataRelation = new DataRelation("IdPhone", dataSet.Tables["Phone_covnersation"].Columns["id"], dataSet.Tables["Abonent"].Columns["id"]);
            dataSet.Relations.Add(dataRelation);
        }
        public void Update(Form form)
        {
            dataSet = new DataSet();

            DataApater_Fill();
            form.Controls.Add(DataGridviewId);
            DataGridviewId.Dock = DockStyle.Fill;
            form.Controls.Add(bindingNavigatorId);
            bindingNavigatorId.Dock = DockStyle.Bottom;

            BindingId.DataSource = dataSet.Tables["Phone_covnersation"];
            DataGridviewId.DataSource = BindingId;
            bindingNavigatorId.BindingSource = BindingId;
        }
       
    }
}
