using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace RefEditor
{
    public partial class Form3 : Form
    {
        private OleDbConnection connection;
        private OleDbCommand command;
        private OleDbDataAdapter adapter;
        private DataSet dataset;
        private DataTable t; 
     
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConnectToData();
            t = dataset.Tables["Reference"];
            // http://www.codeproject.com/Articles/7858/This-is-a-simple-C-program-that-illustrate-the-usa 
            //t.RowChanged += new DataRowChangeEventHandler(Row_Changed); 
            dataGridView1.DataSource = t;  
        
        }

        public void ConnectToData()
        {
            connection = new OleDbConnection();
            command = new OleDbCommand();
            adapter = new OleDbDataAdapter();
            dataset = new DataSet();
            // edit database location using  AppDomain.CurrentDomain.BaseDirectory   of similar constants
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "/refs.accdb;" +
            "Persist Security Info=False";
            command.Connection = connection;
            command.CommandText = "SELECT * FROM reference";
            adapter.SelectCommand = command;
            try
            {
                adapter.Fill(dataset, "Reference");
                /*
                DataRowCollection rows = dataset.Tables["reference"].Rows;
                foreach (DataRow row in rows)
                {
                    String str = row[2].ToString();
                    MessageBox.Show(str);
                }
                // MessageBox.Show(AppDomain.CurrentDomain.BaseDirectory);  
                */ 
            }
            catch (OleDbException)
            {
                MessageBox.Show("Error occured while connecting to database.");
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form3_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'gbua_x_datac31aDataSet.reference' table. You can move, or remove it, as needed.
            //this.referenceTableAdapter.Fill(this.gbua_x_datac31aDataSet.reference);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataset != null)
            {
                string search_string = textBox1.Text.ToString();
                try
                {
                    string expression = "title LIKE `*" + search_string + "*`";
                    DataRow[] results = dataset.Tables["Reference"].Select(expression);
                    dataGridView2.DataSource = results.CopyToDataTable();
                    //ShowDialog
                }catch(Exception exc){};


            }
        } 


    }
}
