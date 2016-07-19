using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Globalization;

namespace RefEditor
{
    public partial class Form1 : Form
    {
        private const String myConnectionString = @"server=mysql301.1gb.ua;port=3306;database=gbua_x_datac31a;userid=gbua_x_datac31a;password=a6b0a10bem1;";
        private MySqlConnection con;
        private MySqlDataReader reader;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            con = null;
            reader = null;
            try
            {
                con = new MySqlConnection(myConnectionString);
                con.Open();
                DateTime saveNow = DateTime.Now;


                CultureInfo enUS = CultureInfo.CreateSpecificCulture("en-US");
                DateTimeFormatInfo dtfi = enUS.DateTimeFormat;
                dtfi.ShortDatePattern = "yyyy-MM-dd HH:mm:ss";               
                String data_values = "1, '"+textBox1.Text+"', 'paper', '"+saveNow.ToString("d", enUS)+"', 1";
                String data_table = "reference";
                String data_columns = "ID, title, type, edit_time, ID_owner";
                String cmdText = "INSERT INTO " + data_table + "(" + data_columns + ") VALUES(" + data_values + ")";

                MySqlCommand cmd = new MySqlCommand(cmdText, con);
                //cmd.Prepare();
                //cmd.Parameters.AddWithValue("@name", "your value here");
                cmd.ExecuteNonQuery();
            }
            catch (MySqlException err)
            {
                Console.WriteLine("Error: " + err.ToString());
            }
            finally
            {
                if (con != null)
                {
                    con.Close(); //close the connection
                }
            } //remember to close the connection after accessing the database

            this.Close();
        }
    }
}
