using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace testBD
{
    public partial class Form1 : Form
    {
        // строка подключения к MS Access
        // вариант 1
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;";
        // вариант 2
        //public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Workers.mdb;";

        private OleDbConnection myConnect;
        public Form1()
        {
            InitializeComponent();
            myConnect = new OleDbConnection(connectString);
            myConnect.Open();
        }

        private void refreshView()
        {
            string query = "SELECT name_med, counter, arrival, sell FROM medic ORDER BY name_med";
            OleDbCommand command = new OleDbCommand(query, myConnect);
            OleDbDataReader reader = command.ExecuteReader();

            listView1.Items.Clear();
            int i = 1;
            while (reader.Read())
            {
                ListViewItem item =
                    listView1.Items.Add(i.ToString());
                item.SubItems.Add(reader["name_med"].ToString());
                item.SubItems.Add(reader["counter"].ToString());
                item.SubItems.Add(reader["arrival"].ToString());
                item.SubItems.Add(reader["sell"].ToString());
                i++;
            }

            reader.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            refreshView();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            myConnect.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "SELECT name_med FROM medic WHERE ID=1";

            OleDbCommand command = new OleDbCommand(query, myConnect);
            textBox1.Text = command.ExecuteScalar().ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            refreshView();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string insert = "INSERT INTO medic ([name_med], [counter], [arrival], [sell]) Values ('Баралгин', 10, 5, 15)";

            OleDbCommand command = new OleDbCommand(insert, myConnect);
            command.ExecuteNonQuery();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string query = "UPDATE medic SET [arrival] = 3, [counter] = 8 WHERE [name_med] = 'Анальгин'";
            OleDbCommand command = new OleDbCommand(query, myConnect);
            command.ExecuteNonQuery();
        }
    }
}
