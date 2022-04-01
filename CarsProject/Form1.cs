using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;


namespace CarsProject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

       

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        string conectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Tanya\OneDrive\Documents\Cars3.mdb";
        OleDbConnection DBConect = new OleDbConnection();
        private void Form2_Load(object sender, EventArgs e)
        {
            displaydata();
           // renamecolumn();
        }

        //private void renamecolumn()
        //{
        //    dataGridView1.Columns[0].HeaderText = "ID номер";
        //    dataGridView1.Columns[1].HeaderText = "Рег. номер";
        //    dataGridView1.Columns[2].HeaderText = "Марка";
        //    dataGridView1.Columns[3].HeaderText = "Година на производство";
        //    dataGridView1.Columns[4].HeaderText = "Цена";
        //    dataGridView1.Columns[5].HeaderText = "Цвят";
        //}

        private void displaydata()
        {
            string myselect = "Select * From Koli";
            DBConect.ConnectionString = conectionstring;
            DBConect.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter(myselect, DBConect);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            DBConect.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DBConect.ConnectionString = conectionstring;
            string myselect = "Insert into Koli(RegNomer, Marka, GodinaProizvodstvo, Cena, Color) values('" + textBox2.Text + "', '" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "', '" + textBox6.Text + "')";
            OleDbCommand dbComand = new OleDbCommand(myselect, DBConect);
            DBConect.Open();
            dbComand.CommandText = myselect;
            dbComand.Connection = DBConect;
            dbComand.ExecuteNonQuery();
            MessageBox.Show("Записано", "Поздравления");
            DBConect.Close();
            displaydata();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DBConect.ConnectionString = conectionstring;
            string myselect = "update Koli set RegNomer= '" + textBox2.Text + "', Marka= '" + textBox3.Text + "', GodinaProizvodstvo= '" + textBox4.Text + "', Cena= '" + textBox5.Text + "', Color = '" + textBox6.Text + "' where KoliID = " + textBox1.Text;
            OleDbCommand dbComand = new OleDbCommand(myselect, DBConect);
            DBConect.Open();
            dbComand.CommandText = myselect;
            dbComand.Connection = DBConect;
            dbComand.ExecuteNonQuery();
            MessageBox.Show("Променено", "Поздравления");
            DBConect.Close();
            displaydata();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DBConect.ConnectionString = conectionstring;
            string myselect = "Delete from Koli where KoliID = " + textBox1.Text;
            OleDbCommand dbComand = new OleDbCommand(myselect, DBConect);
            DBConect.Open();
            dbComand.CommandText = myselect;
            dbComand.Connection = DBConect;
            dbComand.ExecuteNonQuery();
            MessageBox.Show("Записът е изтрит", "Поздравления");
            DBConect.Close();
            displaydata();
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            foreach (DataGridViewTextBoxColumn column in dataGridView1.Columns)
            {
                dt.Columns.Add(column.Name, column.ValueType);
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dr = dt.NewRow();
                foreach (DataGridViewTextBoxColumn column in dataGridView1.Columns)
                {
                    if (row.Cells[column.Name].Value != null)
                    {
                        dr[column.Name] = row.Cells[column.Name].Value.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            string filePath = saveFileDialog1.FileName;
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = saveFileDialog1.FileName;
            }
            DataTableToTextFile(dt, filePath);

        }

        private void DataTableToTextFile(DataTable dt, string outputFilePath)
        {
            int[] maxLengths = new int[dt.Columns.Count];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                maxLengths[i] = dt.Columns[i].ColumnName.Length;
                foreach (DataRow row in dt.Rows)
                {
                    if (!row.IsNull(i))
                    {
                        int length = row[i].ToString().Length;
                        if (length > maxLengths[i])
                        {
                            maxLengths[i] = length;
                        }
                    }
                }
            }

            using (StreamWriter sw = new StreamWriter(outputFilePath, false))
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
                }

                sw.WriteLine();

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (!row.IsNull(i))
                        {
                            sw.Write(row[i].ToString().PadRight(maxLengths[i] + 2));
                        }
                        else
                        {
                            sw.Write(new string(' ', maxLengths[i] + 2));
                        }
                    }
                    sw.WriteLine();
                }
                sw.Close();
            }
        }

    }
}

