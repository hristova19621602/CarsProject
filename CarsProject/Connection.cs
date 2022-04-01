using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;


namespace CarsProject
{
    internal class Connection
    {
        OleDbConnection connection;
        OleDbCommand command;
        private void ConnectTo()
        {
            connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\Tanya\\OneDrive\\Documents\\Cars.accdb");
            command = connection.CreateCommand();
        }
        public Connection()
        {
            ConnectTo();
        }

        public void Insert(Prodajbi p)
        {
            try
            {
                //command.CommandText = "INSERT INTO Bikes(ModelID, VersionID, Price, [Note]) VALUES (1, 2, 1, 'asa')";
                command.CommandText = "INSERT INTO Prodajbi(ProdajbiID, KodProdajba, EGN, RegNomer, DataProdajba) VALUES (" + p.ProdajbiID 
                    + "," + p.KodProdajba + "," + p.EGN + "," + "'" + p.RegNomer + p.DataProdajba + "'" + ")";
                command.CommandType = CommandType.Text;
                // command.Connection = connection;
                connection.Open();
                command.ExecuteNonQuery();
            }
            catch (Exception)
            {
                //throw;
                MessageBox.Show("Некоректни данни! Моля въведете отново!");
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
        }



    }
}
