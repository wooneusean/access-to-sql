using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AccessToSQL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }
        string src = "";
        string table = "";
        List<object> cols = new List<object>();
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            src = srcTxt.Text;
            table = tableCB.Text;
            string output = outTxt.Text;
            string scriptString = "";
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + src;
            Console.WriteLine(connString);
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT * FROM " + table, connection);
                Console.WriteLine(command);
                reader = command.ExecuteReader();



                var tabName = reader.GetSchemaTable();
                var nameCol = tabName.Columns["ColumnName"];
                foreach (DataRow row in tabName.Rows)
                {
                    cols.Add("`" + row[nameCol] + "`");
                }
                var result = string.Join(",", cols);
                Console.WriteLine(result);



                scriptString = "INSERT INTO `" + table + "` (" + result + ") VALUES\n";
                Console.WriteLine(scriptString);
                List<object> stringArr = new List<object>();
                foreach (DataRow row in tabName.Rows)
                {
                    stringArr = tabName.AsEnumerable().Select(r => r["ColumnName"]).ToList();
                }
                var strValue = string.Join(",",stringArr);
                Console.WriteLine(strValue);



                while (reader.Read())
                {
                    scriptString += "(" + strValue + "),\n";
                    Console.WriteLine(scriptString);
                }
            }
            File.WriteAllText(output, scriptString += ";");
        }

        private void srcTxt_TextChanged(object sender, TextChangedEventArgs e)
        {
            
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            string src = srcTxt.Text;
            try
            {
                string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + src;
                using (OleDbConnection connection = new OleDbConnection(connString))
                {
                    connection.Open();
                    DataTable schema = connection.GetSchema("Tables");
                    foreach (DataRow row in schema.Rows)
                    {
                        tableCB.Items.Add(row.Field<string>("TABLE_NAME"));
                    }

                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
