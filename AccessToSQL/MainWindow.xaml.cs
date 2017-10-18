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
        List<string> stringArr = new List<string>();
        List<string> scriptValues = new List<string>();
        string parsedValues = "";
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            src = srcTxt.Text;
            table = tableCB.Text;
            string output = outTxt.Text;
            string scriptString = "";
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + src;
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT * FROM " + table, connection);
                reader = command.ExecuteReader();



                var tabName = reader.GetSchemaTable();
                var nameCol = tabName.Columns["ColumnName"];
                foreach (DataRow row in tabName.Rows)
                {
                    cols.Add("`" + row[nameCol] + "`");
                }
                var result = string.Join(",", cols);
                cols.Clear();

                scriptString = "INSERT INTO `" + table + "` (" + result + ") VALUES \n";


                while (reader.Read())
                {
                    int i = 0 - 1;
                    while (i++ < tabName.Rows.Count - 1)
                    {
                        stringArr.Add("'" + reader.GetValue(i).ToString() + "'");
                    }
                    var resultRow = string.Join(",", stringArr);
                    scriptValues.Add("(" + resultRow + ")");
                    parsedValues = string.Join(",\n", scriptValues);
                    stringArr.Clear();
                }
                scriptString += parsedValues;
                File.WriteAllText(output, scriptString += ";");
            }
            
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
                    string[] restrictions = new string[4];
                    restrictions[3] = "Table";
                    // Get list of user tables
                    var tabName = connection.GetSchema("Tables", restrictions);
                    for (int i = 0; i < tabName.Rows.Count; i++)
                    {
                        tableCB.Items.Add(tabName.Rows[i][2].ToString());
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
