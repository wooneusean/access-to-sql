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
        public string GetConverterOutput(string sourceDB, string sourceTable, string outputPath)
        {
            List<object> cols = new List<object>();
            List<string> stringArr = new List<string>();
            List<string> scriptValues = new List<string>();
            string parsedValues = "";
            string scriptString = "";
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceDB;
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT * FROM " + sourceTable, connection);
                reader = command.ExecuteReader();



                var tabName = reader.GetSchemaTable();
                var nameCol = tabName.Columns["ColumnName"];
                foreach (DataRow row in tabName.Rows)
                {
                    cols.Add("`" + row[nameCol] + "`");
                }
                var result = string.Join(",", cols);
                cols.Clear();

                scriptString = "INSERT INTO `" + sourceTable + "` (" + result + ") VALUES \n";


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
                File.WriteAllText(outputPath, scriptString += ";");
                string retValue = scriptString += ";";
                scriptString = "";
                return retValue;
            }
        }
        public string RunConverter(string sourceDB, string sourceTable, string outputPath)
        {
            List<object> cols = new List<object>();
            List<string> stringArr = new List<string>();
            List<string> scriptValues = new List<string>();
            string parsedValues = "";
            string scriptString = "";
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceDB;
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT * FROM " + sourceTable, connection);
                reader = command.ExecuteReader();



                var tabName = reader.GetSchemaTable();
                var nameCol = tabName.Columns["ColumnName"];
                foreach (DataRow row in tabName.Rows)
                {
                    cols.Add("`" + row[nameCol] + "`");
                }
                var result = string.Join(",", cols);
                cols.Clear();

                scriptString = "INSERT INTO `" + sourceTable + "` (" + result + ") VALUES \n";


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
                File.WriteAllText(outputPath, scriptString += ";");
                scriptString = "";
            }
        }
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            GetConverterOutput(srcTxt.Text, tableCB.SelectedItem.ToString(), outTxt.Text);
        }

        private void srcTxt_TextChanged(object sender, TextChangedEventArgs e)
        {
            
        }
        public bool GetTables(string sourceTable,ComboBox destinationComboBox)
        {
            destinationComboBox.Items.Clear();
            try
            {
                string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceTable;
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
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (GetTables(srcTxt.Text, tableCB))
            {
                inDirLbl.Content = "Access Database Directory:";
            }
            else
            {
                inDirLbl.Content = "Path Invalid!";
            }
        }

        private void outTxt_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void tableCB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            outTxt.Text = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + tableCB.SelectedItem.ToString() + ".sql";
        }
    }
}
