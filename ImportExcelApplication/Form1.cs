using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportExcelApplication
    {
    public partial class Form1 : Form
        {
        static string currSheet = "";
        public Form1()
            {
            InitializeComponent();

            }

        private void button1_Click(object sender, EventArgs e)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

                if(openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string file_path = openFileDialog.FileName;
                    textBox1.Text = openFileDialog.FileName;
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Workbook workBook;
                    //Worksheet workSheet;

                    workBook = excel.Workbooks.Open(file_path);

                    //this for loop iterates through all the Excel sheets in the file and adds it to the combobox dropdown options
                    foreach(Worksheet curr_worksheet in workBook.Worksheets)
                    {
                        comboBox1.Items.Add(curr_worksheet.Name);
                    }

                    //Functionality for grid view preview of the Excel file
                    timer1.Start();

                }
            }

        private void upload_Click(object sender, EventArgs e)
            {
                //first check to make sure all the fields are filled before the user attempts to upload an Excel file
                if(textBox1.Text.Trim().Length == 0)
                {
                    MessageBox.Show("You need to select a file.", "File Select Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (string.IsNullOrEmpty(comboBox1.Text))
                {
                    MessageBox.Show("You need to select a sheet.", "Sheet Select Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if(textBox2.Text.Trim().Length == 0)
                {
                    MessageBox.Show("You need to enter a connection string.", "Connection String Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if(textBox3.Text.Trim().Length == 0)
                {
                    MessageBox.Show("You need to enter a schema.", "Schema Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (textBox4.Text.Trim().Length == 0)
                {
                    MessageBox.Show("You need to enter a table.", "Table Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    processExcelFile();
                }
        }

        private void processExcelFile()
        {
            MessageBox.Show("Attempting to start connection with the server...");
            try
            {
                //Setting up information for the SQl server connection string
                string sqlConnectionString = textBox2.Text;
                SqlConnection sqlConnection = new SqlConnection(sqlConnectionString);

                //Creating the Excel Connection String
                string excelConnectionString;
                string HDR = "YES";
                string fileFullPath = textBox1.Text;
                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + fileFullPath + "; Extended Properties = \"Excel 12.0 Xml; HDR=YES\";";
                string sheetName = comboBox1.Text;

                //Database Info
                string schemaName = textBox3.Text;
                string tableName = textBox4.Text;

                // Import data from Excel Sheet into a datatable object
                using (OleDbConnection conn = new OleDbConnection(excelConnectionString))
                {
                    conn.Open();
                    OleDbCommand dbCommand = new OleDbCommand("SELECT * FROM [" + sheetName + "$]", conn);

                    //OleDbDataAdapter acts as the bridge between the data source and creating a data set/data table
                    OleDbDataAdapter adapter = new OleDbDataAdapter(dbCommand);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    conn.Close();
                    sqlConnection.Open();

                    // Import data from the DataTable object into the specified SQL server table
                    using (SqlBulkCopy bulkcopy = new SqlBulkCopy(sqlConnection))
                    {
                        bulkcopy.DestinationTableName = schemaName + "." + tableName;
                        bulkcopy.WriteToServer(dataTable);

                    }
                    sqlConnection.Close();
                    MessageBox.Show("Finished processing Excel File");
                }

            }
            catch(Exception exception)
            {
                // If anything fails (e.g. connection to SQL server), error message will be logged to the console
                MessageBox.Show(exception.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            /* Every 100 milliseconds checkSheetSelect() is called to determine if the preview of the data table in the grid view
             * needs to be updated*/
            checkSheetSelect();
        }

        private void checkSheetSelect()
        {
            /* first make sure that a Excel file and sheet has been selected and that the selected sheet is different from the current displayed sheet
             * to prevent unecessary updates*/
            if(string.IsNullOrEmpty(comboBox1.Text) == false && comboBox1.Text != currSheet && textBox1.Text.Trim().Length > 0)
            {
                currSheet = comboBox1.Text;
                string excelConnectionString;
                string HDR = "YES";
                string fileFullPath = textBox1.Text;
                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + fileFullPath + "; Extended Properties = \"Excel 12.0 Xml; HDR=YES\";";
                string sheetName = comboBox1.Text;

                // Import data from Excel Sheet into a datatable object
                using (OleDbConnection conn = new OleDbConnection(excelConnectionString))
                {
                    conn.Open();
                    OleDbCommand dbCommand = new OleDbCommand("SELECT * FROM [" + sheetName + "$]", conn);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(dbCommand);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    // binds the dataTable to the dataGridView
                    dataGridView1.DataSource = dataTable;
                    conn.Close();
                }
            }
        }
    }
    }
