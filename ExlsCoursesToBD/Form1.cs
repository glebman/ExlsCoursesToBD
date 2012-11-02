using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExlsCoursesToBD
{
    public partial class Form1 : Form
    {
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }

        public void read()
        {
            
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.xlsx)|*.xlsx";
            opf.ShowDialog();
            string filename = opf.FileName;
            if (filename == "")
            {
                MessageBox.Show("Файл не выбран");
            }
            else
            {
                OleDbConnection theConnection = new OleDbConnection(string.Format("provider=Microsoft.ACE.OLEDB.12.0;data source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"", filename));
                //"HDR=Yes;" indicates that the first row contains columnnames, not data. "HDR=No;" indicates the opposite.
 /* If you want to read the column headers into the result set (using HDR=NO even though there is a header) and the column data is numeric, use IMEX=1 to avoid crash.
To always use IMEX=1 is a safer way to retrieve data for mixed data columns. Consider the scenario that one Excel file might work fine cause that file's data causes the driver to guess one data type while another file, containing other data, causes the driver to guess another data type. This can cause your app to crash.
 */
                theConnection.Open();
                OleDbDataAdapter theDataAdapter = new OleDbDataAdapter("SELECT * FROM [Студенты$]", theConnection);
                //DataSet theDS = new DataSet();
                dt = new DataTable();
                theDataAdapter.Fill(dt);
                this.dataGridView1.DataSource = dt.DefaultView;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            read();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.xlsx)|*.xlsx";
            opf.ShowDialog();
            string filename = opf.FileName;
            if (filename == "")
            {
                MessageBox.Show("Файл не выбран");
            }
            else
            {
                string fio;
                fio = XLGetCellValue(filename, "Студенты", "E3");
                textBox1.Text = fio;

                richTextBox1.Text += XLGetCellValue(filename, "Студенты", "E3");
                ParseFIO(fio);
            }
        }

        private void ParseFIO(string fio)
        {
            if (fio != null || fio != String.Empty)
            {

            }

        }




        public static string XLGetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that Sheet
                // object to retrieve a reference to the appropriate worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part, and then use its 
                // Worksheet property to get a reference to the cell whose 
                // address matches the address you supplied:
                WorksheetPart wsPart =
                  (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string:
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents a numeric value, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and Booleans
                    // individually. For shared strings, the code looks up the 
                    // corresponding value in the shared string table. For Booleans, 
                    // the code converts the value into the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                // For shared strings, look up the value in the shared 
                                // strings table.
                                var stringTable = wbPart.
                                  GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                // If the shared string table is missing, something is 
                                // wrong. Return the index that you found in the cell.
                                // Otherwise, look up the correct text in the table.
                                if (stringTable != null)
                                {
                                    value = stringTable.SharedStringTable.
                                      ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Date:
                                value = theCell.CellValue.ToString();
                                break;


                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                        //datetime
                    else
                    {

                    }
                }
            }
            return value;
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
