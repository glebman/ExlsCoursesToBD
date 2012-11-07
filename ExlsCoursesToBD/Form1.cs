using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data.OleDb;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExlsCoursesToBD
{
    public partial class Form1 : Form
    {
        private string[] SheetsNames;
        private string[] FirstSubjectAdress;
        private string fName;
        private string lName;
        private string mName;
        private string bYear;
        private string mark;
        private string subject;

        private const  int subjectRowIndex =  3;
        private const int startIndex = 4;
        private const int NumberOfClasses = 24;

        DataTable dt;
        //        private string MyConnectionString = @"Server=GCOMPAQ\GLEB; Database=webdev_glekuz ;Trusted_Connection=Yes";
        private const string MyConnectionString = @"Server=MORIA; Database=webdev_glekuz ;Trusted_Connection=Yes";
        public Form1()
        {
            InitializeComponent();
            
            SheetsNames = new string[5];
            SheetsNames[0] = "ВПО";
            SheetsNames[1] = "Студенты";
            SheetsNames[2] = "ВПО_Академка";
            SheetsNames[3] = "Студенты_Академка";
            SheetsNames[4] = "ВПО_Отч.";

            FirstSubjectAdress = new string[5];
            FirstSubjectAdress[0] = "BS";
            FirstSubjectAdress[1] = "BO";
            FirstSubjectAdress[2] = "BS";
            FirstSubjectAdress[3] = "BO";
            FirstSubjectAdress[4] = "BS";
         }

        /*
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
 /
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
*/
        
        //open xml file
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
                for(int i=0 ; i<SheetsNames.Count() ; i++)
                {
                    //TODO post i in cycle
                    ReadSheet(filename, SheetsNames[0], i);
                }
                
            }
        }

        private void ReadSheet(string filename, string sheetName, int sheetNumber)
        {
            int ind = startIndex;//3
            string adressFIO = "E" + ind;
            string adressYear;
            string leterOfAdressubject = FirstSubjectAdress[sheetNumber];

            string fio = XLGetCellValue(filename, sheetName, adressFIO);
           
            while ( fio != null)
            {
                
                adressYear = "I" + ind; //year
                PrepareBirthdayDate(XLGetCellValue(filename, sheetName, adressYear));
                adressFIO = "E" + ind; //fio
                fio = XLGetCellValue(filename, sheetName, adressFIO);
                ParseFIO(fio);

                GetClassAndMark(filename, sheetName, leterOfAdressubject, ind); //отсюда вызов запроса к бд
                
                ind++;
            }
            
        }

        private string FindFirstSubject(string fileName, string sheetName)
        {
            string value = null;
            string firstSubject = "Алгоритмизация";

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
                    (WorksheetPart) (wbPart.GetPartById(theSheet.Id));
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                    Where(c => c.CellValue.ToString() == firstSubject).FirstOrDefault(); //error here
                if (theCell != null)
                    return theCell.CellReference;
                else //todo think about this stupid return statment
                    return " ";
            }
        }

        private void GetClassAndMark(string filename, string sheetName, string leterOfAdressubject, int ind)
        {
            //ищем предметы и оценки
            for (int i = 0; i < NumberOfClasses; i++)
            {
                if (i == 3 || i == 16) { leterOfAdressubject = Increment(leterOfAdressubject); }
                else
                {
                    //5 10 22 курсачи
                    subject = XLGetCellValue(filename, sheetName, leterOfAdressubject + subjectRowIndex);
                    mark = XLGetCellValue(filename, sheetName, leterOfAdressubject + ind);
                    leterOfAdressubject = Increment(leterOfAdressubject);

                    ExecStoredProcedure(" ", lName, fName, mName, subject, bYear, mark);
                }

            }
        }

        

        private void ParseFIO(string fio)
        {
            if (fio != null || fio != String.Empty)
            {
                // Define a regular expression for repeated words.
                Regex rx = new Regex(@"[a-zA-Zа-яА-я]+",
                  RegexOptions.Compiled | RegexOptions.IgnoreCase);

                // Find matches.
                MatchCollection matches = rx.Matches(fio);
                try
                {
                    //TODO call select
                    lName = matches[0].ToString();
                    fName = matches[1].ToString();
                    mName = matches[2].ToString();
                    
                    //ExecStoredProcedure("Persons_SelectByFullName", matches[0].ToString(), matches[1].ToString(),
                      //                  matches[2].ToString());
                }
                catch(Exception ex)
                {
                       
                }
            }

        }

  
        public void ExecStoredProcedure(string storeProcedureName, string ln, string fn, string mn, string courseName, string birthdayYear, string mark)
        {
            SqlConnection connection = new SqlConnection(MyConnectionString );
            try
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(storeProcedureName, connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("lName", ln);
                    command.Parameters.AddWithValue("fName", fn);
                    command.Parameters.AddWithValue("mName", mn);
                    command.Parameters.AddWithValue("year", birthdayYear);
                    command.Parameters.AddWithValue("course", courseName);
                    command.Parameters.AddWithValue("mark", mark);
                    
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            finally
            {
                connection.Close();
            }

        }
       



        public static string XLGetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                //Sheets sheets = wbPart.Workbook.GetFirstChild<Sheets>();
                //sheets.Elements<Sheet>(); 

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



        private DateTime ParseXLSDate(string date)
        {
            int y;
            int m;
            int d;
            if (date != null || date != String.Empty)
            {
                // Define a regular expression for repeated words.
                Regex rx = new Regex(@"[0-9]+",
                  RegexOptions.Compiled | RegexOptions.IgnoreCase);

                // Find matches.
                MatchCollection matches = rx.Matches(date);
                try
                {
                    //TODO call select
                    d = int.Parse(matches[0].ToString());
                    m = int.Parse(matches[1].ToString());
                    y = int.Parse(matches[2].ToString());
                    return new DateTime(y, m, d);
                }
                catch (Exception ex)
                {

                }
            }
            return new DateTime(1900,1,1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime oldDate = new DateTime(1900, 1, 1);
            oldDate = oldDate.AddDays(32497-2);
            string s = oldDate.ToShortDateString();



        }
        
        private void PrepareBirthdayDate(string bDate)
        {
            try
            {
                double d = double.Parse(bDate);
                DateTime date = new DateTime(1900, 1, 1);
                date = date.AddDays(d - 2);
                bYear = date.Year + @"-" + date.Month + @"-" + date.Day;
            }
            catch(Exception exception)
            {
                bYear= @"1900-01-01";
            }
           
        }

        static string Increment(string s)
        {

            // first case - string is empty: return "a"

            if ((s == null) || (s.Length == 0))

                return "a";

            // next case - last char is less than 'z': simply increment last char

            char lastChar = s[s.Length - 1];

            string fragment = s.Substring(0, s.Length - 1);

            if (lastChar < 'z')
            {

                ++lastChar;

                return fragment + lastChar;

            }

            // next case - last char is 'z': roll over and increment preceding string

            return Increment(fragment) + 'a';

        }
    }
}
