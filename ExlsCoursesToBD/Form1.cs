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
        private int mark;
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
                if(!PrepareBirthdayDate(XLGetCellValue(filename, sheetName, adressYear)))
                {
                    Logger("Wrong date at"+sheetName+" in row №"+ind);
                    continue;
                }
                adressFIO = "E" + ind; //fio
                fio = XLGetCellValue(filename, sheetName, adressFIO);
                ParseFIO(fio);

                if (!CheckPerson())
                {
                    Logger("Wrong person info or DB has duplicates " + sheetName + " in row №" + ind);
                    continue;
                }
                GetClassAndMark(filename, sheetName, leterOfAdressubject, ind); //отсюда вызов запроса к бд
                
                ind++;
            }
            
        }

   

        private void GetClassAndMark(string filename, string sheetName, string leterOfAdressubject, int ind)
        {
            string textMark;
            //ищем предметы и оценки
            for (int i = 0; i < NumberOfClasses; i++)
            {
                if (i == 3 || i == 16) { leterOfAdressubject = Increment(leterOfAdressubject); }
                else
                {
                    //5 10 22 курсачи
                    subject = XLGetCellValue(filename, sheetName, leterOfAdressubject + subjectRowIndex);
                    textMark = XLGetCellValue(filename, sheetName, leterOfAdressubject + ind);
                    ParseMark(textMark);

                    leterOfAdressubject = Increment(leterOfAdressubject);
                    if(i==5 || i== 10 || i ==22)
                        ExecStoredProcedure("GradeBook_InsertByFIO_YEAR_CourseName", lName, fName, mName, subject, bYear, mark,0, false);
                    else ExecStoredProcedure("GradeBook_InsertByFIO_YEAR_CourseName", lName, fName, mName, subject, bYear, mark,0, true);
                }

            }
        }

        private void ParseMark(string textMark)
        {
            switch (textMark)
            {
                case "отлично":
                    mark = 5;
                    break;
                case "хорошо":
                    mark = 4; break;
                case "удовлетворительно":
                    mark = 3; break;
                case "зачтено":
                    mark = -1; break; //?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????
                default:
                    mark = 0; break;

            }
        }

        private bool CheckPerson()
        {
            SqlConnection connection = new SqlConnection(MyConnectionString);
            try
            {
                connection.Open();
                string cmdText = String.Format("SELECT COUNT(*) FROM APersons WHERE  LastName Like N'{0}' and FirstName Like N'{1}' and MiddleName Like N'{2}' and Birthday = '{3}';",lName,fName,mName,bYear);
                
                using (SqlCommand command = new SqlCommand(cmdText, connection))
                {


                  int res = (int)command.ExecuteScalar();
                  if (res == 1) return true;
                    
                    return false;

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
            return false;

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

  
        public void ExecStoredProcedure(string storeProcedureName, string ln, string fn, string mn, string courseName, string birthdayYear, int mark, int procentMark, bool exam = true)
        {
            SqlConnection connection = new SqlConnection(MyConnectionString );
            try
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(storeProcedureName, connection))
                {
                    int ex = exam ? 1 : 0;
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("lName", ln);
                    command.Parameters.AddWithValue("fName", fn);
                    command.Parameters.AddWithValue("mName", mn);
                    command.Parameters.AddWithValue("year", birthdayYear);
                    command.Parameters.AddWithValue("course", courseName);
                    command.Parameters.AddWithValue("gradeID", mark);
                    command.Parameters.AddWithValue("mark", procentMark);
                    command.Parameters.AddWithValue("exam",ex );
                    
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
            
            textBox1.Text = AppDomain.CurrentDomain.BaseDirectory;

        }
        
        private bool PrepareBirthdayDate(string bDate)
        {
            try
            {
                double d = double.Parse(bDate);
                DateTime date = new DateTime(1900, 1, 1);
                date = date.AddDays(d - 2);
                bYear = date.Year + @"-" + date.Month + @"-" + date.Day;
                return true;
            }
            catch(Exception exception)
            {
                
                bYear= @"1900-01-01";
                return false;
            }
           
        }

        public void Logger(String lines)
        {

            // Write the string to a file.append mode is enabled so that the log
            // lines get appended to  test.txt than wiping content and writing the log

            System.IO.StreamWriter file = new System.IO.StreamWriter(AppDomain.CurrentDomain.BaseDirectory+"test.txt", true);
            file.WriteLine(lines);

            file.Close();

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
