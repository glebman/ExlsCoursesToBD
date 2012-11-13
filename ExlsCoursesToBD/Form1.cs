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
        private List<string> SheetsNames;
        private List<string> FirstSubjectAdress;
        private string fName;
        private string lName;
        private string mName;
        private string bYear;
        private int mark;
        private string subject;

        private const  int subjectRowIndex =  3;
        private const int startIndex = 4;
        private  int NumberOfClasses = 24;

        private List<string> subjectsList;


        DataTable dt;
        //        private string MyConnectionString = @"Server=GCOMPAQ\GLEB; Database=webdev_glekuz ;Trusted_Connection=Yes";
        private const string MyConnectionString = @"Server=MORIA; Database=webdev_glekuz ;Trusted_Connection=Yes";
        public Form1()
        {
            InitializeComponent();
            
           
           
            

            
         }
        private bool InitCollections()
        {
            SheetsNames = new List<string>();
            FirstSubjectAdress = new List<string>();
            subjectsList = new List<string>();

            #region pageName init
            if (tbpage1.Text != string.Empty && tbcolumn1.Text != string.Empty)
            {
                SheetsNames.Add(tbpage1.Text);
                FirstSubjectAdress.Add(tbcolumn1.Text);
            }
            if (tbpage2.Text != string.Empty && tbcolumn2.Text != string.Empty)
            {
                SheetsNames.Add(tbpage2.Text);
                FirstSubjectAdress.Add(tbcolumn2.Text);
            }
            if (tbpage3.Text != string.Empty && tbcolumn3.Text != string.Empty)
            {
                SheetsNames.Add(tbpage3.Text);
                FirstSubjectAdress.Add(tbcolumn3.Text);
            }
            if (tbpage4.Text != string.Empty && tbcolumn4.Text != string.Empty)
            {
                SheetsNames.Add(tbpage4.Text);
                FirstSubjectAdress.Add(tbcolumn4.Text);
            }
            if (tbpage5.Text != string.Empty && tbcolumn5.Text != string.Empty)
            {
                SheetsNames.Add(tbpage5.Text);
                FirstSubjectAdress.Add(tbcolumn5.Text);
            }
            if (tbpage6.Text != string.Empty && tbcolumn6.Text != string.Empty)
            {
                SheetsNames.Add(tbpage6.Text);
                FirstSubjectAdress.Add(tbcolumn6.Text);
            }
            if (tbpage7.Text != string.Empty && tbcolumn7.Text != string.Empty)
            {
                SheetsNames.Add(tbpage7.Text);
                FirstSubjectAdress.Add(tbcolumn7.Text);
            }
            if (tbpage8.Text != string.Empty && tbcolumn8.Text != string.Empty)
            {
                SheetsNames.Add(tbpage8.Text);
                FirstSubjectAdress.Add(tbcolumn8.Text);
            }
            if (tbpage9.Text != string.Empty && tbcolumn9.Text != string.Empty)
            {
                SheetsNames.Add(tbpage9.Text);
                FirstSubjectAdress.Add(tbcolumn9.Text);
            }
            if (tbpage10.Text != string.Empty && tbcolumn10.Text != string.Empty)
            {
                SheetsNames.Add(tbpage10.Text);
                FirstSubjectAdress.Add(tbcolumn10.Text);
            }
            try
            {
                NumberOfClasses = int.Parse(tbNumberOfColumns.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("In 'NumberOfClasses' field must be a number");
                return false;
            }
            #endregion

            return true;
        }
      
        //open xml file
        private void button2_Click(object sender, EventArgs e)
        {
            if (InitCollections())
            {
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel (*.xlsx)|*.xlsx";
                opf.ShowDialog();
                string filename = opf.FileName;
                if (filename == "")
                {
                    MessageBox.Show("select the file");
                }
                else
                {
                    for (int i = 0; i < SheetsNames.Count(); i++)
                    {
                        ReadSheet(filename, SheetsNames[i], i);
                    }
                    MessageBox.Show("Work done!");

                }
            }
        }

        private void ReadSheet(string filename, string sheetName, int sheetNumber)
        {
            int ind = startIndex-1;//3
            string adressFIO = "E" + (ind+1);
            string adressYear;
            string leterOfAdressubject = FirstSubjectAdress[sheetNumber];
            string fio;

            for (; ;)
            {
                ind++;
                adressFIO = "E" + ind; //fio
                fio = XLGetCellValue(filename, sheetName, adressFIO);
                if(string.IsNullOrEmpty(fio)) break;
                ParseFIO(fio);
                
                adressYear = "I" + ind; //year
                if (!PrepareBirthdayDate(XLGetCellValue(filename, sheetName, adressYear)))
                {
                    Logger(DateTime.Now + " - Wrong date, sheet:" + sheetName + ", Cell: I" + ind);
                    continue;
                }

                if (!CheckPerson())
                {
                    Logger(DateTime.Now + " - There is no person or There are more then one person whith this FIO and year, sheet: " + sheetName + ", Cell: E" + ind);
                    continue;
                }
                
                GetClassAndMark(filename, sheetName, leterOfAdressubject, ind); //отсюда вызов запроса к бд
             }
        }

   

        private void GetClassAndMark(string filename, string sheetName, string leterOfAdressubject, int ind)
        {
            subjectsList = new List<string>();
            string textMark;
            //ищем предметы и оценки
            for (int i = 0; i < NumberOfClasses; i++)
            {

                //5 10 22 курсачи
                subject = XLGetCellValue(filename, sheetName, leterOfAdressubject + subjectRowIndex);

                textMark = XLGetCellValue(filename, sheetName, leterOfAdressubject + ind);
                ParseMark(textMark);


                leterOfAdressubject = Increment(leterOfAdressubject);

                if (subjectsList.Contains(subject))
                    ExecStoredProcedure("GradeBook_InsertByFIO_YEAR_CourseName", lName, fName, mName, subject,
                                        bYear, mark, 0, 0, sheetName, ind);
                else if (textMark == "зачтено")
                    ExecStoredProcedure("GradeBook_InsertByFIO_YEAR_CourseName", lName, fName, mName, subject, bYear,
                                        mark, 0, 2, sheetName, ind);

                else if (!string.IsNullOrEmpty(subject))
                    ExecStoredProcedure("GradeBook_InsertByFIO_YEAR_CourseName", lName, fName, mName, subject,
                                        bYear, mark, 0, 1, sheetName, ind);
                subjectsList.Add(subject);


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
                    mark = 5; break; //?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????
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
//                MessageBox.Show(exception.Message);
                Logger(DateTime.Now + "   " + exception.Message + "  On operation whith -  FIO: "+lName+" "+fName+" "+mName+" year: "+bYear+" subject: "+ subject);
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

                string input = fio;
                string pattern = "ё";
                string replacement = "е";
                Regex rgx = new Regex(pattern);
                string result = rgx.Replace(input, replacement);
                fio = result;

                // Define a regular expression for repeated words.
                Regex rx = new Regex(@"[a-zA-Zа-яА-яёЁ]+",
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

         /// <summary>
  /// 
  /// </summary>
  /// <param name="storeProcedureName"></param>
  /// <param name="ln"></param>
  /// <param name="fn"></param>
  /// <param name="mn"></param>
  /// <param name="courseName"></param>
  /// <param name="birthdayYear"></param>
  /// <param name="mark"></param>
  /// <param name="procentMark"></param>
  /// <param name="examType">0-курсовик  
  ///                        1 - экзамен
  ///                        2 - зачет
  /// </param>
        public void ExecStoredProcedure(string storeProcedureName, string ln, string fn, string mn, string courseName, string birthdayYear, int mark, int procentMark, int examType, string sheetName, int rowNumber)
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
                    command.Parameters.AddWithValue("gradeID", mark);
                    command.Parameters.AddWithValue("mark", procentMark);
                    command.Parameters.AddWithValue("exam", examType);
                    
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception exception)
            {
                Logger(DateTime.Now + " - exeption in store procedure at sheet:" + sheetName + ", Row:" + rowNumber + ".  Exeption message:");
                Logger(exception.Message);
                Logger("Parameters: " + "lName - " + ln + ", fName - " + fn + ", mName - " + mn + ", year - "+ birthdayYear + ", course - "+ courseName + ", gradeID - "+ mark + ", mark - "+ procentMark);
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
            
            tbpage1.Text = AppDomain.CurrentDomain.BaseDirectory;

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

            System.IO.StreamWriter file = new System.IO.StreamWriter(AppDomain.CurrentDomain.BaseDirectory+"Logg.txt", true);
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

        private void btnDefaultValues_Click(object sender, EventArgs e)
        {
            
            tbpage1.Text = "ВПО";
            tbpage2.Text = "Студенты";
            tbpage3.Text = "ВПО_Академка";
            tbpage4.Text = "Студенты_Академка";
            tbpage5.Text = "ВПО_Отч.";
            tbpage6.Text = string.Empty;
            tbpage7.Text = string.Empty;
            tbpage8.Text = string.Empty;
            tbpage9.Text = string.Empty;
            tbpage10.Text = string.Empty;
            
            tbcolumn1.Text = "BS";
            tbcolumn2.Text = "BO";
            tbcolumn3.Text = "BS";
            tbcolumn4.Text = "BO";
            tbcolumn5.Text = "BS";
            tbcolumn6.Text = string.Empty;
            tbcolumn7.Text = string.Empty;
            tbcolumn8.Text = string.Empty;
            tbcolumn9.Text = string.Empty;
            tbcolumn10.Text = string.Empty;

            tbNumberOfColumns.Text = 24.ToString();
        }
    }
}
