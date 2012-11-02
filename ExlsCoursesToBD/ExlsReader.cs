using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExlsCoursesToBD
{
    class ExlsReader
    {

       



        /*
        public void ToExcel()
        {
            object misValue = System.Reflection.Missing.Value;

            var xlApp = new Application();
            var xlWorkBook = xlApp.Workbooks.Add(misValue);

            if (xlWorkBook != null)
            {
                var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];
                #region CELLS
                xlWorkSheet.Cells[1, 1] = "Число";
                xlWorkSheet.Cells[1, 2] = "Время";

                xlWorkSheet.Cells[1, 3] = "U";
                //xlWorkSheet.Cells[1, 4] = "(%)";
                xlWorkSheet.Cells[1, 5] = "P";
                //xlWorkSheet.Cells[1, 6] = "(мм.рт.ст.)";
                xlWorkSheet.Cells[1, 7] = "t";
                //xlWorkSheet.Cells[1, 8] = "(С)";
                xlWorkSheet.Cells[1, 9] = "Осадки";

                xlWorkSheet.Cells[2, 3] = "от";
                xlWorkSheet.Cells[2, 4] = "до";
                xlWorkSheet.Cells[2, 5] = "от";
                xlWorkSheet.Cells[2, 6] = "до";
                xlWorkSheet.Cells[2, 7] = "от";
                xlWorkSheet.Cells[2, 8] = "до";
                xlWorkSheet.Cells[2, 9] = "от";
                xlWorkSheet.Cells[2, 10] = "до";

                for (int i = 0; i < _length; i++)
                {
                    for (int j = 0; j < MaxParams - 2; j++)
                    {

                        xlWorkSheet.Cells[i + 3, j + 2] = _parsedMeteoData[i, j];

                    }

                    if (i % 2 == 0)
                    {
                        xlWorkSheet.Cells[i + 3, 9] = "";
                        xlWorkSheet.Cells[i + 3, 10] = "";
                    }
                    else
                    {
                        xlWorkSheet.Cells[i + 3, 9] = _parsedMeteoData[i, 7];
                        xlWorkSheet.Cells[i + 3, 10] = _parsedMeteoData[i, 8];
                    }
                    if (i % 8 == 0) xlWorkSheet.Cells[i + 3, 1] = (i / 8) + 1;

                }
                #endregion


                var saveFileDialog1 = new SaveFileDialog
                {
                    Filter = @"xls files (*.xls)|*.xls",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // Code to write the stream goes here."zontik.xls"
                    xlWorkBook.SaveAs(saveFileDialog1.FileName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }


                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                ReleaseObject(xlWorkSheet);
            }
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

        }
        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Exception Occured while releasing object {0}", ex));
            }
            finally
            {
                GC.Collect();
            }
        }
         * */
    }
         
}
