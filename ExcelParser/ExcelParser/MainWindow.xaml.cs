using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls.Primitives;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab


namespace ExcelParser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public string directoryPath = "";
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkbook;
        public Excel._Worksheet xlWorksheet;
        public Excel.Range xlRange;

        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<string> chapterCollection = new ObservableCollection<string>();

        public MainWindow()
        {
            InitializeComponent();

            //getExcelFile();
        }

        public ObservableCollection<string> ChapterCollection
        {
            get
            {
                return chapterCollection;
            }
            set
            {
                chapterCollection = value;
                NotifyPropertyChanged();
            }
        }

        public void getExcelFile(string directoryPath)
        {
            try
            {
                InitializeExcelComponents(directoryPath);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }

            int startPoint = 0;
            bool startLock = false;
            int endPoint = 0;
            bool endLock = false;
            int currChapter = 0;


            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            currChapter = chapterList.SelectedIndex + 1;

            string[] reqLabels = (configReqTextBox.Text).Split(";");


            for (int i = 1; i <= rowCount; i++)
            {
                if (startPoint != 0 && endPoint != 0)
                {
                    break;
                }

                for (int j = 2; j <= 2; i++)     // 2 instead of colCount
                {
                    if (startPoint != 0 && endPoint != 0)
                    {
                        break;
                    }

                    var sno = (Excel.Range)xlWorksheet.Cells[i, j];

                    if (xlRange.Cells[i, j] != null && sno.Value2 != null)
                    {
                        if (startPoint != 0 && endPoint != 0)
                        {
                            break;
                        }
                        else if (sno.Value2.ToString().Substring(0, 1) == currChapter.ToString() && startLock == false)
                        {
                            startPoint = i;
                            startLock = true;
                        }
                        else if (sno.Value2.ToString().Substring(0, 1) == (currChapter + 1).ToString() && endLock == false)
                        {
                            endPoint = i - 1;
                            endLock = true;
                            break;
                        }
                    }
                }
            }



            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            for (int i = startPoint; i <= endPoint; i++)
            {
                for (int j = 4; j <= colCount; j++)
                {
                    var sno = (Excel.Range)xlWorksheet.Cells[i, j];

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && sno.Value2 != null)
                    {
                        //Console.Write(sno.Value2.ToString() + "\t");
                        foreach (string label in reqLabels)
                        {
                            if (sno.Value2.ToString() == (label + " "))
                            {
                                GenerateCAPLTestcase(directoryPath, xlWorksheet, currChapter, i);
                            }
                        }                        
                    }
                }
            }

            CodeCleanUp(xlApp, xlWorkbook, xlWorksheet, xlRange);

        }

        private void InitializeExcelComponents(string directoryPath)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\prithvi.patel\Downloads\ComponentRequirementSpecificationDU7ComfortDisplayBR12XX_ComponentRequirementSpecificationDocument_313827.xlsx");
            xlWorkbook = xlApp.Workbooks.Open(directoryPath.ToString());
            xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
        }

        private static void CodeCleanUp(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel._Worksheet xlWorksheet, Excel.Range xlRange)
        {
            /*
                Required to clean up after accessing workbook
            */

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void GenerateCAPLTestcase(string directoryPath, Excel._Worksheet xlWorksheet, int currChapter, int i)
        {
            string newDirExt = $"_chapter_{currChapter}.cin";
            string newDirPath = Path.ChangeExtension(directoryPath, newDirExt);

            string code = configCodeTextBox.Text;
            string modifiedCode = "\t" + code.Replace("\n", "\n\t");

            using (StreamWriter sw = File.AppendText(newDirPath))
            {
                //sw.WriteLine("Error_Flag = 'FOR_IMPORT' and location_type =   'Home' and batch_num = {0}", i);
                var req_ID = (Excel.Range)xlWorksheet.Cells[i, 1];
                //sw.WriteLine("Test Case : Requirement Comps ==> Req ID = {0} ", req_ID.Value2.ToString() );
                sw.WriteLine("export testcase_{0}", req_ID.Value2.ToString());
                sw.WriteLine("{");
                sw.WriteLine("\t TestDescription(\"Requirements: {0}\");", req_ID.Value2.ToString());
                sw.WriteLine("");
                sw.WriteLine(modifiedCode);
                sw.WriteLine("}");
                sw.WriteLine("");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            Nullable<bool> result = openFileDialog.ShowDialog();

            directoryPath = openFileDialog.FileName;
            InitializeExcelComponents(directoryPath);

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int NoofChapters = 0;

            for (int i = rowCount; i >= 1; i--)
            {
                if(NoofChapters != 0)
                {
                    break;
                }

                var sno = (Excel.Range)xlWorksheet.Cells[i, 2];

                if (xlRange.Cells[i, 2] != null && sno.Value2 != null)
                {
                    NoofChapters = Int32.Parse(sno.Value2.ToString().Substring(0, 1));
                }
            }

            ChapterCollection.Clear();
            for (int i = 1; i <= NoofChapters; i++)
            {
                ChapterCollection.Add(i.ToString());
            }



        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            getExcelFile(directoryPath);
        }

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged == null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private void verifyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InitializeExcelComponents(directoryPath);

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                string[] reqLabels = (configReqTextBox.Text).Split(";");



                string file = Path.GetFileNameWithoutExtension(directoryPath);
                string newPath = directoryPath.Replace(file, "invalid_requirements");
                string newPathCSV = Path.ChangeExtension(newPath, ".csv");

                System.Data.DataTable table = new System.Data.DataTable();

                for (int i = 1; i <= rowCount; i++)
                {
                    if (i == rowCount)
                        break;
                    for (int j = 4; j <= 4; i++)     // 2 instead of colCount
                    {
                        if (i == rowCount)
                            break;

                        var sno = (Excel.Range)xlWorksheet.Cells[i, j];

                        if (xlRange.Cells[i, j] != null && sno.Value2 != null)
                        {
                            //write the value to the console
                            if (xlRange.Cells[i, j] != null && sno.Value2 != null)
                            {
                                string snoValue = sno.Value2.ToString();

                                if(reqLabels.Any(x => (x + " ") == (snoValue)))
                                {
                                    break;
                                }
                                else if(snoValue == "")
                                {
                                    break;
                                }
                                else
                                {
                                    using (StreamWriter sw = File.AppendText(newPathCSV))
                                    {
                                        sw.WriteLine("{0}; {1}", i, snoValue);
                                    }
                                }
                            }
                        }
                    }
                }

                CodeCleanUp(xlApp, xlWorkbook, xlWorksheet, xlRange);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
