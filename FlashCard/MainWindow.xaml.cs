using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Win32;

namespace FlashCard
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Text2.Visibility = Visibility.Hidden;
        }
        public class Cards
        {

            public string questionList { get; set; }
            public string answerList { get; set; }
        }


       

        public List<Cards> FlashCards = new List<Cards>();

        int CardNumber = 0;
        public void getExcelFile(string filename)
        {

                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            string question = "";
            string answer = "";
            
            

            for (int i = 1; i <= rowCount; i++)
                {
                int k = 0;
               
                    for (int j = 1; j <= colCount ; j++)
                    {
                       
                         

                    //write the value to the console
                                       
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        if (k == 0)
                        {
                            question = i.ToString()+". " +xlRange.Cells[i, j].Value2.ToString();
                            question += "\n";
                            
                        }
                        else
                        {
                            if (k == 1)
                            {
                                answer = i.ToString() + ". " + xlRange.Cells[i, j].Value2.ToString();
                                answer += "\n";
                            }
                            else
                            {
                                answer += xlRange.Cells[i, j].Value2.ToString();
                                answer += "\n";
                            }
                        }
                    
                    if (k > 3) { k = 0; }

                    k++;
                       
                    }
                FlashCards.Add(new Cards()
                {
                    questionList = question,
                    answerList = answer,


                }) ;
                //Text1.AppendText(question);
                //Text2.AppendText(answer);
                
            }

            Text1.Text = FlashCards[CardNumber].questionList;
            Text2.Text = FlashCards[CardNumber].answerList;
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

        private void Next_Card(object sender, RoutedEventArgs e)
        {
            if (CardNumber < FlashCards.Count - 1)     { CardNumber += 1; }
            Text1.Text = FlashCards[CardNumber].questionList;
            Text2.Text = FlashCards[CardNumber].answerList;
        }
        private void Previews_Card(object sender, RoutedEventArgs e)
        {
            if (CardNumber > 0) { CardNumber -= 1; }

            Text1.Text = FlashCards[CardNumber].questionList;
            Text2.Text = FlashCards[CardNumber].answerList;
        }

        private void GotoAnswer(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Text1.Visibility = Visibility.Hidden;
            Text2.Visibility = Visibility.Visible;
        }

        private void GotoQuestion(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Text2.Visibility = Visibility.Hidden;
            Text1.Visibility = Visibility.Visible;
        }

        private void FlipCard(object sender, RoutedEventArgs e)
        {
            if (Text1.Visibility == Visibility.Hidden) { Text1.Visibility = Visibility.Visible; } else { Text1.Visibility = Visibility.Hidden; }
            if (Text2.Visibility == Visibility.Hidden) { Text2.Visibility = Visibility.Visible; } else { Text2.Visibility = Visibility.Hidden; }
           
        }

        private void OpenFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;

            choofdlog.Multiselect = false;
            choofdlog.ShowDialog();

            if (choofdlog.FileName != "") { getExcelFile(choofdlog.FileName); }
            
        }
    }

        
    
}
