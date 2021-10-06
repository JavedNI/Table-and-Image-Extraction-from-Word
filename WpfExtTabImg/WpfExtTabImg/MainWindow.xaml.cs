﻿using System;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Win32;

namespace WpfExtTabImg
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Global variable used in other functions 
        /// </summary>
        #region Global variable 
        static class Globals
        {
            public static string filePath;
        }
        #endregion 

        /// <summary>
        /// Uses openxml to extract the images as a stream of bytes
        /// </summary>
        /// <param name="folderPath"></param>
        #region ExtractImages
        static void ExtractImages(string folderPath)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(Globals.filePath, true))
            {
                //iterates through all the images and stores it as a list
                var imgList = document.MainDocumentPart.ImageParts.GetEnumerator();
                //counts the number of images 
                int imgNum = 0;

                //loops through imgList 
                while (imgList.MoveNext())
                {
                    imgNum++;
                    
                    //gets the current image in imgList
                    ImagePart imagePart = imgList.Current;

                    //gets the current image a as a stream of bytes, finds the length and stores the images bytes in an array 
                    Stream stream = imagePart.GetStream();
                    long length = stream.Length;
                    byte[] byteStream = new byte[length];
                    stream.Read(byteStream, 0, (int)length);

                    //define new folderpath and filetype string then split string based on '.'
                    char deliminterChar = '.';
                    string[] path = folderPath.Split(deliminterChar);

                    //saves the stream of bytes based on the file location selected, uses imgNum to iterate through the list of images 
                    if (imgNum <= 9)
                    {
                        using (var fstream = new FileStream($"{path[0]}00{imgNum}.{path[1]}", FileMode.OpenOrCreate, FileAccess.Write)) //images are in a different order to the number assigned to them 
                        {
                            stream.CopyTo(fstream);
                            fstream.Write(byteStream, 0, (int)length);
                            fstream.Close();
                        }
                    }
                    else if (imgNum > 10 && imgNum < 100)
                    {
                        using (var fstream = new FileStream($"{path[0]}0{imgNum}.{path[1]}", FileMode.OpenOrCreate, FileAccess.Write)) 
                        {
                            stream.CopyTo(fstream);
                            fstream.Write(byteStream, 0, (int)length);
                            fstream.Close();
                        }
                    }
                    else
                    {
                        using (var fstream = new FileStream($"{path[0]}{imgNum}.{path[1]}", FileMode.OpenOrCreate, FileAccess.Write)) 
                        {
                            stream.CopyTo(fstream);
                            fstream.Write(byteStream, 0, (int)length);
                            fstream.Close();
                        }
                    }                   
                }
            }
        }
        #endregion 

        #region ExtractTables
        static void ExtractTables(string excelPath)
        {
            #region App Checks
            //initialize excel application
            Excel.Application xlApp = new Excel.Application();
            Word.Application wApp = new Word.Application();

            //check if excel is installed on the device 
            if (xlApp == null | wApp == null) // can add specififc messages depending on which ones aren't installed 
            {
                MessageBox.Show("Excel and/or Word is not properly installed!!");
                return;
            }
            #endregion

            #region Create Documents
            Word.Document wDoc;
            object filename = Globals.filePath;
            object missing = Type.Missing;

            // excel workbook/sheet variables 
            Excel.Workbook xlBook;
            Excel.Worksheet xlSheet;
            object misValue = System.Reflection.Missing.Value;

            //create the excel file and sheet 
            xlBook = xlApp.Workbooks.Add(misValue);
            xlSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(1);

            xlSheet.Cells[1, 1] = "Image ID";
            xlSheet.Cells[1, 2] = "Age";
            xlSheet.Cells[1, 3] = "Gender";
            xlSheet.Cells[1, 4] = "Skin Tone";
            xlSheet.Cells[1, 5] = "Expression";
            xlSheet.Cells[1, 6] = "Shadow";
            xlSheet.Cells[1, 7] = "Glasses";
            xlSheet.Cells[1, 8] = "Beard";
            xlSheet.Cells[1, 9] = "Overweight";
            xlSheet.Cells[1, 10] = "Additional Comments";
            #endregion

            #region Read/Write Table Contents 
            try
            {
                // Open the Word file in "ReadOnly" mode.
                wDoc = wApp.Documents.Open(ref filename, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                // Check if Word file has any tables.
                if (wDoc.Tables.Count != 0)       
                {
                    //get the total number of tables in the document
                    int tableCount = wDoc.Tables.Count;

                    // Get total rows in the table.
                    int totalRowCount;
                    totalRowCount = wDoc.Tables[1].Rows.Count;

                    object[,] arrDoc = new object[tableCount, totalRowCount];

                    for (int i = 1; i <= tableCount; i++)
                    {
                        // Check if the table has columns.
                       // if (wDoc.Tables[i].Columns.Count > 0)     
                        //{
                            int rowCount;
                          
                            for (rowCount = 1; rowCount <= totalRowCount; rowCount++)
                            {
                                arrDoc[i - 1, rowCount - 1] = wDoc.Tables[i].Cell(rowCount, 2).Range.Text;                                
                            }                         
                        //}
                    }

                    Excel.Range startCell = (Excel.Range)xlSheet.Cells[2, 1];
                    Excel.Range endCell = (Excel.Range)xlSheet.Cells[tableCount + 1, totalRowCount];
                    Excel.Range writeRange = xlSheet.Range[startCell, endCell];
                    writeRange.Value2 = arrDoc;
                    xlBook.SaveAs(excelPath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);

                }
            }
            catch (Exception ex)
            {
                // Catch errors
                MessageBox.Show("error");
            }
            finally
            {
                // Clean up.
                wApp.Quit(); wApp = null;
                wDoc = null;           
                
                xlBook.Close(true, misValue, misValue);
                xlApp.Quit();

                //deals with how data is passed in args and return values between managed and unmanaged memory calls
                // after use sets everything in the excel file (cells, sheets etc) to null
                Marshal.ReleaseComObject(xlSheet);
                Marshal.ReleaseComObject(xlBook);
                Marshal.ReleaseComObject(xlApp);
            }           
        }
        #endregion 
        #endregion //comment this tomorrow if it doesn't get chnaged 

        /// <summary>
        /// Select File Button. Uses file dialog box to select a file to extract images and tables from
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        #region SelectDocButton
        private void SelectDocButton_Click(object sender, RoutedEventArgs e)
        {
            //uses dialog box to select a file to extract the images and tables information 
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                //specifies starting directory, type of file that can be saved and if the starting director is reset 
                InitialDirectory = @"C:\",
                Filter = "docx files (*.docx)|*.docx", //|All files (*.*)|*.*",
                RestoreDirectory = true,
            };

            //checks if a file has been selected 
            if (openFileDialog.ShowDialog() == true)
            {
                //stores filePath as a global variable to be acessed by other functions 
                Globals.filePath = openFileDialog.FileName;

                //pops up a message box to confirm the selected filepath and enabling other options is 'yes' is selected 
                if (MessageBox.Show($"Use this path? {Globals.filePath}", "File Selection", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    ExtImgButton.IsEnabled = true;
                    ExtTabButton.IsEnabled = true;
                }
            }
        }
        #endregion

        /// <summary>
        /// Extract Images Button. Uses a file dialog box as a path to save the images to a folder 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        #region ExtImgButton
        private void ExtImgButton_Click(object sender, RoutedEventArgs e)
        {
            //uses dialog box to select a location to save the image files 
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                //specifies starting directory, type of file that can be saved and if the starting director is reset 
                InitialDirectory = @"C:\",
                Filter = "Image files (*.jpg)|*.jpg",
                RestoreDirectory = true,
            };

            //checks if a file has been selected 
            if (saveFileDialog.ShowDialog() == true)
            {
                //saves the file dialog path as a string
                string folderPath = saveFileDialog.FileName;

                //calls the extract images function
                ExtractImages(folderPath);

                //confirms the images have been saved 
                MessageBox.Show($"Images have been saved", "Images Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                             
            }
        }
        #endregion

        /// <summary>
        /// Extract Tables Button. Uses a file dialog box as a path to save the excel file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        #region ExtTabButton
        private void ExtTabButton_Click(object sender, RoutedEventArgs e)
        {
            //uses dialog box to select a location to save the excel file 
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {        
                //specifies starting directory, type of file that can be saved and if the starting director is reset 
                InitialDirectory = @"C:\",
                Filter = "Excel files (*.xlsx)|*.xlsx",
                RestoreDirectory = true,
            };

            //checks if a file has been selected 
            if (saveFileDialog.ShowDialog() == true)
            {
                //saves the file dialog path as a string
                string excelPath = saveFileDialog.FileName;

                //calls the extract tables function
                ExtractTables(excelPath);

                //confirms the path in a message box 
                MessageBox.Show($"File Saved to: {excelPath}", "Excel File Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
    #endregion 
}
