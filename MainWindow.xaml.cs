using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows.Media;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OfficeOpenXml;
using System.Data;
using System.Security.Cryptography.X509Certificates;
using System.Drawing;
using System.Text.RegularExpressions;
using OfficeOpenXml.Style;
using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.Windows.Controls.Primitives;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace INTANT_Task
{
    public partial class MainWindow : Window
    {
        private DataTable dataTable1;
        private DataTable dataTable2;
        private DataTable dataTable3;
        private int SelectedColumnIndex;
        private List<string> Conflict;
        public MainWindow()
        {
            InitializeComponent();
            dataTable1 = new DataTable();
            dataTable2 = new DataTable();
            dataTable3 = new DataTable();
            Conflict = new List<string>();
        }
        public string filePath1;
        public string filePath2;
        public string newFilePath;
        public int differencesCount;
        public int position = -1;



        private void LoadingFirstFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";

            bool? result = openFileDialog.ShowDialog();

            if (result == true) 
            {
                string SelectedFile1Path = openFileDialog.FileName;
                filePath1 = openFileDialog.FileName;
                try 
                {
                    using (var package = new ExcelPackage(new FileInfo(SelectedFile1Path))) 
                    {
                        dataTable1.Clear();
                        dataTable3.Clear();

                        var workbook = package.Workbook;
                        if (workbook.Worksheets.Count >0) 
                        {
                            var worksheet = workbook.Worksheets[1];
                            DataGrid1.Columns.Clear();
                            DataGrid1.Items.Clear();

                            DataGrid3.Columns.Clear();
                            DataGrid3.Items.Clear();

                            dataTable1 = new DataTable();
                            dataTable3 = new DataTable();

                            //Заполняем загаловки столбцов
                            foreach (var cell  in worksheet.Cells[1,1,1, worksheet.Dimension.Columns]) 
                            {
                                dataTable1.Columns.Add(cell.Value?.ToString());
                                dataTable3.Columns.Add(cell.Value?.ToString());
                            }

                            //Заполняем данные строк
                            for (int row = 2; row <= worksheet.Dimension.Rows; row++) 
                            {
                                var data1Row = dataTable1.NewRow();
                                var data3Row = dataTable3.NewRow();
                                for (int col = 1; col <= worksheet.Dimension.Columns; col++) 
                                {
                                    data1Row[col - 1] = worksheet.Cells[row, col].Value;
                                    data3Row[col - 1] = worksheet.Cells[row, col].Value;
                                }
                                dataTable1.Rows.Add(data1Row);
                                dataTable3.Rows.Add(data3Row);
                                
                            }
                            if (string.IsNullOrEmpty(dataTable1.Columns[dataTable1.Columns.Count - 1].ColumnName)) 
                            {
                                dataTable1.Columns.RemoveAt(dataTable1.Columns.Count - 1);
                            }

                            if (string.IsNullOrEmpty(dataTable3.Columns[dataTable3.Columns.Count - 1].ColumnName))
                            {
                                dataTable3.Columns.RemoveAt(dataTable3.Columns.Count - 1);
                            }

                            if (DataGrid1.ItemsSource is DataView dataView1)
                            {
                                dataView1.Table.Rows.Clear();
                                foreach (DataRow row in dataTable1.Rows)
                                {
                                    dataView1.Table.ImportRow(row);
                                }
                                dataView1.Table.AcceptChanges();
                                DataGrid1.Items.Refresh();
                            }
                            else
                            {
                                DataGrid1.ItemsSource = dataTable1.DefaultView;
                                DataGrid1.Items.Refresh();
                            }


                            if (DataGrid3.ItemsSource is DataView dataView3)
                            {
                                dataView3.Table.Rows.Clear();
                                foreach (DataRow row in dataTable3.Rows)
                                {
                                    dataView3.Table.ImportRow(row);
                                }
                                dataView3.Table.AcceptChanges();
                                DataGrid3.Items.Refresh();
                            }
                            else
                            {
                                DataGrid3.ItemsSource = dataTable3.DefaultView;
                                DataGrid3.Items.Refresh();
                            }
                        }
                    }
                }
                catch 
                {
                    MessageBox.Show($"Ошибка при чтении файла");
                }
            }
        }

        private void LoadingSecondFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                string SelectedFile2Path = openFileDialog.FileName;
                filePath2 = openFileDialog.FileName;
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(SelectedFile2Path)))
                    {
                        dataTable2.Clear();

                        var workbook = package.Workbook;
                        if (workbook.Worksheets.Count > 0)
                        {
                            var worksheet = workbook.Worksheets[1];
                            DataGrid2.Columns.Clear();
                            DataGrid2.Items.Clear();

                            dataTable2 = new DataTable();

                            //Заполняем загаловки столбцов
                            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                            {
                                dataTable2.Columns.Add(cell.Value?.ToString());
                            }

                            //Заполняем данные строк
                            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                            {
                                var dataRow = dataTable2.NewRow();
                                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                                {
                                    dataRow[col - 1] = worksheet.Cells[row, col].Value;
                                }
                                dataTable2.Rows.Add(dataRow);
                            }
                            if (string.IsNullOrEmpty(dataTable2.Columns[dataTable2.Columns.Count - 1].ColumnName))
                            {
                                dataTable2.Columns.RemoveAt(dataTable2.Columns.Count - 1);
                            }

                            if (DataGrid2.ItemsSource is DataView dataView2)
                            {
                                dataView2.Table.Rows.Clear();
                                foreach (DataRow row in dataTable2.Rows)
                                {
                                    dataView2.Table.ImportRow(row);
                                }
                                dataView2.Table.AcceptChanges();
                                DataGrid2.Items.Refresh();
                            }
                            else
                            {
                                DataGrid2.ItemsSource = dataTable2.DefaultView;
                                DataGrid2.Items.Refresh();
                            }
                        }
                    }
                }
                catch
                {
                    MessageBox.Show($"Ошибка при чтении файла");
                }
            }
        }
        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            CompareExcelFiles(filePath1,filePath2);
        }

        private void SelectFromFirstFileButton_Click(object sender, RoutedEventArgs e) 
        {
            if (DataGrid3.SelectedItem is DataRowView selectedRow)
            {
                string cellAddress = Conflict[position];

                int rowIndex = GetRowIndex(cellAddress);
                rowIndex = rowIndex - 2;

                string valueFromSecondFile = dataTable1.Rows[rowIndex][SelectedColumnIndex]?.ToString();

                dataTable3.Rows[rowIndex][SelectedColumnIndex] = valueFromSecondFile;

                DataGrid3.ItemsSource = dataTable3.DefaultView;
                DataGrid3.Items.Refresh();
            }
        }

        private void SelectFromSecondFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid3.SelectedItem is DataRowView selectedRow)
            {
                string cellAddress = Conflict[position];

                int rowIndex = GetRowIndex(cellAddress);
                rowIndex = rowIndex - 2;


                string valueFromSecondFile = dataTable2.Rows[rowIndex][SelectedColumnIndex]?.ToString();

                dataTable3.Rows[rowIndex][SelectedColumnIndex] = valueFromSecondFile;

                DataGrid3.ItemsSource = dataTable3.DefaultView;
                DataGrid3.Items.Refresh();
            }
        }

        private void ShowValueNextConflictButton_Click(object sender, RoutedEventArgs e) 
        {
            
            if (position >= differencesCount-1)
            {
                MessageBox.Show("Вы дошли до крайнего конфликта");
            }
            else 
            {
                position++;
                DisplayConflict(position, Conflict);
            }

            
        }

        private void ShowValuePreviousConflictButton_Click(object sender, RoutedEventArgs e)
        {

            if (position <= 0)
            {
                MessageBox.Show("Вы дошли до первого конфликта");
            }
            else
            {
                position--;
                DisplayConflict(position, Conflict);
            }

            
        }

        private void DisplayConflict(int conflictIndex,List<string> conflicts) 
        {
            //Очищаем выбранные ячейки в третьем DataGrid
            DataGrid3.UnselectAllCells();

            if (dataTable3.Rows != null) 
            {
                string cellAddress = conflicts[conflictIndex];

                int rowIndex = GetRowIndex(cellAddress);

                DataGrid3.SelectedIndex = rowIndex-2;
                DataGrid3.ScrollIntoView(DataGrid3.SelectedItem);

                //Обновляем индекс текущей выбранной колонки
                SelectedColumnIndex = rowIndex - 2;
            }
        }

        private int GetRowIndex(string address) 
        {
            return int.Parse(Regex.Replace(address,"[^0-9]+",""));
        }
        private int GetColumnIndex(string address)
        {
            foreach (char c in address) 
            {
                if (char.IsLetter(c)) 
                {
                    char uppercaseChar = char.ToUpper(c);

                    return (int)uppercaseChar - 65;
                }
            }
            return 0;
        }

        #region SelectAll
        private void SaveUserChoice(string selectedFilePath) 
        {
            FileInfo selectedFile = new FileInfo(selectedFilePath);

            using (ExcelPackage selectedPackage = new ExcelPackage(selectedFile)) 
            {
                ExcelWorksheet selectedWorksheet = selectedPackage.Workbook.Worksheets[1];

                string cellAddress = Conflict[position];

                int rowIndex = GetRowIndex(cellAddress);
                int columnIndex = GetColumnIndex(cellAddress);
                DataRow dataRow = dataTable3.Rows[rowIndex-2];

                for (int row = 2; row <= rowIndex; row++)
                {
                    for (int col = 0; col <= columnIndex; col++)
                    {
                        DataColumn dataColumn = dataTable3.Columns[col];

                        string cellValue = selectedWorksheet.Cells[row, col+1].GetValue<string>()?.Trim();

                        dataRow[dataColumn] = cellValue;
                    }
                }
                DataGrid3.ItemsSource = dataTable3.DefaultView;
                DataGrid3.Items.Refresh();
            }
        }
        #endregion

        private void SaveNewFileButton_Click(object sender, RoutedEventArgs e) 
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Выберите место сохранения нового файла";

            bool? result = saveFileDialog.ShowDialog();
            if (result == true) 
            {
                string savePath = saveFileDialog.FileName;

                FileInfo newFile = new FileInfo(savePath);
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    int rowCount = dataTable3.Rows.Count;
                    int colCount = dataTable3.Columns.Count;

                    //Заполняем заголовки столбцов
                    for (int col = 1; col <= colCount; col++)
                    {
                        worksheet.Cells[1, col].Value = dataTable3.Columns[col - 1].ColumnName;
                    }

                    //Заполняем данные строк
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            string cellValue = dataTable3.Rows[row-1][col-1]?.ToString();
                            worksheet.Cells[row + 1, col].Value = cellValue;
                            Console.WriteLine($"Cell [{row+1},{col}]:{cellValue}");
                        }
                    }
                    package.Save();
                    MessageBox.Show("Новый файл сохранен.");
                    /*DataGrid1.Columns.Clear();
                    DataGrid2.Columns.Clear();
                    DataGrid3.Columns.Clear();
                    
                    dataTable1.Clear();
                    dataTable2.Clear();
                    dataTable3.Clear();*/
                }
            }
        }

        public  void CompareExcelFiles(string filePath1, string filePath2)
        {

            FileInfo file1 = new FileInfo(filePath1);
            FileInfo file2 = new FileInfo(filePath2);

            using (ExcelPackage package1 = new ExcelPackage(file1))
            using (ExcelPackage package2 = new ExcelPackage(file2))
            {
                ExcelWorksheet worksheet1 = package1.Workbook.Worksheets[1];
                ExcelWorksheet worksheet2 = package2.Workbook.Worksheets[1];

                List<string> differences = new List<string>();

                int rowCount = Math.Max(worksheet1.Dimension.End.Row, worksheet2.Dimension.End.Row);
                int colCount = Math.Max(worksheet1.Dimension.End.Column, worksheet2.Dimension.End.Column);

                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue1 = worksheet1.Cells[row, col].GetValue<string>()?.Trim();
                        string cellValue2 = worksheet2.Cells[row, col].GetValue<string>()?.Trim();

                        if (!AreValuesEqual(cellValue1, cellValue2))
                        {
                            string cellAdress = $"{GetColumnLetter(col)}{row}";
                            //differences.Add($"Различие в ячейке {cellAdress}: {cellValue1} != {cellValue2}");

                            differencesCount++;
                            Conflict.Add( cellAdress );
                        }
                    }
                }
                
                HighlightDifferences(worksheet1, Conflict);
                HighlightDifferences(worksheet2, Conflict);

                package1.Save();
                package2.Save();

                MessageBox.Show($"Сравнение завершено, количество различий {differencesCount}");
            }
        }

        private static bool AreValuesEqual(object value1, object value2)
        {
            if (value1 == null && value2 == null)
            {
                return true;
            }
            else if (value1 == null || value2 == null)
            {
                return false;
            }
            else if (value1 is string && value2 is string)
            {
                return string.Equals((string)value1, (string)value2);
            }
            else if (IsNumericType(value1) && IsNumericType(value2))
            {
                double numValue1 = Convert.ToDouble(value1);
                double numValue2 = Convert.ToDouble(value2);
                return numValue1 == numValue2;
            }
            else
            {
                return false;
            }
        }

        private static bool IsNumericType(object value)
        {
            return value is int || value is double || value is float || value is decimal;
        }

        private static string GetColumnLetter(int columnNumber)
        {
            if (columnNumber < 1)
            {
                throw new ArgumentException("номер столбца не может быть меньше чем 1");
            }

            int dividend = columnNumber;
            string columnLetter = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnLetter = (char)('A' + modulo) + columnLetter;
                dividend = (dividend - modulo) / 26;
            }
            return columnLetter;
        }

        private static int GetColumnNumber(string address)
        {
            foreach (char c in address)
            {
                if (char.IsLetter(c))
                {
                    char uppercaseChar = char.ToUpper(c);

                    return (int)uppercaseChar - 65;
                }
            }
            return 0;
        }

        private static T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }

        private  void HighlightDifferences(ExcelWorksheet worksheet, List<string> differences)
        {
            var highlightedFill = System.Drawing.Color.FromArgb(255,255,0,0);
            foreach (var address in differences)
            {
                int columnNumber = GetColumnNumber(address);
                string rowNumber = Regex.Replace(address, @"[^\d]+", string.Empty);

                if (!int.TryParse(rowNumber, out int rowIndex))
                {
                    continue;
                }

                int columnIndex = columnNumber;

                ExcelRange cell = worksheet.Cells[rowIndex, columnIndex + 1];

                /*dataTable2.Rows[rowIndex][SelectedColumnIndex]?*/
                DataRow dr = (DataRow)dataTable1.Rows[rowIndex][columnIndex];
                //DataGrid1.SelectedCells[0].Item.


                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(highlightedFill);

                /*var mergedCells = worksheet.MergedCells;
                foreach (var mergedCell in mergedCells)
                {
                    ExcelCellAddress startAddress = new ExcelCellAddress(mergedCell);
                    ExcelCellAddress endAddress = new ExcelCellAddress(mergedCell.Split('!')[1]);

                    if (startAddress.Row <= rowIndex && rowIndex <= endAddress.Row &&
                        startAddress.Column <= columnIndex + 1 && columnIndex + 1 <= endAddress.Column)
                    {
                        for (int i = startAddress.Row; i <= endAddress.Row; i++)
                        {
                            for (int j = startAddress.Column; j <= endAddress.Column; j++)
                            {
                                ExcelRange range = worksheet.Cells[i, j];
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(highlightedFill);
                            }
                        }
                        break;
                    }
                }*/


            }
        }
    }
}
