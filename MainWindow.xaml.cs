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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

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
            DataContext = this;
        }
        public string filePath1;
        public string filePath2;
        public string newFilePath;
        public int differencesCount;

        public int compliteCount = 0;
        public int ostatocCount = 0;

        public int position = -1;

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e) 
        {
            AdjustDataGridSize();
        }

        public void AdjustDataGridSize()
        {
            double newWidht1 = this.ActualWidth*0.49;
            //double newWidht2 = this.ActualHeight;

            double newHight = this.ActualHeight*0.4;

            DataGrid1.Width = newWidht1;
            DataGrid2.Width = newWidht1;
            //DataGrid3.Width = newWidht2;

            DataGrid1.Height = newHight;
            DataGrid2.Height = newHight;
            DataGrid3.Height = newHight;
        }

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
                int colIndex = GetColumnIndex(cellAddress);
                rowIndex = rowIndex - 2;

                string valueFromSecondFile = dataTable1.Rows[rowIndex][colIndex]?.ToString();

                dataTable3.Rows[rowIndex][colIndex] = valueFromSecondFile;

                ColorCompliteConflict(rowIndex, colIndex);
                if (ostatocCount <= differencesCount) 
                {
                    string myVariable2 = (differencesCount - ostatocCount).ToString();
                    TextBox2.Text = myVariable2;

                    string myVariable3 = compliteCount.ToString();
                    TextBox3.Text = myVariable3;
                }
            }
        }

        private void SelectFromSecondFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid3.SelectedItem is DataRowView selectedRow)
            {
                string cellAddress = Conflict[position];

                int rowIndex = GetRowIndex(cellAddress);
                int colIndex = GetColumnIndex(cellAddress);

                rowIndex = rowIndex - 2;


                string valueFromSecondFile = dataTable2.Rows[rowIndex][colIndex]?.ToString();

                dataTable3.Rows[rowIndex][colIndex] = valueFromSecondFile;

                ColorCompliteConflict(rowIndex, colIndex);
                if (ostatocCount <= differencesCount) 
                {
                    string myVariable2 = (differencesCount - ostatocCount).ToString();
                    TextBox2.Text = myVariable2;

                    string myVariable3 = compliteCount.ToString();
                    TextBox3.Text = myVariable3;
                }
            }
        }
        private void ColorCompliteConflict(int positionRow,int positionCol) 
        {
            Style myStyle3 = new Style(typeof(TextBlock));//TODO: изменить на DataGridCell и выделить в отдельную функцию

            myStyle3.Setters.Add(new Setter(TextBlock.BackgroundProperty, new SolidColorBrush(Colors.Green)));
            DataGridRow dr4 = (DataGrid3.ItemContainerGenerator.ContainerFromItem(DataGrid3.Items[positionRow]) as DataGridRow);
            FrameworkElement gridCell3 = null;
            if (dr4 != null)
                gridCell3 = DataGrid3.Columns[positionCol].GetCellContent(dr4);
            if (gridCell3 != null)
            {
                gridCell3.Style = myStyle3;
                gridCell3.UpdateLayout();
            }
            compliteCount++;
            ostatocCount++;

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
                //HighlightDifferences(worksheet2, Conflict);

                package1.Save();
                package2.Save();

                string myVariable1 = differencesCount.ToString();
                TextBox1.Text = myVariable1;

                string myVariable2 = differencesCount.ToString();
                TextBox2.Text = myVariable2;

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
                int rowNumber = GetRowIndex(address);

                rowNumber = rowNumber - 2;

                ColorRows(rowNumber,columnNumber);

                ExcelRange cell = worksheet.Cells[rowNumber + 1, columnNumber + 1];
                
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(highlightedFill);
            }
        }
        private void ColorRows(int positionRow,int positionCol) 
        {
            Style myStyle1 = new Style(typeof(TextBlock));//TODO: изменить на DataGridCell и выделить в отдельную функцию

            myStyle1.Setters.Add(new Setter(TextBlock.BackgroundProperty, new SolidColorBrush(Colors.Red)));

            DataGridRow dr1 = (DataGrid1.ItemContainerGenerator.ContainerFromItem(DataGrid1.Items[positionRow]) as DataGridRow);
            FrameworkElement gridCell = null;
            if (dr1 != null)
                gridCell = DataGrid1.Columns[positionCol].GetCellContent(dr1);
            if (gridCell != null)
            {
                gridCell.Style = myStyle1;
            }


            DataGridRow dr2 = (DataGrid2.ItemContainerGenerator.ContainerFromItem(DataGrid2.Items[positionRow]) as DataGridRow);
            FrameworkElement gridCell2 = null;
            if (dr2 != null)
                gridCell2 = DataGrid1.Columns[positionCol].GetCellContent(dr2);
            if (gridCell2 != null)
                gridCell2.Style = myStyle1;

            Style myStyle2 = new Style(typeof(TextBlock));//TODO: изменить на DataGridCell и выделить в отдельную функцию

            myStyle2.Setters.Add(new Setter(TextBlock.BackgroundProperty, new SolidColorBrush(Colors.Yellow)));
            DataGridRow dr3 = (DataGrid3.ItemContainerGenerator.ContainerFromItem(DataGrid3.Items[positionRow]) as DataGridRow);
            FrameworkElement gridCell3 = null;
            if (dr3 != null)
                gridCell3 = DataGrid3.Columns[positionCol].GetCellContent(dr3);
            if (gridCell3 != null)
            {
                gridCell3.Style = myStyle2;
            }
        }

    }
}
