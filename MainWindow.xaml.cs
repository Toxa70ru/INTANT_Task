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

namespace INTANT_Task
{
    public partial class MainWindow : Window
    {
        private DataTable dataTable1;
        private DataTable dataTable2;
        private DataTable dataTable3;
        private int SelectedColumnIndex;
        private int differencesCount;
        private List<int> ConflictsRow;
        private List<int> ConflictsColumn;
        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            dataTable1 = new DataTable();
            dataTable2 = new DataTable();
            dataTable3 = new DataTable();
            ConflictsRow = new List<int>();
            ConflictsColumn = new List<int>();
        }
        public string filePath1;
        public string filePath2;
        public string newFilePath;
        public int position = 0;

        private void MainWindow_Loaded(object sender, RoutedEventArgs e) 
        {
            Button_Previos_Conflict.IsEnabled = false;
            Button_Next_Conflict.IsEnabled = false;
            Button_First_File.IsEnabled = false;
            Button_Second_File.IsEnabled = false;
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

                        var workbook = package.Workbook;
                        if (workbook.Worksheets.Count >0) 
                        {
                            var worksheet = workbook.Worksheets[1];
                            DataGrid1.Columns.Clear();
                            DataGrid1.Items.Clear();

                            dataTable1 = new DataTable();

                            //Заполняем загаловки столбцов
                            foreach (var cell  in worksheet.Cells[1,1,1, worksheet.Dimension.Columns]) 
                            {
                                dataTable1.Columns.Add(cell.Value?.ToString());
                            }

                            //Заполняем данные строк
                            for (int row = 2; row <= worksheet.Dimension.Rows; row++) 
                            {
                                var dataRow = dataTable1.NewRow();
                                for (int col = 1; col <= worksheet.Dimension.Columns; col++) 
                                {
                                    dataRow[col - 1] = worksheet.Cells[row, col].Value;
                                }
                                dataTable1.Rows.Add(dataRow);
                            }
                            if (string.IsNullOrEmpty(dataTable1.Columns[dataTable1.Columns.Count - 1].ColumnName)) 
                            {
                                dataTable1.Columns.RemoveAt(dataTable1.Columns.Count - 1);
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
            CompareExcelFiles(filePath1,filePath2, differencesCount,ConflictsRow,ConflictsColumn);
            DataGrid3.ItemsSource = dataTable1.DefaultView;
            DataGrid3.Items.Refresh();

        }

        private void SelectFromFirstFileButton_Click(object sender, RoutedEventArgs e) 
        {
            if (DataGrid3.SelectedItem is DataRowView selectedRow) 
            {
                int rowIndex = dataTable3.Rows.IndexOf(selectedRow.Row);

                string valueFromFirstFile = dataTable1.Rows[rowIndex][SelectedColumnIndex]?.ToString();

                dataTable3.Rows[rowIndex][SelectedColumnIndex] = valueFromFirstFile;

                if (!dataTable3.Rows.Contains(selectedRow.Row)) 
                {
                    dataTable3.ImportRow(selectedRow.Row);
                }
                SaveUserChoice(filePath1);
                DataGrid3.ItemsSource = dataTable3.DefaultView;
                DataGrid3.Items.Refresh();
            }
        }

        private void SelectFromSecondFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid3.SelectedItem is DataRowView selectedRow)
            {
                int rowIndex = dataTable3.Rows.IndexOf(selectedRow.Row);

                string valueFromSecondFile = dataTable2.Rows[rowIndex][SelectedColumnIndex]?.ToString();

                dataTable3.Rows[rowIndex][SelectedColumnIndex] = valueFromSecondFile;

                if (!dataTable3.Rows.Contains(selectedRow.Row))
                {
                    dataTable3.ImportRow(selectedRow.Row);
                }

                SaveUserChoice(filePath2);

                DataGrid3.ItemsSource = dataTable3.DefaultView;
                DataGrid3.Items.Refresh();
            }
        }

        private void ShowValueNextConflictButton_Click(object sender, RoutedEventArgs e) 
        {
            
            if (position >= differencesCount)
            {
                IsEnabled = false;
            }
            else 
            {
                position++;
            }

            DisplayConflict(position);
        }

        private void ShowValuePreviousConflictButton_Click(object sender, RoutedEventArgs e)
        {

            if (position <= differencesCount)
            {
                IsEnabled = false;
            }
            else 
            {
                position--;
            }
            
            DisplayConflict(position);
        }

        private void DisplayConflict(int conflictIndex) 
        {
            //Очищаем выбранные ячейки в третьем DataGrid
            DataGrid3.UnselectAllCells();

            if (DataGrid3.SelectedItem != null) 
            {
                int positionRowConflict = ConflictsRow[conflictIndex];
                int positionColumnConflict = ConflictsColumn[conflictIndex];
                //Выделяем текущий конфликт
                DataGrid3.SelectedIndex = positionRowConflict;
                DataGrid3.SelectedItem = positionColumnConflict;
                DataGrid3.ScrollIntoView(DataGrid3.SelectedItem);

                //Обновляем индекс текущей выбранной колонки
                SelectedColumnIndex = DataGrid3.CurrentCell.Column.DisplayIndex;
            }
        }

        private void SaveUserChoice(string selectedFilePath) 
        {
            FileInfo selectedFile = new FileInfo(selectedFilePath);

            using (ExcelPackage selectedPackage = new ExcelPackage(selectedFile)) 
            {
                ExcelWorksheet selectedWorksheet = selectedPackage.Workbook.Worksheets[1];

                for (int row = 0; row <= dataTable3.Rows.Count; row++) 
                {
                    DataRow dataRow = dataTable3.Rows[row];

                    int rowIndex = row + 1;

                    for (int col = 0; col <= dataTable3.Columns.Count; col++) 
                    {
                        DataColumn dataColumn = dataTable3.Columns[col];

                        string cellValue = selectedWorksheet.Cells[rowIndex, col + 1].GetValue<string>()?.Trim();

                        dataRow[dataColumn] = cellValue;
                    }
                }
                DataGrid3.ItemsSource = dataTable3.DefaultView;
                DataGrid3.Items.Refresh();
            }
        }

        private void SaveNewFileButton_Click(object sender, RoutedEventArgs e) 
        {
            if (!string.IsNullOrEmpty(newFilePath))
            {
                FileInfo newFile = new FileInfo(newFilePath);
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
                            worksheet.Cells[row + 1, col].Value = dataTable3.Rows[row - 1][col - 1]?.ToString();
                        }
                    }
                    package.Save();
                    MessageBox.Show("Новый файл сохранен.");
                }
            }
            else 
            {
                MessageBox.Show("Новый файл не найден.");
            }
        }

        public static void CompareExcelFiles(string filePath1, string filePath2,int differencesCount,List<int> conflictRow, List<int> conflictColumn)
        {
            //conflictColumn.Clear();
            //conflictRow.Clear();

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
                            differences.Add($"Различие в ячейке {cellAdress}: {cellValue1} != {cellValue2}");

                            differencesCount++;

                            conflictRow.Add(row);
                            conflictColumn.Add(col);
                        }
                    }
                }

                HighlightDifferences(worksheet1, differences);
                HighlightDifferences(worksheet2, differences);

                package1.Save();
                package2.Save();

                MessageBox.Show($"Сравнение завершено, количество различий {differencesCount}");
            }
        }

        public static void Main() 
        {

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
            string columnLetter = Regex.Replace(address, @"[\d]", string.Empty);
            int columnNumber = 0;

            foreach (char c in columnLetter)
            {
                columnNumber *= 26;
                columnNumber += char.ToUpper(c) - 'A' + 1;
            }
            return columnNumber;
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

        private static void HighlightDifferences(ExcelWorksheet worksheet, List<string> differences)
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

                int columnIndex = columnNumber - 1;

                if (columnIndex < 0 || columnIndex >= worksheet.Dimension.Columns)
                {
                    continue;
                }

                ExcelRange cell = worksheet.Cells[rowIndex, columnIndex + 1];

                var mergedCells = worksheet.MergedCells;
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
                }
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(highlightedFill);
            }
        }
    }
}
/*private static void HighlightDifferences(ExcelWorksheet worksheet, List<string> differences)
        {
            System.Windows.Media.Color highlightedColor = Colors.Red;

            foreach (var address in differences)
            {
                int columnNumber = GetColumnNumber(address);
                string rowNumber = Regex.Replace(address, @"[^\d]+", string.Empty);

                if (!int.TryParse(rowNumber, out int rowIndex))
                {
                    continue;
                }

                int columnIndex = columnNumber - 1;

                if (columnIndex < 0 || columnIndex >= worksheet.Dimension.Columns)
                {
                    continue;
                }

                // Изменяем цвет ячейки в DataGrid1 или DataGrid2
                var dataGridCell1 = GetCellFromDataGrid(DataGrid1, rowIndex - 1, columnIndex);
                dataGridCell1.Background = new SolidColorBrush(highlightedColor);

                var dataGridCell2 = GetCellFromDataGrid(DataGrid2, rowIndex - 1, columnIndex);
                dataGridCell2.Background = new SolidColorBrush(highlightedColor);

                ExcelRange cell = worksheet.Cells[rowIndex, columnIndex + 1];

                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(255, 255, 0, 0);
            }

            // Обновляем DataGrid1 и DataGrid2
            DataGrid1.Items.Refresh();
            DataGrid2.Items.Refresh();
        }*/