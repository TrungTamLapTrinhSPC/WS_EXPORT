using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace MOU_CSV
{
    public partial class Form_Export_Multi : Window
    {
        public ObservableCollection<CategoryItem> Categories { get; set; }
        private List<Dictionary<string, string>> _sourceData;

        public Form_Export_Multi(List<Dictionary<string, string>> sourceData)
        {
            InitializeComponent();

            _sourceData = sourceData;

            // Tự động tạo categories từ sourceData
            LoadCategoriesFromSourceData();

            lstCategories.ItemsSource = Categories;

            btnExport.Click += BtnExport_Click;
            btnCancel.Click += BtnCancel_Click;
            lstCategories.PreviewMouseLeftButtonDown += LstCategories_PreviewMouseLeftButtonDown;



            _lb1.Text = T("Choose Categories");
            btnExport.Content = T("Export");
            btnCancel.Content = T("Cancel");


        }

        private void LoadCategoriesFromSourceData()
        {
            Categories = new ObservableCollection<CategoryItem>();

            if (_sourceData == null || _sourceData.Count == 0)
            {
                return;
            }

            // Nhóm theo Category và đếm số lượng
            var categoryGroups = _sourceData
                .GroupBy(row => row.ContainsKey(T("Category")) ? row[T("Category")] : T("Unknown"))
                .Select(g => new CategoryItem
                {
                    DisplayName = g.Key,
                    Count = g.Count(),
                    CountText = $"({g.Count()})",
                    IsSelected = false
                })
                .OrderBy(c => c.DisplayName)
                .ToList();

            foreach (var category in categoryGroups)
            {
                Categories.Add(category);
            }
        }

        private void LstCategories_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement(lstCategories, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (item != null && item.Content is CategoryItem category)
            {
                category.IsSelected = !category.IsSelected;
                e.Handled = true;
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            // Lấy các category được chọn
            var selectedCategories = Categories.Where(c => c.IsSelected).ToList();

            if (!selectedCategories.Any())
            {
                System.Windows.MessageBox.Show(T("Please select at least one category."), T("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Chọn folder
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = T("Select folder to save CSV files");
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string folderPath = dialog.SelectedPath;

                    try
                    {
                        int successCount = 0;
                        int errorCount = 0;
                        StringBuilder errorMessages = new StringBuilder();

                        // Xuất từng category
                        foreach (var category in selectedCategories)
                        {
                            try
                            {
                                ExportCategoryToCSV(category.DisplayName, folderPath);
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                errorCount++;
                                errorMessages.AppendLine($"- {category.DisplayName}: {ex.Message}");
                            }
                        }

                        // Hiển thị kết quả
                        string message = T("Export completed") + "!\n\n" +
                                       "✓ " + T("Success") + ": " + successCount + T("file(s)") + "\n\n" +
                                       "✗ " + T("Failed") + ": " + errorCount + T("file(s)");

                        if (errorCount > 0)
                        {
                            message += "\n\n" + T("Errors") + ": " + "\n" + errorMessages;
                            System.Windows.MessageBox.Show(message, T("Export Result"), MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        else
                        {
                            System.Windows.MessageBox.Show(message, T("Export Result"), MessageBoxButton.OK, MessageBoxImage.Information);
                        }

                        DialogResult = true;
                        Close();
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show(SimpleTranslator.T("Export error") + ": " + ex.Message, SimpleTranslator.T("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void ExportCategoryToCSV(string categoryName, string folderPath)
        {
            // Tạo DataTable cho category này
            DataTable elementsTable = new DataTable();
            bool columnsCreated = false;

            // Lọc dữ liệu theo category
            foreach (var row in _sourceData)
            {
                string rowCategory = row.ContainsKey(T("Category")) ? row[T("Category")] : T("Unknown");

                if (rowCategory != categoryName)
                    continue;

                // Tạo cột (chỉ lần đầu)
                if (!columnsCreated)
                {
                    foreach (var key in row.Keys)
                    {
                        if (!elementsTable.Columns.Contains(key))
                        {
                            elementsTable.Columns.Add(key, typeof(string));
                        }
                    }
                    columnsCreated = true;
                }

                // Thêm dữ liệu với try-catch để xử lý cột thiếu
                DataRow dataRow = elementsTable.NewRow();
                foreach (var kvp in row)
                {
                    try
                    {
                        // Kiểm tra xem cột có tồn tại không
                        if (elementsTable.Columns.Contains(kvp.Key))
                        {
                            dataRow[kvp.Key] = kvp.Value ?? "";
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }

                // Đảm bảo các cột không có trong row được gán giá trị rỗng
                foreach (DataColumn col in elementsTable.Columns)
                {
                    if (dataRow[col] == DBNull.Value || dataRow[col] == null)
                    {
                        dataRow[col] = "";
                    }
                }

                elementsTable.Rows.Add(dataRow);
            }

            // Nếu không có dữ liệu, bỏ qua
            if (elementsTable.Rows.Count == 0)
            {
                throw new Exception(T("No data found for this category"));
            }

            // Tạo tên file
            string fileName = $"Data_{categoryName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string filePath = Path.Combine(folderPath, fileName);

            // Xuất file CSV
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // Write header
                List<string> headers = new List<string>();
                foreach (DataColumn col in elementsTable.Columns)
                {
                    headers.Add(EscapeCsvValue(col.ColumnName));
                }
                sw.WriteLine(string.Join(";", headers));

                // Write data
                foreach (DataRow row in elementsTable.Rows)
                {
                    List<string> values = new List<string>();
                    foreach (DataColumn col in elementsTable.Columns)
                    {
                        string v = row[col]?.ToString() ?? "";
                        values.Add(EscapeCsvValue(v));
                    }
                    sw.WriteLine(string.Join(";", values));
                }
            }
        }

        private string EscapeCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            return value;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private string T(string key)
        {
            return SimpleTranslator.T(key);
        }

    }
}