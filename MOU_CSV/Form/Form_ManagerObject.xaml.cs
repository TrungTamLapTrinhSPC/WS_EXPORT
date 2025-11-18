using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MOU_CSV
{
    public partial class Form_ManagerObject : Window
    {
        private Document _doc;
        private UIDocument _uidoc;
        private ObservableCollection<CategoryItem> _categories;
        private ObservableCollection<CategoryItem> _filteredCategories;
        private ObservableCollection<dynamic> _elements;
        private List<dynamic> _allElements;
        private List<Dictionary<string, string>> _sourceData;
        private List<Dictionary<string, string>> _updateData;
        private DataTable _elementsTable;
        private bool _isUpdatingSelection = false;
        private CategoryItem _selectedCategory = null;
        private string _currentExportMode = "single";

        // Biến để lưu thông tin thay đổi
        private List<ChangeInfo> _changes;
        private List<string> _changedCategories;
        private List<string> _changedObjectIds;
        private List<string> _changedParameters;

        public Form_ManagerObject(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            TranslateUI();

            _doc = doc;
            _uidoc = uidoc;
            _categories = new ObservableCollection<CategoryItem>();
            _filteredCategories = new ObservableCollection<CategoryItem>();

            // Gộp dữ liệu từ các biến riêng lẻ
            MergeAllData();

            _elementsTable = new DataTable();

            lstCategories.ItemsSource = _filteredCategories;
            dgElements.ItemsSource = _elementsTable.DefaultView;

            LoadCategories();
            AttachEventHandlers();

            InitializeDataGridEvents();
        }

        public void TranslateUI()
        {
            _lb1.Text = SimpleTranslator.T("OBJECT LIST");
            _lb2.Text = SimpleTranslator.T("Select a category to view");
            _lb3.Text = "🔍 " + SimpleTranslator.T("Search") + "...";
            btnRefresh.Content = "🔄 " + SimpleTranslator.T("REFRESH");
            lblSelectedCategory.Text = SimpleTranslator.T("Select a category to view information");
            _lb4.Text = SimpleTranslator.T("Filter by") + ":";
            _cbb1.Content = SimpleTranslator.T("All");
            _cbb2.Content = SimpleTranslator.T("By Level");
            _cbb3.Content = SimpleTranslator.T("By Type");
            btnImport.Content = "📥 " + SimpleTranslator.T("IMPORT CSV");
            btnExportMain.Content = "📤 " + SimpleTranslator.T("Export Current Category CSV");
            lblSelectedCount.Text = SimpleTranslator.T("No objects selected");
            btnHighlight.Content = "👁 " + SimpleTranslator.T("HIGHLIGHT");
            btnApply.Content = "✓ " + SimpleTranslator.T("Apply Changes");
            _btn1.Content = "📤 " + SimpleTranslator.T("Export Current Category CSV");
            _btn2.Content = "📤 " + SimpleTranslator.T("Export Multiple Categories CSV");

        }

        private void AttachEventHandlers()
        {
            // TextBox
            txtSearch.TextChanged -= TxtSearch_TextChanged;
            txtSearch.TextChanged += TxtSearch_TextChanged;

            // ComboBox
            cboFilter.SelectionChanged -= CboFilter_SelectionChanged;
            cboFilter.SelectionChanged += CboFilter_SelectionChanged;

            // DataGrid
            dgElements.SelectionChanged -= DgElements_SelectionChanged;
            dgElements.SelectionChanged += DgElements_SelectionChanged;

            dgElements.MouseDoubleClick -= DgElements_MouseDoubleClick;
            dgElements.MouseDoubleClick += DgElements_MouseDoubleClick;

            // Buttons
            btnRefresh.Click -= BtnRefresh_Click;
            btnRefresh.Click += BtnRefresh_Click;
        }

        private void LoadCategories()
        {
            _categories.Clear();
            _filteredCategories.Clear();

            if (_sourceData == null || _sourceData.Count == 0)
            {
                MessageBox.Show(T("No data available. Please run the Export command first")+".");
                return;
            }

            var categoryGroups = _sourceData
                .GroupBy(row => row.ContainsKey(T("Category")) ? row[T("Category")] : T("Unknown"))
                .Select(g => new
                {
                    CategoryName = g.Key,
                    Count = g.Count()
                })
                .OrderBy(x => x.CategoryName)
                .ToList();

            foreach (var group in categoryGroups)
            {
                var categoryItem = new CategoryItem
                {
                    DisplayName = group.CategoryName,
                    Count = group.Count,
                    CountText = $"({group.Count})"
                };

                _categories.Add(categoryItem);
                _filteredCategories.Add(categoryItem);
            }

            MessageBox.Show(T("Loaded") + " "+ _categories.Count +" " +T("categories"));
        }

        private void CategoryButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as System.Windows.Controls.Button;
            if (button == null) return;

            var clickedCategory = button.Tag as CategoryItem;
            if (clickedCategory == null) return;

            // Deselect all categories first
            foreach (var cat in _filteredCategories)
            {
                cat.IsSelected = false;
            }

            // Update selected state
            if (_selectedCategory == clickedCategory)
            {
                // If clicking the currently selected category, deselect it
                _selectedCategory = null;
            }
            else
            {
                // Select new category
                _selectedCategory = clickedCategory;
                clickedCategory.IsSelected = true;
            }

            // Load data
            LoadSelectedCategories();
        }

        private void LoadSelectedCategories()
        {
            try
            {
                _elementsTable.Clear();
                _elementsTable.Columns.Clear();
                dgElements.ItemsSource = null;

                if (_selectedCategory == null)
                {
                    lblSelectedCategory.Text = T("No category selected");
                    lblElementCount.Text = "";
                    dgElements.ItemsSource = _elementsTable.DefaultView;
                    return;
                }

                lblSelectedCategory.Text = _selectedCategory.DisplayName;

                int totalCount = 0;

                // Bước 1: Thu thập tất cả các keys theo thứ tự xuất hiện
                List<string> orderedKeys = new List<string>();
                List<Dictionary<string, string>> categoryRows = new List<Dictionary<string, string>>();

                foreach (var row in _sourceData)
                {
                    string categoryName = row.ContainsKey("Category") ? row["Category"] : "Unknown";

                    if (categoryName != _selectedCategory.DisplayName)
                        continue;

                    categoryRows.Add(row);

                    // Thu thập keys theo thứ tự xuất hiện, không trùng lặp
                    foreach (var key in row.Keys)
                    {
                        if (!orderedKeys.Contains(key))
                        {
                            orderedKeys.Add(key);
                        }
                    }
                }

                if (categoryRows.Count == 0)
                {
                    dgElements.ItemsSource = _elementsTable.DefaultView;
                    lblElementCount.Text = T("No elements found");
                    return;
                }

                // Bước 2: Tạo columns theo thứ tự đã thu thập
                _elementsTable.Columns.Add("_ElementId", typeof(int));

                foreach (var key in orderedKeys)
                {
                    if (!_elementsTable.Columns.Contains(key))
                    {
                        _elementsTable.Columns.Add(key, typeof(string));
                    }
                }

                // Bước 3: Thêm dữ liệu
                foreach (var row in categoryRows)
                {
                    DataRow dataRow = _elementsTable.NewRow();

                    string revitIdStr = row.ContainsKey("Revit ID") ? row["Revit ID"] : "0";
                    int.TryParse(revitIdStr, out int revitId);
                    dataRow["_ElementId"] = revitId;

                    foreach (var kvp in row)
                    {
                        // Chỉ gán nếu column tồn tại
                        if (_elementsTable.Columns.Contains(kvp.Key))
                        {
                            dataRow[kvp.Key] = kvp.Value ?? "";
                        }
                    }

                    _elementsTable.Rows.Add(dataRow);
                    totalCount++;
                }

                dgElements.ItemsSource = _elementsTable.DefaultView;

                // Ẩn column _ElementId
                if (dgElements.Columns.Count > 0)
                {
                    var hiddenColumn = dgElements.Columns.FirstOrDefault(c => c.Header.ToString() == "_ElementId");
                    if (hiddenColumn != null)
                    {
                        hiddenColumn.Visibility = System.Windows.Visibility.Collapsed;
                    }
                }

                // Đếm số types
                int distinctTypes = 0;
                if (_elementsTable.Columns.Contains("Type Name"))
                {
                    distinctTypes = _elementsTable.AsEnumerable()
                        .Select(r => r.Field<string>("Type Name"))
                        .Where(t => !string.IsNullOrEmpty(t))
                        .Distinct()
                        .Count();
                }

                lblElementCount.Text = T("Total") + ": " + totalCount + " " + T("objects") + "| " + distinctTypes + " " + T("types");

                UpdateSelectedCount();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}\n\n{ex.StackTrace}");
            }
        }
        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string searchText = txtSearch.Text.ToLower();
            _filteredCategories.Clear();

            foreach (var category in _categories)
            {
                if (category.DisplayName.ToLower().Contains(searchText))
                {
                    _filteredCategories.Add(category);
                }
            }
        }

        private void CboFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboFilter.SelectedIndex == -1 || _elementsTable.Rows.Count == 0)
                return;

            var selectedFilter = (cboFilter.SelectedItem as ComboBoxItem)?.Content.ToString();

            DataView dv = _elementsTable.DefaultView;

            if (selectedFilter == T("All"))
            {
                dv.Sort = "";
            }
            else if (selectedFilter == T("By Level"))
            {
                if (_elementsTable.Columns.Contains("Level"))
                    dv.Sort = "Level ASC";
            }
            else if (selectedFilter == T("By Type"))
            {
                if (_elementsTable.Columns.Contains("Type Name"))
                    dv.Sort = "[Type Name] ASC";
            }
        }

        private void DgElements_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSelectedCount();
        }

        private void DgElements_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            BtnHighlight_Click(sender, e);
        }

        private void UpdateSelectedCount()
        {
            int selectedCount = dgElements.SelectedItems.Count;
            if (selectedCount == 0)
            {
                lblSelectedCount.Text = T("No objects selected");
            }
            else if (selectedCount == 1)
            {
                lblSelectedCount.Text = "1 " + T("object selected");
            }
            else
            {
                lblSelectedCount.Text = selectedCount + " " + T("objects selected");
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            MergeAllData();
            _selectedCategory = null;
            LoadCategories();
            _elementsTable.Clear();
            lblSelectedCategory.Text = T("No category selected");
            lblElementCount.Text = "";
            lblSelectedCount.Text = T("No objects selected");
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = T("Select CSV File(s)"),
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    Multiselect = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _updateData = new List<Dictionary<string, string>>();

                    foreach (string filePath in openFileDialog.FileNames)
                    {
                        var lines = File.ReadAllLines(filePath, Encoding.UTF8); // Thêm UTF8 encoding

                        if (lines.Length == 0)
                        {
                            MessageBox.Show(T("File") + " " + Path.GetFileName(filePath) + " "+ T("is empty") + ".",
                                T("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                            continue;
                        }

                        // Dòng đầu tiên là header
                        string[] headers = ParseCsvLine(lines[0]);

                        // Xử lý các dòng dữ liệu (bỏ qua dòng header)
                        for (int i = 1; i < lines.Length; i++)
                        {
                            // Bỏ qua dòng trống
                            if (string.IsNullOrWhiteSpace(lines[i]))
                                continue;

                            string[] values = ParseCsvLine(lines[i]);

                            Dictionary<string, string> row = new Dictionary<string, string>();

                            for (int j = 0; j < headers.Length && j < values.Length; j++)
                            {
                                string key = headers[j].Trim();
                                string value = values[j].Trim();

                                row[key] = value;
                            }

                            // Chỉ thêm row nếu có ít nhất 1 giá trị không rỗng
                            if (row.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                            {
                                _updateData.Add(row);
                            }
                        }
                    }

                    MessageBox.Show(T("Successfully imported") +" " +_updateData.Count + " "+T("rows from") + " "+ openFileDialog.FileNames.Length+ " "+T("file(s)")+".",
                        T("Import Successful"), MessageBoxButton.OK, MessageBoxImage.Information);

                    CheckDataUpdate();
                    ApplyUpdates();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing CSV file(s): {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Hàm helper để parse CSV line với delimiter ;
        private string[] ParseCsvLine(string line)
        {
            List<string> result = new List<string>();
            bool inQuotes = false;
            StringBuilder currentValue = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    // Kiểm tra double quotes ""
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentValue.Append('"');
                        i++; // Skip next quote
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (!inQuotes && (c == ',' || c == ';')) // <== Nhận diện cả , và ;
                {
                    result.Add(currentValue.ToString());
                    currentValue.Clear();
                }
                else
                {
                    currentValue.Append(c);
                }
            }

            // Thêm giá trị cuối cùng
            result.Add(currentValue.ToString());

            return result.ToArray();
        }


        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            // Call based on current mode
            if (_currentExportMode == "single")
            {
                BtnExportSingle_Click(sender, e);
            }
            else
            {
                BtnExportMulti_Click(sender, e);
            }
        }

        private void BtnExportDropdown_Click(object sender, RoutedEventArgs e)
        {
            popupExport.IsOpen = !popupExport.IsOpen;
        }

        private void BtnExportSingle_Click(object sender, RoutedEventArgs e)
        {
            popupExport.IsOpen = false;
            _currentExportMode = "single";
            btnExportMain.Content = "📤 " + T("Export Current Category CSV");
            btnExportDropdown.Margin = new Thickness(205, 0, 0, 0);


            if (_elementsTable == null || _elementsTable.Rows.Count == 0)
            {
                MessageBox.Show(T("No data to export."));
                return;
            }
            string categoryName = _selectedCategory.DisplayName;

            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = $"Data_{categoryName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (StreamWriter sw = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                    {
                        // Write header
                        List<string> headers = new List<string>();
                        foreach (DataColumn col in _elementsTable.Columns)
                        {
                            if (col.ColumnName == "_ElementId") continue;
                            headers.Add(EscapeCsvValue(col.ColumnName));
                        }
                        sw.WriteLine(string.Join(";", headers));

                        // Write data
                        foreach (DataRow row in _elementsTable.Rows)
                        {
                            List<string> values = new List<string>();
                            foreach (DataColumn col in _elementsTable.Columns)
                            {
                                if (col.ColumnName == "_ElementId") continue;
                                string v = row[col]?.ToString() ?? "";
                                values.Add(EscapeCsvValue(v));
                            }
                            sw.WriteLine(string.Join(";", values));
                        }
                    }

                    MessageBox.Show(T("Data exported successfully!"), T("Notification"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting data: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnExportMulti_Click(object sender, RoutedEventArgs e)
        {
            Form_Export_Multi exportForm = new Form_Export_Multi(_sourceData);
            exportForm.ShowDialog();
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

        private void BtnHighlight_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<ElementId> listIds = new List<ElementId>();

                foreach (var row in dgElements.SelectedItems)
                {
                    object idValue = null;

                    if (row is DataRowView drv)
                    {
                        idValue = drv.Row["_ElementId"];
                    }

                    if (idValue == null) continue;

                    if (int.TryParse(idValue.ToString(), out int idInt))
                    {
                        listIds.Add(new ElementId(idInt));
                    }
                }

                if (!listIds.Any())
                {
                    MessageBox.Show(T("Please select at least one object."));
                    return;
                }

                _uidoc.Selection.SetElementIds(listIds);
                _uidoc.ShowElements(listIds);

                MessageBox.Show(T("Highlighted") +" "+ listIds.Count +" "+ T("objects in Revit")+".");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            if (_changes == null || _changes.Count == 0)
            {
                MessageBox.Show(T("No changes to apply")+".", T("Information"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show(
                T("Are you sure you want to apply") + " " + _changes.Count + " " + T("changes")+"?\n\n" +
                T("This will update parameters in Revit after closing this form")+".",
                T("Confirm Apply Changes"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            // Chuyển đổi _changes thành format Dictionary cho cls_Bien.dataImport
            cls_Bien.dataImport = ConvertChangesToDictionary(_changes);

            // Lưu danh sách element IDs đã thay đổi
            cls_Bien.changedElementIds = _changes
                .Select(c => c.RevitId)
                .Distinct()
                .ToList();

            MessageBox.Show(
                T("Prepared") + " " + _changes.Count + T("changes for") + " " + cls_Bien.changedElementIds.Count + " " + T("objects") + ".\n\n" +
                T("Changes will be applied after closing this form") + ".",
                T("Ready to Apply"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            // Đóng form
            this.Close();
        }

        /// <summary>
        /// Chuyển đổi List<ChangeInfo> thành List<Dictionary<string, string>>
        /// </summary>
        private List<Dictionary<string, string>> ConvertChangesToDictionary(List<ChangeInfo> changes)
        {
            var result = new List<Dictionary<string, string>>();

            // Group changes theo Revit ID
            var groupedByElement = changes.GroupBy(c => c.RevitId);

            foreach (var group in groupedByElement)
            {
                var dict = new Dictionary<string, string>();

                // Thêm Revit ID
                dict["Revit ID"] = group.Key;

                // Thêm Category (lấy từ change đầu tiên trong group)
                dict["Category"] = group.First().Category;

                // Thêm tất cả parameters cần thay đổi
                foreach (var change in group)
                {
                    dict[change.ParameterName] = change.NewValue;
                }

                result.Add(dict);
            }

            return result;
        }


        private void MergeAllData()
        {
            _sourceData = new List<Dictionary<string, string>>();

            // Kiểm tra dữ liệu cơ bản
            if (cls_Bien._basicInfo == null || cls_Bien._basicInfo.Count == 0)
            {
                MessageBox.Show(T("No basic information available")+".");
                return;
            }

            for (int i = 0; i < cls_Bien._basicInfo.Count; i++)
            {
                Dictionary<string, string> mergedRow = new Dictionary<string, string>();

                // 1. Add basic info first (Level, Category, Revit ID, Type Name)
                if (cls_Bien._basicInfo != null && i < cls_Bien._basicInfo.Count)
                {
                    foreach (var kvp in cls_Bien._basicInfo[i])
                    {
                        mergedRow[kvp.Key] = kvp.Value;
                    }
                }

                // 2. Add constraints
                if (cls_Bien._constraintsInfo != null && i < cls_Bien._constraintsInfo.Count)
                {
                    foreach (var kvp in cls_Bien._constraintsInfo[i])
                    {
                        if (!mergedRow.ContainsKey(kvp.Key))
                            mergedRow[kvp.Key] = kvp.Value;
                    }
                }

                // 3. Add dimensions
                if (cls_Bien._dimensionsInfo != null && i < cls_Bien._dimensionsInfo.Count)
                {
                    foreach (var kvp in cls_Bien._dimensionsInfo[i])
                    {
                        if (!mergedRow.ContainsKey(kvp.Key))
                            mergedRow[kvp.Key] = kvp.Value;
                    }
                }

                // 4. Add identity data last
                if (cls_Bien._identityInfo != null && i < cls_Bien._identityInfo.Count)
                {
                    foreach (var kvp in cls_Bien._identityInfo[i])
                    {
                        if (!mergedRow.ContainsKey(kvp.Key))
                            mergedRow[kvp.Key] = kvp.Value;
                    }
                }

                _sourceData.Add(mergedRow);
            }
        }

        // Biến để lưu thông tin thay đổi

        public void CheckDataUpdate()
        {
            _changes = new List<ChangeInfo>();

            if (_sourceData == null || _updateData == null)
            {
                return;
            }

            foreach (var updateRow in _updateData)
            {
                string revitId = updateRow.ContainsKey("Revit ID") ? updateRow["Revit ID"] : "";

                if (string.IsNullOrEmpty(revitId))
                    continue;

                var sourceRow = _sourceData.FirstOrDefault(r =>
                    r.ContainsKey("Revit ID") && r["Revit ID"] == revitId);

                if (sourceRow == null)
                    continue;

                string category = sourceRow.ContainsKey(T("Category")) ? sourceRow[T("Category")] : "";

                foreach (var updateColumn in updateRow)
                {
                    string parameterName = updateColumn.Key;
                    string newValue = updateColumn.Value;

                    if (parameterName == "Revit ID" || parameterName == T("Category"))
                        continue;

                    string oldValue = sourceRow.ContainsKey(parameterName) ? sourceRow[parameterName] : "";

                    // ===== THÊM PHẦN NÀY ĐỂ CHUẨN HÓA =====
                    // Chuẩn hóa: Trim và coi null/empty như nhau
                    string normalizedOldValue = string.IsNullOrWhiteSpace(oldValue) ? "" : oldValue.Trim();
                    string normalizedNewValue = string.IsNullOrWhiteSpace(newValue) ? "" : newValue.Trim();

                    // So sánh sau khi chuẩn hóa
                    if (normalizedOldValue != normalizedNewValue)
                    // ===== KẾT THÚC PHẦN THÊM =====
                    {
                        _changes.Add(new ChangeInfo
                        {
                            Category = category,
                            RevitId = revitId,
                            ParameterName = parameterName,
                            OldValue = oldValue,  // Giữ nguyên giá trị gốc để hiển thị
                            NewValue = newValue
                        });
                    }
                }
            }

            _changedCategories = _changes
                .Select(c => c.Category)
                .Distinct()
                .ToList();

            _changedObjectIds = _changes
                .Select(c => c.RevitId)
                .Distinct()
                .ToList();

            _changedParameters = _changes
                .Select(c => c.ParameterName)
                .Distinct()
                .ToList();
        }
        private void ApplyUpdates()
        {
            if (_changes == null || _changes.Count == 0)
            {
                MessageBox.Show(T("No changes detected") + ".", T("Information"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Hiển thị thông tin tổng quan về thay đổi
            string summary = T("Changes Summary") + " :\n" +
                            T("Categories affected") + ": " + _changedCategories.Count + "\n" +
                            T("Objects affected") + ": " + _changedObjectIds.Count + "\n" +
                            T("Parameters affected") + ": " + _changedParameters.Count + "\n" +
                            T("Total changes") + ": " + _changes.Count + "\n\n" +
                            T("Do you want to highlight these changes") + "?";

            var result = MessageBox.Show(summary, T("Apply Changes"), MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            // 1. Highlight categories trong list
            HighlightChangedCategories();

            // 2. Reload lại category hiện tại để highlight elements
            if (_selectedCategory != null)
            {
                LoadSelectedCategories();
            }

            MessageBox.Show(T("Changes have been highlighted successfully") + "!", T("Success"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void HighlightChangedCategories()
        {
            // Đổi màu nền các category có thay đổi
            foreach (var category in _categories)
            {
                if (_changedCategories.Contains(category.DisplayName))
                {
                    category.HasChanges = true;
                }
            }

            // Refresh lại filtered list
            foreach (var category in _filteredCategories)
            {
                if (_changedCategories.Contains(category.DisplayName))
                {
                    category.HasChanges = true;
                }
            }

            // Trigger refresh UI
            lstCategories.Items.Refresh();
        }

        // Thêm vào constructor
        private void InitializeDataGridEvents()
        {
            dgElements.LoadingRow -= DgElements_LoadingRow;
            dgElements.LoadingRow += DgElements_LoadingRow;
        }

        private void DgElements_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (_changes == null || _changes.Count == 0)
            {
                // Reset về màu mặc định
                e.Row.Background = Brushes.White;
                return;
            }

            var dataRow = e.Row.Item as DataRowView;
            if (dataRow == null)
                return;

            string revitId = dataRow.Row["_ElementId"]?.ToString();
            if (string.IsNullOrEmpty(revitId))
                return;

            var objectChanges = _changes.Where(c => c.RevitId == revitId).ToList();

            if (objectChanges.Any())
            {
                // Màu đỏ nhạt cho failed changes (sau khi Apply)
                e.Row.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 230, 230)); // #FFE6E6

                e.Row.Loaded += (s, args) =>
                {
                    HighlightFailedCellsInRow(e.Row, objectChanges);
                };
            }
            else
            {
                e.Row.Background = Brushes.White;
            }
        }

        /// <summary>
        /// Highlight cells thất bại với màu đỏ đậm hơn
        /// </summary>
        private void HighlightFailedCellsInRow(DataGridRow row, List<ChangeInfo> objectChanges)
        {
            try
            {
                foreach (var change in objectChanges)
                {
                    string paramName = change.ParameterName;

                    for (int colIndex = 0; colIndex < dgElements.Columns.Count; colIndex++)
                    {
                        var column = dgElements.Columns[colIndex];
                        if (column.Header?.ToString() == paramName)
                        {
                            var cell = GetCell(row, colIndex);
                            if (cell != null)
                            {
                                // Màu đỏ đậm hơn cho failed changes
                                cell.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 150, 150)); // #FF9696
                                cell.FontWeight = FontWeights.Bold;

                                // Thêm tooltip để hiển thị lý do fail
                                cell.ToolTip = "This parameter could not be updated (read-only or type parameter)";
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error highlighting failed cells: {ex.Message}");
            }
        }


        // Helper method để lấy DataGridCell
        private DataGridCell GetCell(DataGridRow row, int columnIndex)
        {
            if (row == null) return null;

            var presenter = GetVisualChild<DataGridCellsPresenter>(row);
            if (presenter == null)
            {
                // Force generate cells
                row.ApplyTemplate();
                presenter = GetVisualChild<DataGridCellsPresenter>(row);
            }

            if (presenter == null) return null;

            var cell = presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
            if (cell == null)
            {
                // Scroll into view to generate cell
                dgElements.ScrollIntoView(row.Item, dgElements.Columns[columnIndex]);
                dgElements.UpdateLayout();
                cell = presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
            }

            return cell;
        }

        // Helper method để tìm visual child
        private T GetVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T typedChild)
                    return typedChild;

                var childOfChild = GetVisualChild<T>(child);
                if (childOfChild != null)
                    return childOfChild;
            }
            return null;
        }

        // Update CategoryItem class để thêm property HasChanges
        public class CategoryItem : INotifyPropertyChanged
        {
            private bool _isSelected;
            private bool _hasChanges;
            private bool _hasFailedChanges;

            public string DisplayName { get; set; }
            public string CountText { get; set; }
            public int Count { get; set; }

            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    if (_isSelected != value)
                    {
                        _isSelected = value;
                        OnPropertyChanged(nameof(IsSelected));
                    }
                }
            }

            public bool HasChanges
            {
                get => _hasChanges;
                set
                {
                    if (_hasChanges != value)
                    {
                        _hasChanges = value;
                        OnPropertyChanged(nameof(HasChanges));
                    }
                }
            }

            public bool HasFailedChanges
            {
                get => _hasFailedChanges;
                set
                {
                    if (_hasFailedChanges != value)
                    {
                        _hasFailedChanges = value;
                        OnPropertyChanged(nameof(HasFailedChanges));
                    }
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        private string T(string key)
        {
            return SimpleTranslator.T(key);
        }

    }

    public class CategoryItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private bool _hasChanges;
        private bool _hasFailedChanges;

        public string DisplayName { get; set; }
        public string CountText { get; set; }
        public int Count { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public bool HasChanges
        {
            get => _hasChanges;
            set
            {
                if (_hasChanges != value)
                {
                    _hasChanges = value;
                    OnPropertyChanged(nameof(HasChanges));
                }
            }
        }

        public bool HasFailedChanges
        {
            get => _hasFailedChanges;
            set
            {
                if (_hasFailedChanges != value)
                {
                    _hasFailedChanges = value;
                    OnPropertyChanged(nameof(HasFailedChanges));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    public class ChangeInfo
    {
        public string Category { get; set; }
        public string RevitId { get; set; }
        public string ParameterName { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }

}