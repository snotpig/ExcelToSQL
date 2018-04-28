using Microsoft.Win32;
using System.Collections.Generic;
using System.Windows;
using System.Linq;

namespace ExceltoSQL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SqlBuilder _sqlBuilder;
        private IEnumerable<string> _worksheets;

        public MainWindow()
        {
            InitializeComponent();
            if(App.Args != null)
                LoadFile(App.Args[0]);
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx; *.csv)|*.xlsx;*.csv";
            var result = openFileDialog.ShowDialog();
            if (result.Value)
                LoadFile(openFileDialog.FileName);
        }

        private void DropPanel_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            LoadFile(files[0]);     
        }

        private void LoadFile(string filePath)
        {
            PanelWorksheet.Visibility = Visibility.Collapsed;
            dgColumns.Visibility = Visibility.Collapsed;
            btnSql.Visibility = Visibility.Collapsed;
            var extension = filePath.Substring(filePath.Length - 4).ToLower();
            if (extension != ".csv" && extension != "xlsx")
            {
                ShowMessage("Can only open .XLSX or .CSV files");
                return;
            }

            var worksheets = FileReader.ReadFile(filePath, ShowMessage);

            if (worksheets == null)
            {
                ShowMessage($"Can't open {filePath}");
                return;
            }
            _sqlBuilder = new SqlBuilder(worksheets);
            if (worksheets.Count() > 1)
            {
                _worksheets = worksheets.Select((w, i) => $" {i + 1}   {w.Title}");
                ComboWorksheet.ItemsSource = _worksheets;
                ComboWorksheet.SelectedItem = _worksheets.First();
                PanelWorksheet.Visibility = Visibility.Visible;
            }
            dgColumns.ItemsSource = _sqlBuilder.Columns;
            dgColumns.Visibility = Visibility.Visible;
            btnSql.Visibility = Visibility.Visible;
            this.SizeToContent = SizeToContent.Height;
            this.MaxHeight = 109.5 + _sqlBuilder.Columns.Count() * 19 + (worksheets.Count() > 1? 26 : 0);
        }

        public void ShowMessage(string message)
        {
            MessageBox.Show(message, "Error!");
        }

        private void CbAll_Checked(object sender, RoutedEventArgs e)
        {
            _sqlBuilder.Columns.ForEach(c => c.Include = true);
        }

        private void CbAll_Unchecked(object sender, RoutedEventArgs e)
        {
            _sqlBuilder.Columns.ForEach(c => c.Include = false);
        }

        private void BtnSql_Click(object sender, RoutedEventArgs e)
        {
            if (_sqlBuilder.Columns.Any(c => c.Include))
                Clipboard.SetText(_sqlBuilder.GetSql());
            else
                ShowMessage("You must include at least one column");
        }

        private void ComboSheet_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            _sqlBuilder.OpenWorksheet(ComboWorksheet.SelectedIndex + 1);
            dgColumns.ItemsSource = _sqlBuilder.Columns;
        }
    }
}
