using Microsoft.Win32;
using System.Collections.Generic;
using System.Windows;
using System.Linq;
using System.ComponentModel;
using System.IO;

namespace ExceltoSQL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string[] extensions = { ".xlsx", ".xls", ".csv" };
        private SqlBuilder _sqlBuilder;
        private IEnumerable<string> _worksheets;
        private BackgroundWorker backgroundWorker;
        private string sql;

        public MainWindow()
        {
            InitializeComponent();
            backgroundWorker = ((BackgroundWorker)this.FindResource("backgroundWorker"));
            if (App.Args != null)
                LoadFile(App.Args[0]);
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            sql = _sqlBuilder.GetSql();            
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Spinner.Visibility = Visibility.Collapsed;
            btnSql.IsEnabled = true;
            Clipboard.SetText(sql);
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = $"Excel files (*{string.Join("; *", extensions)})|*{string.Join("; *", extensions)}";
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
            panelTableName.Visibility = Visibility.Collapsed;
            dgColumns.Visibility = Visibility.Collapsed;
            btnSql.Visibility = Visibility.Collapsed;
            var extension = Path.GetExtension(filePath).ToLower();
            if (!extensions.Contains(extension))
            {
                ShowMessage($"Can't open {extension} files");
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
            panelTableName.Visibility = Visibility.Visible;
            btnSql.Visibility = Visibility.Visible;
            this.SizeToContent = SizeToContent.Height;
            this.MaxHeight = 147 + _sqlBuilder.Columns.Count() * 19 + (worksheets.Count() > 1? 26 : 0);
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
            {
                btnSql.IsEnabled = false;
                Spinner.Visibility = Visibility.Visible;
                _sqlBuilder.TableName = txtTableName.Text;
                backgroundWorker.RunWorkerAsync();
            }
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
