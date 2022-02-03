using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace ExceltoSQL
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
    {
        private readonly string[] _extensions = { ".xlsx", ".xls", ".csv" };
        private SqlBuilder _sqlBuilder;
		private string _filePath;
        private IEnumerable<Worksheet> _worksheets;
        private BackgroundWorker _backgroundWorker;
		private DispatcherTimer _timer;
		private string _sql;

        public MainWindow()
        {
            InitializeComponent();
			_backgroundWorker = (BackgroundWorker)FindResource("backgroundWorker");
			_timer = (DispatcherTimer)FindResource("timer");
			_timer.Interval = TimeSpan.FromSeconds(0.3);

			if (App.Args != null)
			{
				_filePath = App.Args[0];
                LoadFile();
			}
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = $"Excel files (*{string.Join("; *", _extensions)})|*{string.Join("; *", _extensions)}";
            var result = openFileDialog.ShowDialog();
            if (result.Value)
			{
				_filePath = openFileDialog.FileName;
                LoadFile();
			}
        }

        private void DropPanel_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
			if (files != null)
			{
				_filePath = files[0];
				LoadFile();
			}
        }

		private void BtnOpen_MouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
		{
			var text = Clipboard.GetText();
			if (string.IsNullOrEmpty(text))
				return;

			var lines = text.Replace("\r", "").Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
			if(_extensions.Contains(Path.GetExtension(lines[0])))
			{
				_filePath = lines[0];
			}
			_worksheets = new List<Worksheet> { new Worksheet { Rows = GetRows(new List<string>{ "value" }.Concat(lines)) } };
			PopulateGrid();
			btnSql.Visibility = Visibility.Visible;
			Resize();
		}

		private IEnumerable<IEnumerable<string>> GetRows(IEnumerable<string> lines)
		{
			return lines.Select(s => new List<string> { s });
		}

		private void LoadFile()
        {
            PanelWorksheet.Visibility = Visibility.Collapsed;
            panelTableName.Visibility = Visibility.Collapsed;
            dgColumns.Visibility = Visibility.Collapsed;
            btnSql.Visibility = Visibility.Collapsed;
            var extension = Path.GetExtension(_filePath).ToLower();
            if (!_extensions.Contains(extension))
            {
                ShowMessage($"Can't open {extension} files");
                return;
            }
			Spinner.Visibility = Visibility.Visible;
			_backgroundWorker.RunWorkerAsync(new [] { "file" });
        }

		private void UpdateUi()
		{
			if (_worksheets == null)
			{
				ShowMessage($"Can't open {_filePath}");
				return;
			}
			UpdateWorksheetDropdown();
			PopulateGrid();
			btnSql.Visibility = Visibility.Visible;
			Resize();
		}

		private void UpdateWorksheetDropdown()
		{
			if (_worksheets.Count() > 1)
			{
				var worksheetNames = _worksheets.Select((w, i) => $" {i + 1}   {w.Title}");
				ComboWorksheet.ItemsSource = worksheetNames;
				ComboWorksheet.SelectedItem = worksheetNames.First();
				PanelWorksheet.Visibility = Visibility.Visible;
			}
		}

		private void PopulateGrid()
		{
			var worksheetIndex = _worksheets.Count() > 1 ? ComboWorksheet.SelectedIndex : 0;
			_sqlBuilder = new SqlBuilder(_worksheets);
			_sqlBuilder.OpenWorksheet(worksheetIndex, cbUnderscore.IsChecked ?? false, cbIgnoreEmpty.IsChecked ?? false);
			dgColumns.ItemsSource = _sqlBuilder.Columns;
			dgColumns.Visibility = Visibility.Visible;
			panelTableName.Visibility = Visibility.Visible;
			panelOptions.Visibility = Visibility.Visible;
		}

		public void ShowMessage(string message)
			=> MessageBox.Show(message, "Error!");

		private void Resize()
		{
			SizeToContent = SizeToContent.Height;
			MaxHeight = 194 + _sqlBuilder.Columns.Count() * 19 + (_worksheets.Count() > 1 ? 26 : 0);
		}

		private void CbAll_Checked(object sender, RoutedEventArgs e)
        {
            _sqlBuilder.Columns.ForEach(c => c.Include = true);
        }

        private void CbAll_Unchecked(object sender, RoutedEventArgs e)
        {
            _sqlBuilder.Columns.ForEach(c => c.Include = false);
        }

        private void ComboSheet_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
			PopulateGrid();
			Resize();
        }

        private void BtnSql_Click(object sender, RoutedEventArgs e)
        {
            if (_sqlBuilder.Columns.Any(c => c.Include))
            {
                btnSql.IsEnabled = false;
				progress.Value = 0;
				SizeToContent = SizeToContent.Manual;
				btnSql.Visibility = Visibility.Collapsed;
				progress.Visibility = Visibility.Visible;
				SizeToContent = SizeToContent.Height;
				_sqlBuilder.TableName = txtTableName.Text;
                _backgroundWorker.RunWorkerAsync(new [] { "sql", $"{cbUnderscore.IsChecked.Value}"});
            }
            else
                ShowMessage("You must include at least one column");
        }

		private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			var args = e.Argument as string[];

			if (args[0] == "sql")
			{
				_sql = _sqlBuilder.GetSql(sender as BackgroundWorker, args[1] == "True");
				e.Result = "sql";
			}
			else
			{
				_worksheets = FileReader.ReadFile(_filePath, ShowMessage);
				e.Result = "file";
			}
		}

		private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			if (e.Result as string == "sql")
			{
				Clipboard.SetText(_sql);
				_timer.Start();
			}
			else
			{
				Spinner.Visibility = Visibility.Collapsed;
				if (!_worksheets.Any())
					ShowMessage("No worksheets found.");

				else
					UpdateUi();
			}
		}

		private void progressChanged(object sender, ProgressChangedEventArgs e)
		{
			progress.Value = e.ProgressPercentage;
		}

		private void timer_Tick(object sender, EventArgs e)
		{
			SizeToContent = SizeToContent.Manual;
			progress.Visibility = Visibility.Collapsed;
			btnSql.Visibility = Visibility.Visible;
			SizeToContent = SizeToContent.Height;
			btnSql.IsEnabled = true;
		}

		private void CbUnderscore_Click(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrEmpty(_filePath))
				PopulateGrid();
		}

		private void CbIgnoreEmpty_Click(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrEmpty(_filePath))
			{
				PopulateGrid();
				Resize();
			}
		}
	}
}
