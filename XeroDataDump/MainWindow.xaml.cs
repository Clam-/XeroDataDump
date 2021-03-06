﻿
using System.Windows;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading;
using Microsoft.Win32;
using System.IO;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace XeroDataDump
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>

	public partial class MainWindow : Window
	{
		BackgroundWorker worker;
		string oldOrgName = "";
		string oldConsKey = "";
		int oldYear = 0;
		int oldMonth = 0;

		// http://stackoverflow.com/a/15865842
		public static readonly DependencyProperty MonthsProperty = DependencyProperty.Register(
			"Months", typeof(List<string>), typeof(MainWindow),
			new PropertyMetadata(Thread.CurrentThread.CurrentCulture.DateTimeFormat.MonthNames.Take(12).ToList())
		);

		public List<string> Months
		{
			get { return (List<string>)GetValue(MonthsProperty); }
			set { SetValue(MonthsProperty, value); }
		}


		public MainWindow()
		{
			// setup UI
			InitializeComponent();
			ToDate.SelectedDate = DateTime.Today;
			initStoredValues();
			oldMonth = Month.SelectedIndex;
			oldOrgName = OrgName.Text;
			int.TryParse(Year.Text, out oldYear);
		}

		private void initStoredValues()
		{
			CostBudgetFname.Text = Options.Default.CostBudgetFile;
			TimeBudgetFname.Text = Options.Default.TimeBudgetFile;
			TimeSheetFname.Text = Options.Default.TimesheetFile;
			OrgName.Text = Options.Default.OrganisationName ?? OrgName.Text;
			Month.SelectedIndex = Options.Default.Month;
			Year.Text = Options.Default.Year.ToString();
			SaveName.Text = Options.Default.OutputFile;
			CertFname.Text = Options.Default.CertFile;
			ConsumerKey.Password = Options.Default.ConsumerKey;
			CertPass.Password = Options.Default.CertPassword;

			CostBudgetYearCol.Text = Options.Default.CostBudgetYearCol.ToString();
			TimeBudgetPosCol.Text = Options.Default.TimeBudgetPosCol.ToString();
			TimeBudgetYearCol.Text = Options.Default.TimeBudgetYearCol.ToString();

			IncAccts.Text = Options.Default.IncAccts.ToString();
			Positions.Text = Options.Default.Positions.ToString();
			Collation.Text = Options.Default.Collation.ToString();
			IgnoreSheets.Text = Options.Default.IgnoreSheets.ToString();
			OverheadProjs.Text = Options.Default.OverheadProjs.ToString();
			MergeAccounts.Text = Options.Default.MergeAccounts.ToString();
			HideAccounts.Text = Options.Default.HideAccounts.ToString();
			HoursDay.Text = Options.Default.HoursDay.ToString();
			MonthRows.Text = Options.Default.MonthRows.ToString();

			CostBudgetACHeader.Text = Options.Default.CostBudgetACHeader.ToString();
			CostBudgetAccCol.Text = Options.Default.CostBudgetAccCol.ToString();
			TBSearchCol.Text = Options.Default.TBSearchCol.ToString();
			TBEndRowText.Text = Options.Default.TBEndRowText.ToString();
			TBStartRowText.Text = Options.Default.TBStartRowText.ToString();

			LogoFname.Text = Options.Default.LogoFname.ToString();
			IgnoreBudgetSheets.Text = Options.Default.IgnoreBudgetSheets.ToString();
			NetProfitLabel.Text = Options.Default.NetProfitLabel.ToString();
			DeleteAcctRadio.IsChecked = Options.Default.DelEmptyAccounts;
			HideAcctRadio.IsChecked = !Options.Default.DelEmptyAccounts;
		}

		private void disableUI()
		{
			GetDataButton.IsEnabled = false;
			SaveName.IsEnabled = false;
			SaveBrowse.IsEnabled = false;
			BudgetBrowse.IsEnabled = false;
			OpenReport.IsEnabled = false;
		}
		private void enableUI()
		{
			GetDataButton.IsEnabled = true;
			SaveName.IsEnabled = true;
			SaveBrowse.IsEnabled = true;
			BudgetBrowse.IsEnabled = true;
			OpenReport.IsEnabled = true;
		}

		private void GetDataButton_Click(object sender, RoutedEventArgs e)
		{
			Log.Text = "";
			if (string.IsNullOrWhiteSpace(Options.Default.CostBudgetFile) || string.IsNullOrWhiteSpace(Options.Default.TimeBudgetFile) || string.IsNullOrWhiteSpace(Options.Default.OutputFile))
			{
				Log.Text = Log.Text + "Require Time Budget file, Cost Budget file and Output file.\n";
				return;
			}
			if (!File.Exists(TimeBudgetFname.Text))
			{
				Log.Text = Log.Text + "Budget Time File doesn't exist.\n";
				return;
			}
			List<object> args = new List<object> { ToDate.SelectedDate };

			// check year
			int year = -1;
			if (int.TryParse(Year.Text, out year)) { args.AddRange(new object[] { year, Month.SelectedIndex + 1 }); }
			else {
				Log.Text = Log.Text + "Year is not a number.\n";
				return;
			}

			// Check year to date is not in the future.
			if (ToDate.SelectedDate > DateTime.Now.AddDays(32))
			{
				Log.Text = Log.Text + "Year to date is set 32+ days in the future. You said you wouldn't do this.\n";
				return;
			}

			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;

			//worker.DoWork += new DoWorkEventHandler(Logic.CoreTest);
			// Actual logic.
			worker.DoWork += new DoWorkEventHandler(Logic.YTDDataDump);
				
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(worker_Done);
			worker.ProgressChanged += worker_ProgressChanged;
			disableUI();
			worker.RunWorkerAsync(args.ToArray());

		}

		private void TimeBudgetBrowse_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "Excel Workbook|*.xlsx";
			openDialog.Title = "Open Time Budget File...";
			openDialog.ShowDialog();

			// If the file name is not an empty string open it for saving.
			if (openDialog.FileName != "")
			{
				// get filename
				TimeBudgetFname.Text = openDialog.FileName;
				Options.Default.TimeBudgetFile = openDialog.FileName;
				Options.Default.Save();
			}
		}
		private void SaveBrowse_Click(object sender, RoutedEventArgs e)
		{
			// Displays a SaveFileDialog so the user can save the Image
			// assigned to Button2. https://msdn.microsoft.com/en-us/library/sfezx97z(v=vs.110).aspx
			//CommonOpenFileDialog cofd = new CommonOpenFileDialog();
			//cofd.IsFolderPicker = true;
			//cofd.Title = "Save to Excel file...";
			//cofd.ShowDialog();

			SaveFileDialog saveDialog = new SaveFileDialog();
			saveDialog.Filter = "Excel Workbook|*.xlsx";
			saveDialog.Title = "Save to location...";
			saveDialog.ShowDialog();

			// If the file name is not an empty string open it for saving.
			try
			{
				if (saveDialog.FileName != "")
				{
					// get filename
					SaveName.Text = saveDialog.FileName;
					Options.Default.OutputFile = saveDialog.FileName;
					Options.Default.Save();
				} }
			catch (InvalidOperationException) { }
		}

		void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			if (e.UserState != null)
				Log.Text += e.UserState as String;
			//progress.Value = e.ProgressPercentage;

		}

		void worker_Done(object sender, RunWorkerCompletedEventArgs e)
		{
			// First, handle the case where an exception was thrown. 
			if (e.Error != null)
			{
                Log.Text += e.Error.Message;
				Log.Text += "Worker Exception\r\n" + e.Error.GetType() + ":" + e.Error.Message + "\r\n" + e.Error.StackTrace;
			}
			else if (e.Cancelled)
			{
				// Next, handle the case where the user canceled  
				// the operation. 
				// Note that due to a race condition in  
				// the DoWork event handler, the Cancelled 
				// flag may not have been set, even though 
				// CancelAsync was called.
				Log.Text += "Operation canceled";
			}
			else
			{
				// Finally, handle the case where the operation  
				// succeeded.
				if (e.Result != null) { Log.Text += e.Result.ToString(); }
				else { Log.Text += "Done.\n"; }
			}

			enableUI();
        }

		private void OrgName_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(OrgName.Text) && OrgName.Text != oldOrgName)
			{
				Options.Default.OrganisationName = OrgName.Text;
				Options.Default.Save();
				oldOrgName = OrgName.Text;
			} else if (string.IsNullOrWhiteSpace(OrgName.Text))
			{
				OrgName.Text = oldOrgName;
			}
		}
		private void CertPass_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.CertPassword = CertPass.Password;
			Options.Default.Save();
		}
		private void ConsKey_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(ConsumerKey.Password) && ConsumerKey.Password != oldConsKey)
			{
				Options.Default.ConsumerKey = ConsumerKey.Password;
				Options.Default.Save();
				oldConsKey = ConsumerKey.Password;
			}
			else if (string.IsNullOrWhiteSpace(ConsumerKey.Password))
			{
				ConsumerKey.Password = oldConsKey;
			}
		}

		private void Year_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(Year.Text) && Year.Text != oldYear.ToString())
			{
				int year = 0;
				if (int.TryParse(Year.Text, out year)) {
					Options.Default.Year = year;
					Options.Default.Save();
					oldYear = year;
				} else {
					Log.Text = Log.Text + "Year is not a number.\n";
					return;
				}
			}
			else if (string.IsNullOrWhiteSpace(Year.Text))
			{
				Year.Text = oldYear.ToString();
			}
		}
		private void Month_LostFocus(object sender, RoutedEventArgs e)
		{
			if (Month.SelectedIndex != oldMonth)
			{
				Options.Default.Month = Month.SelectedIndex;
				Options.Default.Save();
				oldMonth = Month.SelectedIndex;
			}
		}

		private void CostBudgetBrowse_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "Excel Workbook|*.xlsx";
			openDialog.Title = "Open Cost Budget File...";
			openDialog.ShowDialog();

			// If the file name is not an empty string grab it
			if (openDialog.FileName != "")
			{
				// get filename
				CostBudgetFname.Text = openDialog.FileName;
				Options.Default.CostBudgetFile = openDialog.FileName;
				Options.Default.Save();
			}
		}

		private void CertBrowse_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "PFX Certificate|*.pfx";
			openDialog.Title = "Open Private Certificate...";
			openDialog.ShowDialog();

			// If the file name is not an empty string grab it
			if (openDialog.FileName != "")
			{
				// get filename
				CertFname.Text = openDialog.FileName;
				Options.Default.CertFile = openDialog.FileName;
				Options.Default.Save();
			}
		}

		private void CostBudgetYearCol_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(CostBudgetYearCol.Text))
			{
				int num = 1;
				if (int.TryParse(CostBudgetYearCol.Text, out num))
				{
					Options.Default.CostBudgetYearCol = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Cost Year Column is not a number.\n";
					return;
				}
			}
		}

		private void TimeBudgetPosCol_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(TimeBudgetPosCol.Text))
			{
				int num = 1;
				if (int.TryParse(TimeBudgetPosCol.Text, out num))
				{
					Options.Default.TimeBudgetPosCol = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Time Pos Column is not a number.\n";
					return;
				}
			}
		}

		private void TimeBudgetYearCol_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(TimeBudgetYearCol.Text))
			{
				int num = 1;
				if (int.TryParse(TimeBudgetYearCol.Text, out num))
				{
					Options.Default.TimeBudgetYearCol = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Time Year Column is not a number.\n";
					return;
				}
			}
		}

		private void IncAccts_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.IncAccts = IncAccts.Text;
			Options.Default.Save();
		}

		private void Positions_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.Positions = Positions.Text;
			Options.Default.Save();
		}

		private void TimesheetBrowse_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "Excel Workbook|*.xlsx";
			openDialog.Title = "Open Timesheet File...";
			openDialog.ShowDialog();

			// If the file name is not an empty string grab it
			if (openDialog.FileName != "")
			{
				// get filename
				TimeSheetFname.Text = openDialog.FileName;
				Options.Default.TimesheetFile = openDialog.FileName;
				Options.Default.Save();
			}
		}

		private void Collation_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.Collation = Collation.Text;
			Options.Default.Save();
		}

		private void IgnoreSheets_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.IgnoreSheets = IgnoreSheets.Text;
			Options.Default.Save();
		}

		private void OverheadProjs_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.OverheadProjs = OverheadProjs.Text;
			Options.Default.Save();
		}

		private void HoursDay_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(HoursDay.Text))
			{
				int num = 8;
				if (int.TryParse(HoursDay.Text, out num))
				{
					Options.Default.HoursDay = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Hours per day is not a number.\n";
					return;
				}
			}
		}

		private void MergeAccounts_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.MergeAccounts = MergeAccounts.Text;
			Options.Default.Save();
		}

		private void HideAccounts_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.HideAccounts = HideAccounts.Text;
			Options.Default.Save();
		}

		private void MonthRows_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(MonthRows.Text))
			{
				int num = 0;
				if (int.TryParse(MonthRows.Text, out num))
				{
					Options.Default.MonthRows = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Rows between Months is not a number.\n";
					return;
				}
			}
		}

		private void TBStartRowText_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.TBStartRowText = TBStartRowText.Text;
			Options.Default.Save();
		}

		private void TBEndRowText_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.TBEndRowText = TBEndRowText.Text;
			Options.Default.Save();
		}

		private void TBSearchCol_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(TBSearchCol.Text))
			{
				int num = 0;
				if (int.TryParse(TBSearchCol.Text, out num))
				{
					Options.Default.TBSearchCol = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Search Column is not a number.\n";
					return;
				}
			}
		}

		private void CostBudgetAccCol_LostFocus(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(CostBudgetAccCol.Text))
			{
				int num = 0;
				if (int.TryParse(CostBudgetAccCol.Text, out num))
				{
					Options.Default.CostBudgetAccCol = num;
					Options.Default.Save();
				}
				else {
					Log.Text = Log.Text + "Account Column is not a number.\n";
					return;
				}
			}
		}

		private void CostBudgetACHeader_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.CostBudgetACHeader = CostBudgetACHeader.Text;
			Options.Default.Save();
		}

		private void LogoFname_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.LogoFname = LogoFname.Text;
			Options.Default.Save();
		}

		private void LogoBrowse_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "Image|*.png;*.jpg;*.jpeg";
			openDialog.Title = "Open Logo File...";
			openDialog.ShowDialog();

			// If the file name is not an empty string grab it
			if (openDialog.FileName != "")
			{
				// get filename
				LogoFname.Text = openDialog.FileName;
				Options.Default.LogoFname = openDialog.FileName;
				Options.Default.Save();
			}
		}

		private void IgnoreBudgetSheets_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.IgnoreBudgetSheets = IgnoreBudgetSheets.Text;
			Options.Default.Save();
		}

		private void NetProfitLabel_LostFocus(object sender, RoutedEventArgs e)
		{
			Options.Default.NetProfitLabel = NetProfitLabel.Text;
			Options.Default.Save();
		}

		private void OpenReport_Click(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(Options.Default.OutputFile)) {
				Log.Text = Log.Text + "\nOpening report...";
				System.Diagnostics.Process.Start(Options.Default.OutputFile);
			}
		}

		private void SplitReport_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "Excel Workbook|*.xlsx";
			openDialog.Title = "Open Report to Split...";
			openDialog.ShowDialog();

			Log.Text = "";

			// If the file name is not an empty string grab it
			if (openDialog.FileName != "")
			{
				// do split

				List<object> args = new List<object> { openDialog.FileName };
				
				worker = new BackgroundWorker();
				worker.WorkerSupportsCancellation = false;
				worker.WorkerReportsProgress = true;

				//worker.DoWork += new DoWorkEventHandler(Logic.CoreTest);
				// Actual logic.
				worker.DoWork += new DoWorkEventHandler(Logic.SplitSheets);

				worker.RunWorkerCompleted +=
					new RunWorkerCompletedEventHandler(worker_Done);
				disableUI();
				worker.RunWorkerAsync(args.ToArray());
			}
			else
			{
				Log.Text = Log.Text + "No file selected.\n";
			}


		}

		private void RadioButton_Checked(object sender, RoutedEventArgs e)
		{
			Options.Default.DelEmptyAccounts = (bool)DeleteAcctRadio.IsChecked;
			Options.Default.Save();
		}
	}
}

