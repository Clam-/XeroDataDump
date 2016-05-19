﻿
using System.Security.Cryptography;
using System.Windows;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading;
using Microsoft.Win32;
using System.IO;
using System.Configuration;

namespace XeroDataDump
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>

	public partial class MainWindow : Window
	{
		static string ORGNAMEKEY = "OrganisationName";

		AustralianPayroll ap;
		Core c;
		BackgroundWorker worker;
		string oldOrgName = "";
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
			OrgName.Text = Options.Default.OrganisationName ?? OrgName.Text;
			Month.SelectedIndex = Options.Default.Month;
			Year.Text = Options.Default.Year.ToString();
			oldMonth = Month.SelectedIndex;
			oldOrgName = OrgName.Text;
			int.TryParse(Year.Text, out oldYear);

			try {
				ap = new AustralianPayroll();
				c = new Core();
			} catch (CryptographicException) {
				Log.Text = "Missing certificate information.";
				GetDataButton.IsEnabled = false;
			}
		}

		private void disableUI()
		{
			GetDataButton.IsEnabled = false;
			SaveName.IsEnabled = false;
			SaveBrowse.IsEnabled = false;
			BudgetBrowse.IsEnabled = false;
			BudgetFname.IsEnabled = false;
		}
		private void enableUI()
		{
			GetDataButton.IsEnabled = true;
			SaveName.IsEnabled = true;
			SaveBrowse.IsEnabled = true;
			BudgetBrowse.IsEnabled = true;
			BudgetFname.IsEnabled = true;
		}

		private void GetDataButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(SaveName.Text) || string.IsNullOrWhiteSpace(BudgetFname.Text))
			{
				Log.Text = Log.Text + "Require filename for Save File and Budget Time File.\n";
				return;
			}
			if (!File.Exists(BudgetFname.Text))
			{
				Log.Text = Log.Text + "Budget Time File doesn't exist.\n";
				return;
			}
			List<object> args = new List<object> { OrgName.Text, BudgetFname.Text, SaveName.Text, ap, c };

			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;

			//worker.DoWork += new DoWorkEventHandler(Logic.CoreTest);
			// Actual logic.
			worker.DoWork += new DoWorkEventHandler(Logic.YTDDataDump);
			int year = -1;
			if (int.TryParse(Year.Text, out year)) { args.AddRange(new object[] { year, Month.SelectedIndex+1 });}
			else { 
				Log.Text = Log.Text + "Year is not a number.\n";
				return;
			}
				
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(worder_Done);
			worker.ProgressChanged += worker_ProgressChanged;
			disableUI();
			worker.RunWorkerAsync(args.ToArray());

		}

		private void BudgetBrowse_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openDialog = new OpenFileDialog();
			openDialog.Filter = "Excel Workbook|*.xlsx";
			openDialog.Title = "Open Time Budget File...";
			openDialog.ShowDialog();

			// If the file name is not an empty string open it for saving.
			if (openDialog.FileName != "")
			{
				// get filename
				BudgetFname.Text = openDialog.FileName;
			}
		}
		private void SaveBrowse_Click(object sender, RoutedEventArgs e)
		{
			// Displays a SaveFileDialog so the user can save the Image
			// assigned to Button2. https://msdn.microsoft.com/en-us/library/sfezx97z(v=vs.110).aspx
			SaveFileDialog saveDialog = new SaveFileDialog();
			saveDialog.Filter = "Excel Workbook|*.xlsx";
			saveDialog.Title = "Save to location...";
			saveDialog.ShowDialog();

			// If the file name is not an empty string open it for saving.
			if (saveDialog.FileName != "")
			{
				// get filename
				SaveName.Text = saveDialog.FileName;
			}
		}

		void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			if (e.UserState != null)
				Log.Text += e.UserState as String;
			//progress.Value = e.ProgressPercentage;

		}

		void worder_Done(object sender, RunWorkerCompletedEventArgs e)
		{
			// First, handle the case where an exception was thrown. 
			if (e.Error != null)
			{
                Log.Text += e.Error.Message;
				//Logger.log(TraceEventType.Error, 9, "Worker Exception\r\n" + e.Error.GetType() + ":" + e.Error.Message + "\r\n" + e.Error.StackTrace);
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

			// If the file name is not an empty string open it for saving.
			if (openDialog.FileName != "")
			{
				// get filename
				CostBudgetFname.Text = openDialog.FileName;
			}
		}
	}
}
