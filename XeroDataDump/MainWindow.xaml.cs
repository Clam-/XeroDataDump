
using System.Security.Cryptography;
using System.Windows;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading;
namespace XeroDataDump
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>

	public enum TimeOptions
	{
		YTD, Monthly
	}

	public partial class MainWindow : Window
	{
		AustralianPayroll ap;
		Core c;
		BackgroundWorker worker;

		private TimeOptions timeopt = TimeOptions.Monthly; // set your default value here

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

		public TimeOptions TimeFrame
		{
			get { return timeopt; }
			set { timeopt = value; } // Cannot set this. InvokePropertyChanged("Options");
		}

		public MainWindow()
		{
			// setup UI
			InitializeComponent();
			Year.Text = DateTime.Now.Year.ToString();

			try {
				ap = new AustralianPayroll();
				c = new Core();
			} catch (CryptographicException) {
				Log.Text = "Missing certificate information.";
				GetDataButton.IsEnabled = false;
			}
		}

		private void GetDataButton_Click(object sender, RoutedEventArgs e)
		{
			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;
			List<object> args = new List<object> { ap, c };

			//worker.DoWork += new DoWorkEventHandler(Logic.CoreTest);
			if (TimeFrame == TimeOptions.Monthly) {
				worker.DoWork += new DoWorkEventHandler(Logic.MonthlyHourDump);
				int year = -1;
				if (int.TryParse(Year.Text, out year)) { args.AddRange(new object[] { year, Month.SelectedIndex+1 });}
				else { 
					Log.Text = Log.Text + "Year is not a number.";
					return;
				}

			}
			else if (TimeFrame == TimeOptions.YTD) { worker.DoWork += new DoWorkEventHandler(Logic.CoreTest); }
			
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(worder_Done);
			worker.ProgressChanged += worker_ProgressChanged;
			GetDataButton.IsEnabled = false;
			//disableUI();
			worker.RunWorkerAsync(args.ToArray());

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

			GetDataButton.IsEnabled = true;
		}
	}
}
