using PM = Xero.Api.Payroll.Australia.Model;
using CM = Xero.Api.Core.Model;
using System.ComponentModel;
using System.Linq;
using System.Collections.Generic;
using System;

namespace XeroDataDump
{
	class Logic
	{
		static Dictionary<Guid, string> TCmapping = new Dictionary<Guid, string>();

		static Dictionary<Guid, List<Tuple<Guid, decimal>>> projectsTime = new Dictionary<Guid, List<Tuple<Guid, decimal>>>();

		internal static void CoreTest(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			worker.ReportProgress(0, "Running Core...\n");

			object[] args = (object[])e.Argument;
			AustralianPayroll ap = (AustralianPayroll)args[0];
			Core c = (Core)args[1];
			UpdateTCMapping(c);

			//foreach (CM.Invoice inv in c.Invoices.Find())
			//{
			//	worker.ReportProgress(0, inv.Id + " - " + inv.Status + " - " + inv.Total + "\n");
			//}

			var budget = c.Reports.ProfitAndLoss(DateTime.Now.AddDays(-37), from: DateTime.Now.AddDays(-37), to:DateTime.Now.AddDays(-6), standardLayout: false);
			worker.ReportProgress(0, String.Join(",", budget.Fields.Select(x => x.Value)));
			worker.ReportProgress(0, "\n");
			foreach (var row in budget.Rows)
			{
				if (row.Cells == null)
				{
					foreach (var trow in row.Rows)
					{
						if (trow.Cells != null)
						{
							foreach (var cell in trow.Cells)
							{
								if (cell.Value != null)
									worker.ReportProgress(0, cell.Value + "\t" );
								else
									worker.ReportProgress(0, "NONE");
							}
							worker.ReportProgress(0, "\n");
						}
					}
					continue;
				}
				foreach (var cell in row.Cells)
				{
					if (cell.Value != null)
						worker.ReportProgress(0, cell.Value + "\t");
					else
						worker.ReportProgress(0, "NONE");
				}
				worker.ReportProgress(0, "\n");
			}
			worker.ReportProgress(0, "Done Core.\n");
		}

		internal static void PayrollTest(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			worker.ReportProgress(0, "Running Payroll...\n");

			object[] args = (object[])e.Argument;
			AustralianPayroll ap = (AustralianPayroll)args[0];
			Core c = (Core)args[1];
			UpdateTCMapping(c);

			//c.TrackingCategories.Find("aab");

			//foreach (var pi in ap.PayItems.Find())
			//{
			//	foreach (var er in pi.EarningsRates)
			//	{ worker.ReportProgress(0, "ER ID: " + er.Id + " Name: " + er.Name + " type: " + er.EarningsType + "\n"); }
			//	foreach (var dt in pi.DeductionTypes)
			//	{ worker.ReportProgress(0, "DT ID: " + dt.Id + " Name: " + dt.Name + " type: " + dt.AccountCode + "\n"); }

			//}

			//worker.ReportProgress(0, "From: " + "To: " + "\n");
			foreach (PM.Timesheet it in ap.Timesheets.Where("StartDate == DateTime.Parse(\"2015-12-1\")").Find())
            {
				worker.ReportProgress(0, "E: " + ap.Employees.Find(it.EmployeeId.ToString()).FirstName + "\n");
				foreach (PM.TimesheetLine itl in it.TimesheetLines)
				{
					//ap.PayItems.Find(itl.EarningsRateId.ToString()).EarningsRates
					worker.ReportProgress(0, "\t Project: " + TCmapping[itl.TrackingItemID] + " Units: " + itl.NumberOfUnits.Sum() + "\n");
				}
			}
			worker.ReportProgress(0, "Done Payroll test.\n");
		}

		internal static void MonthlyHourDump(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;

			object[] args = (object[])e.Argument;
			AustralianPayroll ap = (AustralianPayroll)args[0];
			Core c = (Core)args[1];
			int year = (int)args[2];
			int month = (int)args[3];
			UpdateTCMapping(c);

			int days = DateTime.DaysInMonth(year, month);

			var sq = string.Format("StartDate == DateTime.Parse(\"{0}-{1}-1\")", year, month);
			var eq = string.Format("EndDate == DateTime.Parse(\"{0}-{1}-{2}\")", year, month, days);
			worker.ReportProgress(0, "Running Monthly Hour Dump\n");
			worker.ReportProgress(0, sq + "\n" + eq + "\n");
			foreach (PM.Timesheet it in ap.Timesheets.Where(sq).And(eq).Find())
			{
				//worker.ReportProgress(0, "E: " + ap.Employees.Find(it.EmployeeId.ToString()).FirstName + "\n");
				foreach (PM.TimesheetLine itl in it.TimesheetLines)
				{
					//ap.PayItems.Find(itl.EarningsRateId.ToString()).EarningsRates
					//worker.ReportProgress(0, "\t Project: " + TCmapping[itl.TrackingItemID] + " Units: " + itl.NumberOfUnits.Sum() + "\n");
					List<Tuple<Guid, decimal>> projectHours = null;
					if (projectsTime.TryGetValue(itl.TrackingItemID, out projectHours))
					{
						projectHours.Add(new Tuple<Guid, decimal>(it.EmployeeId, itl.NumberOfUnits.Sum()));
					} else
					{
						projectsTime[itl.TrackingItemID] = new List<Tuple<Guid, decimal>> { new Tuple<Guid, decimal>(it.EmployeeId, itl.NumberOfUnits.Sum()) };
					}
                }
			}
			foreach (var project in projectsTime)
			{
				worker.ReportProgress(0, "Project: " + TCmapping[project.Key] + "\n");
				foreach (var emp in project.Value)
				{
					var iemp = ap.Employees.Find(emp.Item1.ToString());
					
                    worker.ReportProgress(0, iemp.FirstName.Substring(0,1) + iemp.LastName.Substring(0,1) + "\t" + emp.Item2+"\n");
                }
			}
			worker.ReportProgress(0, "Done hours.\n");

		}
		internal static void YTDHourDump(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			worker.ReportProgress(0, "Don't know what YTD is.\n");

			object[] args = (object[])e.Argument;
			AustralianPayroll ap = (AustralianPayroll)args[0];
			Core c = (Core)args[1];
			UpdateTCMapping(c);
		}

		static void UpdateTCMapping(Core c)
		{
			if (TCmapping.Count != 0) return; // update Tracking Categories if empty

			foreach (var tc in c.TrackingCategories.Find())
			{
				foreach (var opt in tc.Options) {
					Console.WriteLine(opt.Id.ToString() + "-" + opt.Name);
					TCmapping.Add(opt.Id, opt.Name);
				}
			}
		}
	}
}
