using PM = Xero.Api.Payroll.Australia.Model;
using CM = Xero.Api.Core.Model;
using System.ComponentModel;
using System.Linq;
using System.Collections.Generic;
using System;
using ClosedXML.Excel;
using Xero.Api.Core.Model.Reports;
using System.Security.Cryptography;

namespace XeroDataDump
{
	class Logic
	{
		static Dictionary<Guid, string> projectMap = new Dictionary<Guid, string>();

		static Dictionary<Guid, List<Tuple<Guid, decimal>>> projectsTime = null;

		static Dictionary<Guid, string> employees = new Dictionary<Guid, string>();

		static AustralianPayroll ap = null;
		static Core c = null;

		static BackgroundWorker worker = null;

		static int BUDGETCOLUMN = 2;
		static int ACTUALCOLUMN = 3;
		static int OVERALLROWSTART = 4;
		static int PROJETROWSTART = 4;
		static int BUDGETPROJCOLUMN = 3;
		static int ACTUALPROJCOLUMN = 4;
		static int BUDGETFULLYRPROJCOLUMN = 7;

		// static int BUDGETTIMEROWSTART = 3; Progmatically find
		static int BUDGETTIMECOLSEARCH = 2;
		static string BUDGETTIMESSEARCH = "Total Days YTD";
		static int BUDGETTIMECOLUMNSTART = 3;


		static Guid ProjectTCID;

		internal static void LogBox(string msg)
		{
			if (worker != null)
				worker.ReportProgress(0, msg+"\n");
		}

		private static void initXero()
		{
			try
			{
				ap = new AustralianPayroll();
				c = new Core();
			}
			catch (CryptographicException)
			{
				LogBox("Missing certificate information.");
			}
		}

		private static IXLWorksheet setupWorkbook(XLWorkbook wb, DateTime from, DateTime to, string orgName)
		{
			var ws = wb.Worksheets.Add("Overall PL");
			ws.Cell("A1").Value = "Profit and loss"; ws.Cell("C1").Value = "Begining"; ws.Cell("D1").Value = "Ending";
			ws.Cell("A2").Value = orgName; ws.Cell("C2").Value = from.ToShortDateString(); ws.Cell("D2").Value = to.ToShortDateString();

			ws.Cell(OVERALLROWSTART, BUDGETCOLUMN).Value = "Budget"; ws.Cell(OVERALLROWSTART, ACTUALCOLUMN).Value = "Actual";
			ws.Cell(OVERALLROWSTART, ACTUALCOLUMN+1).Value = "Var $"; ws.Cell(OVERALLROWSTART, ACTUALCOLUMN+2).Value = "Var %";
			return ws;
		}

		private static void addOverallData(IXLWorksheet ws, DateTime start, DateTime end)
		{
			// Overall profit and loss report
			var plReport = c.Reports.ProfitAndLoss(start, from: start, to: end, standardLayout: true);
			LogBox("Getting Overall PL\n");

			var rowNum = OVERALLROWSTART + 1;

			foreach (var row in plReport.Rows)
			{
				if (row.Cells == null)
				{
					foreach (var trow in row.Rows)
					{
						if (trow.Cells != null)
						{
							// get name and value, cell[0] and [1]
							ws.Cell(rowNum, 1).Value = trow.Cells[0].Value;
							var acell = ws.Cell(rowNum, ACTUALCOLUMN);
							acell.Value = trow.Cells[1].Value;
							// insert formulas
							ws.Cell(rowNum, ACTUALCOLUMN + 1).FormulaA1 = "=" + ws.Cell(rowNum, BUDGETCOLUMN).Address + "-" + acell.Address;
							ws.Cell(rowNum, ACTUALCOLUMN + 2).FormulaA1 = "=" + ws.Cell(rowNum, BUDGETCOLUMN).Address + "/" + acell.Address;
							// TODO: format special cells
						}
						rowNum++;
					}
				}
			}
		}

		private static void getInsertBudget(IXLWorksheet ws, DateTime start, DateTime end)
		{
			LogBox("Getting Budget\n");
			var budget = c.Reports.BudgetSummary(date: start, periods: 1, timeFrame: CM.Types.BudgetSummaryTimeframeType.Month);
			Console.WriteLine("getInsertBudget Fields: " + string.Join(",", budget.Fields.Select(x => x.Value)));
			var rownum = OVERALLROWSTART + 1;
			foreach (var row in budget.Rows)
			{
				foreach (var r in row.Rows ?? Enumerable.Empty<ReportRow>())
				{
					var rowdata = r.Cells;
					if ((string)ws.Cell(rownum, BUDGETCOLUMN - 1).Value == rowdata[0].Value)
					{
						Console.WriteLine(string.Format("Looking for ({0}) FOUND ({1})", (string)ws.Cell(rownum, BUDGETCOLUMN - 1).Value, rowdata[0].Value));
						ws.Cell(rownum, BUDGETCOLUMN).Value = rowdata[1].Value;
					}
					else
					{
						Console.WriteLine(string.Format("Looking for ({0}) NOT FOUND...", (string)ws.Cell(rownum, BUDGETCOLUMN - 1).Value));
						// find corresponding row for data
						var irow = rownum; var found = true;
						while ((string)ws.Cell(irow, BUDGETCOLUMN - 1).Value != rowdata[0].Value)
						{
							Console.WriteLine(string.Format("Looking in ({0}) for ({1}) GOT ({2}).", irow, (string)ws.Cell(irow, BUDGETCOLUMN - 1).Value, rowdata[0].Value));
							irow++;
							if (irow >= ws.LastRowUsed().RowNumber())
							{
								Console.WriteLine("REACHED END.");
								found = false;
								break;
							}
						}
						if (found)
						{
							Console.WriteLine(string.Format("FOUND ({0}) at ({1}).", (string)ws.Cell(rownum, BUDGETCOLUMN - 1).Value, irow));
							rownum = irow;
						}
						else
						{
							Console.WriteLine("Didn't find. Inserting.");
							ws.Row(rownum).InsertRowsAbove(1);
							ws.Cell(rownum, ACTUALCOLUMN).Value = 0;
						}
						ws.Cell(rownum, BUDGETCOLUMN-1).Value = rowdata[0].Value;
						ws.Cell(rownum, BUDGETCOLUMN).Value = rowdata[1].Value;
					}
					rownum++;
				}
			}
		}

		private static void initializeState()
		{
			UpdateProjectMapping();
			projectsTime = new Dictionary<Guid, List<Tuple<Guid, decimal>>>();
		}

		internal static void YTDDataDump(object sender, DoWorkEventArgs e)
		{
			initXero();

			worker = sender as BackgroundWorker;

			object[] args = (object[])e.Argument;
			string orgName = (string)args[0];
			string budgetfname = (string)args[1];
			string savefname = (string)args[2];
			ap = (AustralianPayroll)args[3];
			c = (Core)args[4];
			int year = (int)args[5];
			int month = (int)args[6];

			// initialize state
			initializeState();

			int days = DateTime.DaysInMonth(year, month);

			var start = new DateTime(year, month, 1);
			var end = new DateTime(year, month, days);

			LogBox(string.Format("Running Monthly Data Dump for {0}-{1}\n", start.ToShortDateString(), end.ToShortDateString()));
			var budgetwb = new XLWorkbook(budgetfname);
			var wb = new XLWorkbook();

			var overallws = setupWorkbook(wb, start, end, orgName);

			addOverallData(overallws, start, end);

			gatherTimesheetsProjects(start, end); // setup projectsTime

			getInsertBudget(overallws, start, end);

			processProjects(budgetwb, wb, start, end);

			// autosize before save
			foreach (var ws in wb.Worksheets)
			{
				foreach( var col in ws.Columns())
				{
					col.AdjustToContents();
					col.Width = col.Width + 3;
				}
			}
			wb.SaveAs(savefname);

			LogBox("Done Monthly dump.\n");

		}

		private static string getEmployeeName(Guid empID)
		{
			string emp = null;
			if (employees.TryGetValue(empID, out emp))
			{
				return emp;
			}
			else
			{
				var e = ap.Employees.Find(empID.ToString());
				emp = e.FirstName + " " + e.LastName;
				employees[empID] = emp;
				return emp;
			}
		}

		private static void setupProjectSheet(IXLWorksheet ws, string title, DateTime from, DateTime to)
		{
			ws.Cell("A1").Value = "Profit and loss"; ws.Cell("C1").Value = "Begining"; ws.Cell("D1").Value = "Ending";
			ws.Cell("A2").Value = title; ws.Cell("C2").Value = from.ToShortDateString(); ws.Cell("D2").Value = to.ToShortDateString();

			ws.Cell(PROJETROWSTART, BUDGETPROJCOLUMN).Value = "Budget"; ws.Cell(PROJETROWSTART, ACTUALPROJCOLUMN).Value = "Actual"; ws.Cell(PROJETROWSTART, ACTUALPROJCOLUMN+1).Value = "Var"; ws.Cell(PROJETROWSTART, ACTUALPROJCOLUMN+2).Value = "Var %";
			ws.Cell("A4").Value = "Staff Hours"; ws.Cell("B4").Value = "Pos";
		}

		private static Dictionary<string, Tuple<string,double>> getBudgetEmployees(IXLWorksheet bws, DateTime from)
		{
			var col = from.Month + BUDGETTIMECOLUMNSTART - 1;
			var rownum = 1;
			var emps = new Dictionary<string, Tuple<string, double>>();
			if (bws == null) { return emps; }
			
			foreach (var row in bws.RangeUsed().Rows())
			{
				if (rownum == 1) { rownum++; continue; }
				var emp = row.Cell(1).Value.ToString();
				if (!string.IsNullOrWhiteSpace(emp))
				{
					try {
						emps[emp] = new Tuple<string, double>(row.Cell(2).Value.ToString(), row.Cell(col).GetDouble());
					} catch (Exception) { }
				}
				rownum++;
			}
			return emps;
		}

		private static Dictionary<string, List<string>> getProjectsProfitLoss(DateTime start, DateTime end)
		{
			var plReport = c.Reports.ProfitAndLoss(start, from: start, to: end, standardLayout: false, trackingCategory: ProjectTCID);
			List<string> headers = null;
			Dictionary<string, List<string>> preport = new Dictionary<string, List<string>>();

			foreach (var row in plReport.Rows)
			{
				if (headers == null)
					headers = row.Cells.Select(x => x.Value).ToList();
				foreach (var r in row.Rows ?? Enumerable.Empty<ReportRow>())
				{
					int i = 0;
					foreach (var c in r.Cells ?? Enumerable.Empty<ReportCell>())
					{
						if (!preport.ContainsKey(headers[i]))
						{
							preport[headers[i]] = new List<string>();
						}
						preport[headers[i]].Add(c.Value);
						i++;
					}
				}
			}
			return preport;
		}

		private static void processProjects(XLWorkbook bwb, XLWorkbook wb, DateTime from, DateTime to)
		{
			// dump employee hour data
			foreach (var project in projectsTime)
			{
				var ws = wb.Worksheets.Add(projectMap[project.Key]);
				IXLWorksheet bws = null;
				try {
					bws = bwb.Worksheet(projectMap[project.Key]);
				} catch (Exception) {
					
				}
				var budgetemps = getBudgetEmployees(bws, from);

				setupProjectSheet(ws, projectMap[project.Key], from, to);
				var employees = new HashSet<string>();
				var rownum = PROJETROWSTART+1;
				foreach (var emp in project.Value)
				{
					var ename = getEmployeeName(emp.Item1);
					employees.Add(ename);
					ws.Cell(rownum, 1).Value = ename;
					var bcell = ws.Cell(rownum, BUDGETPROJCOLUMN);
					var acell = ws.Cell(rownum, ACTUALPROJCOLUMN);
					acell.Value = emp.Item2;
					ws.Cell(rownum, ACTUALPROJCOLUMN + 1).FormulaA1 = "=" + bcell.Address + "-" + acell.Address;
					ws.Cell(rownum, ACTUALPROJCOLUMN + 2).FormulaA1 = "=" + bcell.Address + "/" + acell.Address;
					if (budgetemps.ContainsKey(ename))
					{
						bcell.Value = budgetemps[ename].Item2;
						ws.Cell(rownum, 2).Value = budgetemps[ename].Item1;
					}
					rownum++;
				}
				// remaining items
				foreach (var ik in budgetemps.Keys.Except(employees))
				{
					ws.Cell(rownum, 1).Value = ik;
					ws.Cell(rownum, BUDGETPROJCOLUMN).Value = budgetemps[ik];
				}
			}
			
			// dump remaining project data...
			var projectReports = getProjectsProfitLoss(from, to);
			var categories = projectReports[""];
			projectReports.Remove("");
			foreach (var k in projectReports.Keys)
			{
				IXLWorksheet ws = null;
				try
				{
					ws = wb.Worksheet(k);
				}
				catch (Exception)
				{
					if (k.Equals("Unassigned") || k.Equals("Total"))
						continue;
					else
					{
						ws = wb.Worksheets.Add(k);
						setupProjectSheet(ws, k, from, to);
					}
				}
				var rownum = ws.LastRowUsed().RowNumber();
				rownum = rownum + 2;
				ws.Cell(rownum, BUDGETPROJCOLUMN).Value = "Budget"; ws.Cell(rownum, ACTUALPROJCOLUMN).Value = "Actual"; ws.Cell(rownum, ACTUALPROJCOLUMN + 1).Value = "Var $"; ws.Cell(rownum, ACTUALPROJCOLUMN + 2).Value = "Var %";
				ws.Cell(rownum, 1).Value = "Others";
				rownum++;

				int i = 0;
				foreach (var line in projectReports[k])
				{
					ws.Cell(rownum, 1).Value = categories[i];
					ws.Cell(rownum, ACTUALPROJCOLUMN).Value = line;
					i++; rownum++;
				}
				
			}
		}

		private static void gatherTimesheetsProjects(DateTime start, DateTime end)
		{
			var sq = string.Format("StartDate == DateTime.Parse(\"{0}-{1}-1\")", start.Year, start.Month);
			var eq = string.Format("EndDate == DateTime.Parse(\"{0}-{1}-{2}\")", end.Year, end.Month, end.Day);
			foreach (PM.Timesheet it in ap.Timesheets.Where(sq).And(eq).Find())
			{
				//LogBox("E: " + ap.Employees.Find(it.EmployeeId.ToString()).FirstName + "\n");
				foreach (PM.TimesheetLine itl in it.TimesheetLines)
				{
					//ap.PayItems.Find(itl.EarningsRateId.ToString()).EarningsRates
					//LogBox("\t Project: " + TCmapping[itl.TrackingItemID] + " Units: " + itl.NumberOfUnits.Sum() + "\n");
					List<Tuple<Guid, decimal>> projectHours = null;
					if (projectsTime.TryGetValue(itl.TrackingItemID, out projectHours))
					{
						projectHours.Add(new Tuple<Guid, decimal>(it.EmployeeId, itl.NumberOfUnits.Sum()));
					}
					else
					{
						projectsTime[itl.TrackingItemID] = new List<Tuple<Guid, decimal>> { new Tuple<Guid, decimal>(it.EmployeeId, itl.NumberOfUnits.Sum()) };
					}
				}
			}
		}

		static void UpdateProjectMapping()
		{
			if (projectMap.Count != 0) return; // update Tracking Categories if empty

			foreach (var tc in c.TrackingCategories.Where("Name == \"Projects\"").Find())
			{
				ProjectTCID = tc.Id;
				foreach (var opt in tc.Options) {
					projectMap.Add(opt.Id, opt.Name);
				}
			}
		}
	}
}
