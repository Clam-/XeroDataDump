using PM = Xero.Api.Payroll.Australia.Model;
using CM = Xero.Api.Core.Model;
using System.ComponentModel;
using System.Linq;
using System.Collections.Generic;
using System;

using Xero.Api.Core.Model.Reports;
using System.Security.Cryptography;
using OfficeOpenXml;
using System.IO;

namespace XeroDataDump
{
	class Logic
	{
		static Dictionary<Guid, string> projectMap = new Dictionary<Guid, string>();

		static Dictionary<Guid, List<Tuple<Guid, decimal>>> projectsTime = null;

		static Dictionary<Guid, string> employees = new Dictionary<Guid, string>();
		static List<string> IncomeAccounts = new List<string>();

		static AustralianPayroll ap = null;
		static Core c = null;

		static BackgroundWorker worker = null;

		static int BUDGETCOLUMN = 3;
		static int ACTUALCOLUMN = 4;
		static int OVERALLROWSTART = 4;
		static int PROJETROWSTART = 4;
		static int PROJ_BUDGET_COLUMN = 3;
		static int PROJ_ACTUAL_COLUMN = 4;
		static int PROJ_HOURS_ROW = 4;
		static int PROJ_COST_ROW = 8;
		static int PROJ_OTHERS_ROW = 12;

		static int POSITIONS = 8; // TODO: CHANGE THIS

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

		private static bool initXero()
		{
			try
			{
				ap = new AustralianPayroll();
				c = new Core();
				return true;
			}
			catch (CryptographicException)
			{
				LogBox("Missing certificate information.");
			}
			catch (ArgumentException)
			{
				LogBox("Invalid certificate information");
			}
			return false;
		}

		private static ExcelWorksheet setupWorkbook(ExcelWorkbook wb, DateTime from, DateTime to)
		{
			var ws = wb.Worksheets.Add("Overall PL");
			ws.Cells["A1"].Value = "Profit and loss"; ws.Cells["C1"].Value = "Begining"; ws.Cells["D1"].Value = "Ending";
			ws.Cells["A2"].Value = Options.Default.OrganisationName; ws.Cells["C2"].Value = from.ToShortDateString(); ws.Cells["D2"].Value = to.ToShortDateString();

			ws.Cells[OVERALLROWSTART, BUDGETCOLUMN-1].Value = "Budget Full";
			ws.Cells[OVERALLROWSTART, BUDGETCOLUMN].Value = "Budget YTD"; ws.Cells[OVERALLROWSTART, ACTUALCOLUMN].Value = "Actual";
			ws.Cells[OVERALLROWSTART, ACTUALCOLUMN+1].Value = "Var $"; ws.Cells[OVERALLROWSTART, ACTUALCOLUMN+2].Value = "Var %";
			return ws;
		}

		private static void addOverallData(ExcelWorksheet ws, DateTime start, DateTime end)
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
							ws.Cells[rowNum, 1].Value = trow.Cells[0].Value;
							var acell = ws.Cells[rowNum, ACTUALCOLUMN];
							acell.Value = double.Parse(trow.Cells[1].Value);
							// insert formulas
							if (IncomeAccounts.Contains(ws.Cells[rowNum, 1].Value))
							{
								ws.Cells[rowNum, ACTUALCOLUMN + 1].Formula = "=" + acell.Address + "-" + ws.Cells[rowNum, BUDGETCOLUMN].Address;
								ws.Cells[rowNum, ACTUALCOLUMN + 2].Formula = "=" + acell.Address + "/" + ws.Cells[rowNum, BUDGETCOLUMN].Address;
							}
							else
							{
								ws.Cells[rowNum, ACTUALCOLUMN + 1].Formula = "=" + ws.Cells[rowNum, BUDGETCOLUMN].Address + "-" + acell.Address;
								ws.Cells[rowNum, ACTUALCOLUMN + 2].Formula = "=" + ws.Cells[rowNum, BUDGETCOLUMN].Address + "/" + acell.Address;
							}
						}
						rowNum++;
					}
				}
			}
			
		}

		private static bool IsEmptyCell(string data)
		{
			return string.IsNullOrWhiteSpace(data);
		}

		private static int searchText(ExcelWorksheet ws, int rowi, int coli, string text)
		{
			var row = rowi;
			string data = ws.Cells[row, coli].GetValue<string>();
			while (!IsEmptyCell(data))
			{
				if (data == text)
				{
					return row;
				}
				row = row + 1;
				data = ws.Cells[row, coli].GetValue<string>();
			}
			throw new ArgumentException("Text not found");
		}
		
		private static double getDouble(ExcelRangeBase rb)
		{
			double val = rb.GetValue<double>();
			return val;
		}

		private static void processOverallCostBudget(ExcelWorksheet ws, ExcelWorkbook budget, int months)
		{
			LogBox("Processing Overall Budget\n");
			var sheet = budget.Worksheets["Overall"];
			var rownum = OVERALLROWSTART + 1;
			var row = Options.Default.CostBudgetRow;

			string data = sheet.Cells[row, Options.Default.CostBudgetACCol].GetValue<string>();

			while (!IsEmptyCell(data))
			{
				int irow = -1;
				try
				{
					irow = searchText(ws, rownum, 1, data);
				} catch (ArgumentException)
				{
					// append
					LogBox("Append: " + data);
				}
				if (irow > 0) {
					LogBox("Summing: " + (Options.Default.CostBudgetYearCol + 1) + " - " + (Options.Default.CostBudgetYearCol + months));
					ws.Cells[irow, BUDGETCOLUMN].Value = sheet.Cells[row, Options.Default.CostBudgetYearCol + 1, row, Options.Default.CostBudgetYearCol + months].Sum(cell => { return getDouble(cell); });
					ws.Cells[irow, BUDGETCOLUMN - 1].Value = sheet.Cells[row, Options.Default.CostBudgetYearCol].Value;
				}
				row = row + 1;
				data = sheet.Cells[row, Options.Default.CostBudgetACCol].GetValue<string>();
			}
			//Format cells
			ws.Cells[OVERALLROWSTART, BUDGETCOLUMN - 1, ws.Dimension.Rows, ACTUALCOLUMN + 1].Style.Numberformat.Format = "$#,##0.00";
			ws.Cells[OVERALLROWSTART, ACTUALCOLUMN +2, ws.Dimension.Rows, ACTUALCOLUMN + 2].Style.Numberformat.Format = "0.00%";
		}

		private static void initializeState()
		{
			UpdateProjectMapping();
			projectsTime = new Dictionary<Guid, List<Tuple<Guid, decimal>>>();
			// unpack income accounts
			IncomeAccounts = Options.Default.IncAccts.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
		}

		internal static void YTDDataDump(object sender, DoWorkEventArgs e)
		{
			worker = sender as BackgroundWorker;

			if (!initXero()) { return; }

			object[] args = (object[])e.Argument;
			//string orgName = (string)args[0];
			//string budgetfname = (string)args[1];
			string savefname = System.IO.Path.Combine(Options.Default.OutputDir, "report.xlsx");
			DateTime to = (DateTime)args[0];
			int year = (int)args[1];
			int month = (int)args[2];

			// initialize state and income accounts
			initializeState();

			int days = DateTime.DaysInMonth(year, month);

			var start = new DateTime(year, month, 1);
			var end = to;
			// calc months
			var months = 0;
			while (start.AddMonths(months) < end) { months = months + 1; if (months > 10000) { throw new ArgumentException("Too many months. Are you sure YTD year is in the past?"); } }

			LogBox(string.Format("Running YTD Data Dump for {0}-{1}\n", start.ToShortDateString(), end.ToShortDateString()));

			var costBudgetwb = new ExcelPackage(new FileInfo(Options.Default.CostBudgetFile)).Workbook;
			var timeBudgetwb = new ExcelPackage(new FileInfo(Options.Default.TimeBudgetFile)).Workbook;
			// delete old
			File.Delete(savefname);
			var pkg = new ExcelPackage(new FileInfo(savefname));
			
			var wb = pkg.Workbook;

			var overallws = setupWorkbook(wb, start, end);

			addOverallData(overallws, start, end);

			//gatherTimesheetsProjects(start, end); // setup projectsTime

			processOverallCostBudget(overallws, costBudgetwb, months);
			processProjectActualCost(wb, start, end);

			//processProjectsTime(costBudgetwb, wb, start, end);

			// autosize before save
			foreach (var ws in wb.Worksheets)
			{
				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				for (var ncol = 1; ncol <= ws.Dimension.Columns; ncol++)
				{
					ws.Column(ncol).Width = ws.Column(ncol).Width + 5;
				}
			}

			pkg.Save();

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

		private static void setupProjectSheet(ExcelWorksheet ws, string title, DateTime from, DateTime to)
		{
			ws.Cells["A1"].Value = "Profit and loss"; ws.Cells["C1"].Value = "Begining"; ws.Cells["D1"].Value = "Ending";
			ws.Cells["A2"].Value = title; ws.Cells["C2"].Value = from.ToShortDateString(); ws.Cells["D2"].Value = to.ToShortDateString();

			// staff Hours
			ws.Cells[PROJ_HOURS_ROW, 1].Value = "Staff Hours"; ws.Cells[PROJ_HOURS_ROW, PROJ_BUDGET_COLUMN-1].Value = "Budget Full"; ws.Cells[PROJ_HOURS_ROW, PROJ_BUDGET_COLUMN].Value = "Budget YTD";
			ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN].Value = "Actual"; ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN + 1].Value = "Var"; ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN + 2].Value = "Var %";
			// staff Cost
			ws.Cells[PROJ_COST_ROW, 1].Value = "Staff Cost"; ws.Cells[PROJ_COST_ROW, PROJ_BUDGET_COLUMN - 1].Value = "Budget Full"; ws.Cells[PROJ_COST_ROW, PROJ_BUDGET_COLUMN].Value = "Budget YTD";
			ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN].Value = "Actual"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 1].Value = "Var"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 2].Value = "Var %"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 3].Value = "Rate?";

			ws.Cells[PROJETROWSTART, PROJ_BUDGET_COLUMN].Value = "Budget"; ws.Cells[PROJETROWSTART, PROJ_ACTUAL_COLUMN].Value = "Actual"; ws.Cells[PROJETROWSTART, PROJ_ACTUAL_COLUMN+1].Value = "Var"; ws.Cells[PROJETROWSTART, PROJ_ACTUAL_COLUMN+2].Value = "Var %";
			ws.Cells["A4"].Value = "Staff Hours"; ws.Cells["B4"].Value = "Pos";
		}

		private static Dictionary<string, Tuple<string,double>> getBudgetEmployees(ExcelWorksheet bws, DateTime from)
		{
			var col = from.Month + BUDGETTIMECOLUMNSTART - 1;
			var emps = new Dictionary<string, Tuple<string, double>>();
			if (bws == null) { return emps; }

			var i = bws.Dimension.Rows;
			for (var rownum = 0; rownum < bws.Dimension.Rows; rownum++)
			{
				if (rownum == 1) { continue; } //increment rownum
				var emp = bws.Cells[rownum, 1].Value.ToString();
				if (!string.IsNullOrWhiteSpace(emp))
				{
					try {
						emps[emp] = new Tuple<string, double>(bws.Cells[rownum, 2].Value.ToString(), bws.Cells[rownum, col].GetValue<double>());
					} catch (Exception) { }
				}
				rownum++;
			}
			return emps;
		}

		private static Dictionary<string, List<string>> getProjectsProfitLoss(DateTime start, DateTime end)
		{
			var plReport = c.Reports.ProfitAndLoss(start, from: start, to: end, standardLayout: true, trackingCategory: ProjectTCID);
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

		private static void processProjectsTime(ExcelWorkbook bwb, ExcelWorkbook wb, DateTime from, DateTime to)
		{
			// dump employee hour data
			foreach (var project in projectsTime)
			{
				var ws = wb.Worksheets.Add(projectMap[project.Key]);
				ExcelWorksheet bws = null;
				try {
					bws = bwb.Worksheets[projectMap[project.Key]];
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
					ws.Cells[rownum, 1].Value = ename;
					var bcell = ws.Cells[rownum, PROJ_BUDGET_COLUMN];
					var acell = ws.Cells[rownum, PROJ_ACTUAL_COLUMN];
					acell.Value = emp.Item2;
					// TODO: ???
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 1].Formula = "=" + bcell.Address + "-" + acell.Address;
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 2].Formula = "=" + bcell.Address + "/" + acell.Address;

					if (budgetemps.ContainsKey(ename))
					{
						bcell.Value = budgetemps[ename].Item2;
						ws.Cells[rownum, 2].Value = budgetemps[ename].Item1;
					}
					rownum++;
				}
				// remaining items
				foreach (var ik in budgetemps.Keys.Except(employees))
				{
					ws.Cells[rownum, 1].Value = ik;
					ws.Cells[rownum, PROJ_BUDGET_COLUMN].Value = budgetemps[ik];
				}
			}
		}

		private static void processProjectActualCost(ExcelWorkbook wb, DateTime from, DateTime to)
		{
			// dump remaining project data...
			var projectReports = getProjectsProfitLoss(from, to);
			var categories = projectReports[""];
			projectReports.Remove("");
			foreach (var k in projectReports.Keys)
			{
				ExcelWorksheet ws = wb.Worksheets[k];
				if (ws == null)
				{
					if (k.Equals("Unassigned") || k.Equals("Total"))
						continue;
					else
					{
						ws = wb.Worksheets.Add(k);
						setupProjectSheet(ws, k, from, to);
					}
				}
				var rownum = ws.Dimension.Rows;
				rownum = rownum + 2;
				ws.Cells[rownum, PROJ_BUDGET_COLUMN].Value = "Budget"; ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Value = "Actual"; ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 1].Value = "Var $"; ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 2].Value = "Var %";
				ws.Cells[rownum, 1].Value = "Others";
				rownum++;

				int i = 0;
				foreach (var line in projectReports[k])
				{
					ws.Cells[rownum, 1].Value = categories[i];
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Value = double.Parse(line);
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
