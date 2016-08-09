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

		static Dictionary<int, Tuple<string, decimal>> posIDmap = new Dictionary<int, Tuple<string, decimal>>();
		static Dictionary<string, decimal> posmap = new Dictionary<string, decimal>();
		static Dictionary<string, Dictionary<string, decimal>> projPosHours = new Dictionary<string, Dictionary<string,decimal>>();
		static Dictionary<string, string> projTranslations = new Dictionary<string, string>();

		static BackgroundWorker worker = null;

		static int BUDGETCOLUMN = 3;
		static int ACTUALCOLUMN = 4;
		static int OVERALLROWSTART = 4;
		static int PROJETROWSTART = 4;
		static int PROJ_BUDGET_COLUMN = 3;
		static int PROJ_ACTUAL_COLUMN = 4;
		static int PROJ_HOURS_ROW = 4;
		static int PROJ_COST_ROW = 6;
		static int PROJ_OTHERS_ROW = 8;
		static int TIMESHEET_HOURS_TOTAL = 38; // AL total column
		static int TIMESHEET_PROJ_COL = 5;

		// static int BUDGETTIMEROWSTART = 3; Progmatically find
		//static int BUDGETTIMECOLSEARCH = 2;
		//static string BUDGETTIMESSEARCH = "Total Days YTD";
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

		private static void initPositions()
		{
			using (StringReader reader = new StringReader(Options.Default.Translations))
			{
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					// split line up in to ID, POS, COST
					var ls = line.Split(' ');
					if (ls.Count() != 2)
					{
						LogBox(string.Format("Invalid number of items in Translations ({0})", line));
					}
					else
					{
						projTranslations[ls[0]] = ls[1];
					}
				}
			}

			using (StringReader reader = new StringReader(Options.Default.Positions))
			{
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					// split line up in to ID, POS, COST
					var ls = line.Split(new char[] { ' ' }, 3);
					if (ls.Count() < 3)
					{
						LogBox(string.Format("Invalid number of items in Positions ({0})", line));
					} else
					{
						var parsed = false;
						int id = 0;
						parsed = int.TryParse(ls[0], out id);
						var pos = ls[2];
						decimal cost;
						parsed = parsed & decimal.TryParse(ls[1], out cost);
						if (!parsed)
						{
							LogBox(string.Format("Don't understand positions line ({0})", line));
						} else
						{
							posIDmap[id] = new Tuple<string, decimal>(pos, cost);
							posmap[pos] = cost;
						}
					}
				}
			}
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
							acell.Value = decimal.Parse(trow.Cells[1].Value);
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
				if (data.Equals(text, StringComparison.Ordinal))
				{
					return row;
				}
				row = row + 1;
				data = ws.Cells[row, coli].GetValue<string>();
			}
			throw new ArgumentException("Text not found");
		}
		
		private static decimal getDecimal(ExcelRangeBase rb)
		{
			decimal val = rb.GetValue<decimal>();
			return val;
		}

		private static void processOverallCostBudget(ExcelWorksheet ws, ExcelWorkbook budget, int months)
		{
			LogBox("Processing Overall Budget\n");
			var budgetSheet = budget.Worksheets["Overall"];
			var rownum = OVERALLROWSTART + 1;
			var row = Options.Default.CostBudgetRow + 1;

			string budgetAccName = budgetSheet.Cells[row, Options.Default.CostBudgetACCol].GetValue<string>();

			while (!IsEmptyCell(budgetAccName))
			{
				int irow = -1;
				try
				{
					irow = searchText(ws, rownum, 1, budgetAccName);
				} catch (ArgumentException)
				{
					// Missing budget
					LogBox(string.Format("MISSING FROM ACTUAL ({0}).", budgetAccName));
				}
				if (irow > 0) {
					//LogBox("Summing: " + (Options.Default.CostBudgetYearCol + 1) + " - " + (Options.Default.CostBudgetYearCol + months));
					ws.Cells[irow, BUDGETCOLUMN].Value = budgetSheet.Cells[row, Options.Default.CostBudgetYearCol + 1, row, Options.Default.CostBudgetYearCol + months].Sum(cell => { return getDecimal(cell); });
					ws.Cells[irow, BUDGETCOLUMN - 1].Value = budgetSheet.Cells[row, Options.Default.CostBudgetYearCol].Value;
				}
				row = row + 1;
				budgetAccName = budgetSheet.Cells[row, Options.Default.CostBudgetACCol].GetValue<string>();
			}
			//Format cells
			ws.Cells[OVERALLROWSTART, BUDGETCOLUMN - 1, ws.Dimension.Rows, ACTUALCOLUMN + 1].Style.Numberformat.Format = "$#,##0.00";
			ws.Cells[OVERALLROWSTART, ACTUALCOLUMN +2, ws.Dimension.Rows, ACTUALCOLUMN + 2].Style.Numberformat.Format = "0.00%";
		}

		private static void processProjectsCostBudget(ExcelWorkbook wb, ExcelWorkbook budget, int months)
		{
			LogBox("Processing Project Budgets - Cost\n");

			foreach (var ws in wb.Worksheets)
			{
				if (ws.Name.Equals("Overall PL", StringComparison.OrdinalIgnoreCase)) { continue; }
				// get budget sheet
				var budgetSheet = budget.Worksheets[ws.Name];
				if (budgetSheet == null) { LogBox(string.Format("No Cost Budget found for ({0})", ws.Name)); continue; }

				var wsrow = PROJ_OTHERS_ROW + 1;
				var budgetrow = Options.Default.CostBudgetRow_Proj + 1;

				// Check sheet
				if (!budgetSheet.Cells[budgetrow-1, Options.Default.CostBudgetACCol].GetValue<string>().Equals("Account Name", StringComparison.OrdinalIgnoreCase))
				{
					LogBox(string.Format("Didn't find \"Account Name\" in expected cell in cost budget for ({0}))", ws.Name));
					continue;
				}

				string budgetAccName = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetACCol].GetValue<string>();

				while (!IsEmptyCell(budgetAccName))
				{
					int irow = -1;
					try
					{
						irow = searchText(ws, wsrow, 1, budgetAccName);
					}
					catch (ArgumentException)
					{
						// Missing budget
						LogBox(string.Format("MISSING FROM COST ACTUAL ({0}) from ({1}).", budgetAccName, ws.Name));
					}
					if (irow > 0)
					{
						//LogBox("Summing: " + (Options.Default.CostBudgetYearCol + 1) + " - " + (Options.Default.CostBudgetYearCol + months));
						ws.Cells[irow, PROJ_BUDGET_COLUMN].Value = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetYearCol + 1, budgetrow, Options.Default.CostBudgetYearCol + months].Sum(cell => { return getDecimal(cell); });
						ws.Cells[irow, PROJ_BUDGET_COLUMN - 1].Value = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetYearCol].Value;
					}
					budgetrow = budgetrow + 1;
					budgetAccName = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetACCol].GetValue<string>();
				}
				//Format cells
				ws.Cells[PROJ_OTHERS_ROW, PROJ_BUDGET_COLUMN - 1, ws.Dimension.Rows, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = "$#,##0.00";
				ws.Cells[PROJ_OTHERS_ROW, PROJ_ACTUAL_COLUMN + 2, ws.Dimension.Rows, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = "0.00%";

			}
		}

		private static void processTimesheet(ExcelWorkbook wb, ExcelWorkbook timesheetwb, int months)
		{
			LogBox("Processing Timesheet\n");
			var igsheets = Options.Default.IgnoreSheets.Replace("\r\n", "\n").Split('\n').ToList();
			foreach (var ws in timesheetwb.Worksheets)
			{
				if (igsheets.Contains(ws.Name)) { continue; }
				// people sheets collate project hours
				var irow = 11;
				for (int i = 0; i < months; i++)
				{
					// get position for month
					var pos = ws.Cells[irow - 2, TIMESHEET_PROJ_COL - 1].GetValue<int>();
					var position = posIDmap.ContainsKey(pos) ? posIDmap[pos].Item1 : "";
					if (position.Equals(""))
					{
						LogBox(string.Format("Missing position for month ({0}) of ({1})", i, ws.Name));
						position = "N/A";
					}
					// check for projects
					for (int x = 0; x < 12; x++)
					{
						irow = irow + x;
						var project = ws.Cells[irow, TIMESHEET_PROJ_COL].GetValue<string>();
						if (!string.IsNullOrEmpty(project))
						{
							if (projTranslations.ContainsKey(project)) { project = projTranslations[project]; }

							if (!projPosHours.ContainsKey(project)) { projPosHours[project] = new Dictionary<string, decimal>(); }
							if (!projPosHours[project].ContainsKey(position)) { projPosHours[project][position] = 0; }
							projPosHours[project][position] = projPosHours[project][position] + 
								ws.Cells[irow, TIMESHEET_HOURS_TOTAL].GetValue<decimal>();
						} else { break; }
					}
					// jump to next
					irow = irow + 38;
				}
			}

		}

		private static void processProjectsTime(ExcelWorkbook wb, ExcelWorkbook budget, int months)
		{
			LogBox("Processing Project - Time\n");

			foreach (var ws in wb.Worksheets)
			{
				if (ws.Name.Equals("Overall PL", StringComparison.OrdinalIgnoreCase)) { continue; }
				// get budget sheet
				var budgetSheet = budget.Worksheets[ws.Name];
				if (budgetSheet == null) { LogBox(string.Format("No Time Budget found for ({0})", ws.Name)); continue; }

				var wsrow = PROJ_HOURS_ROW + 1;
				var budgetrow = Options.Default.TimeBudgetDateRow + 2;

				// Check sheet
				var test = budgetSheet.Cells[budgetrow - 2, Options.Default.TimeBudgetPosCol].GetValue<string>();
				if (test == null || !test.Trim().Equals("Expected days", StringComparison.OrdinalIgnoreCase))
				{
					if (test != null)
						LogBox(string.Format("Didn't find \"Expected days\" in expected cell in time budget for ({0}) found ({1}))", ws.Name, test.Trim()));
					else
						LogBox(string.Format("Didn't find \"Expected days\" in expected cell in time budget for ({0})", ws.Name));
					continue;
				}

				string position = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetPosCol].GetValue<string>();
				var posdone = new List<string>();

				Dictionary<string, decimal> projHours = null;
				if (projPosHours.ContainsKey(ws.Name))
				{
					projHours = projPosHours[ws.Name];
				}

				while (!position.Contains("Total"))
				{
					var sumToDate = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetYearCol, budgetrow, Options.Default.TimeBudgetYearCol + months-1].Sum(cell => { return getDecimal(cell); });
					var sumYear = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetYearCol, budgetrow, Options.Default.TimeBudgetYearCol + 11].Sum(cell => { return getDecimal(cell); });
					//ws.Cells[budgetrow, PROJ_BUDGET_COLUMN].Value
					if (sumToDate != 0 || sumYear != 0)
					{
						// insert position if values
						ws.InsertRow(wsrow, 1);
						ws.Cells[wsrow, PROJ_BUDGET_COLUMN - 1].Value = sumYear;
						ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Value = sumToDate;
						ws.Cells[wsrow, 1].Value = position;
						// Insert Actual time from timesheet mapping
						if (projHours != null)
						{
							if (projHours.ContainsKey(position))
							{
								ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Value = projHours[position];
								posdone.Add(position);
							}
						}
						wsrow = wsrow + 1;
					} else
					{
						LogBox(string.Format("No Time Budget for position ({0}) for project ({1})", position, ws.Name));
					}
					budgetrow = budgetrow + 1;
					position = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetPosCol].GetValue<string>();
				}
				// add missing actual positions from budget:
				if (projHours != null)
				{
					foreach (var pos in projHours)
					{
						if (posdone.Contains(pos.Key)) { continue; }
						ws.InsertRow(wsrow, 1);
						ws.Cells[wsrow, 1].Value = pos.Key;
						ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Value = pos.Value;
						wsrow = wsrow + 1;
					}
				}
				//Format cells
				ws.Cells[PROJ_HOURS_ROW, PROJ_BUDGET_COLUMN - 1, wsrow, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = "#,##0.00";
				ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN + 2, wsrow, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = "0.00%";

				var stoprow = wsrow;
				// Insert Staff Cost section
				wsrow = wsrow + 2;

				for (int row = PROJ_HOURS_ROW + 1; row < stoprow; row++)
				{
					ws.InsertRow(wsrow, 1);
					// insert cost and then formulas
					position = ws.Cells[row, 1].GetValue<string>();
					var rateCell = ws.Cells[wsrow, ACTUALCOLUMN + 3];
					if (posmap.ContainsKey(position))
					{
						rateCell.Value = posmap[position];
					}
					else {
						rateCell.Value = "N/A";
					}
					ws.Cells[wsrow, 1].Value = position;
					// formulas ws.Cells[rowNum, ACTUALCOLUMN + 1].Formula = "=" + acell.Address + "-" + ws.Cells[rowNum, BUDGETCOLUMN].Address;
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN-1].Formula = "=" + ws.Cells[row, PROJ_BUDGET_COLUMN-1].Address + "*" + rateCell.Address;
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Formula = "=" + ws.Cells[row, PROJ_BUDGET_COLUMN].Address + "*" + rateCell.Address;
					ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Formula = "=" + ws.Cells[row, PROJ_ACTUAL_COLUMN].Address + "*" + rateCell.Address;
					// var

					// increment things
					wsrow = wsrow + 1;
				}
				// format cells
				ws.Cells[stoprow + 2, PROJ_BUDGET_COLUMN - 1, wsrow, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = "$#,##0.00";
				ws.Cells[stoprow + 2, PROJ_ACTUAL_COLUMN + 2, wsrow, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = "0.00%";
			}
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
			initPositions();

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
			var timesheetwb = new ExcelPackage(new FileInfo(Options.Default.TimesheetFile)).Workbook;
			// delete old
			File.Delete(savefname);
			var pkg = new ExcelPackage(new FileInfo(savefname));
			
			var wb = pkg.Workbook;

			var overallws = setupWorkbook(wb, start, end);

			//Process timesheet
			processTimesheet(wb, timesheetwb, months);

			// DO OVERALL ACTUALS, THEN BUDGET
			addOverallData(overallws, start, end);
			processOverallCostBudget(overallws, costBudgetwb, months);

			// DO PROJECT ACTUAL COSTS THEN PROJET BUDGET
			processProjectActualCost(wb, start, end);
			processProjectsCostBudget(wb, costBudgetwb, months);

			// Do project time budget
			processProjectsTime(wb, timeBudgetwb, months);

			// autosize before save
			foreach (var ws in wb.Worksheets)
			{
				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				for (var ncol = 1; ncol <= ws.Dimension.Columns; ncol++)
				{
					ws.Column(ncol).Width = ws.Column(ncol).Width + 5;
				}
			}
			// move overheads to first index

			wb.Worksheets.MoveAfter("Overheads", "Overall PL");
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
			ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN].Value = "Actual"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 1].Value = "Var"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 2].Value = "Var %"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 3].Value = "Rate";
		}

		private static Dictionary<string, Tuple<string,decimal>> getBudgetEmployees(ExcelWorksheet bws, DateTime from)
		{
			var col = from.Month + BUDGETTIMECOLUMNSTART - 1;
			var emps = new Dictionary<string, Tuple<string, decimal>>();
			if (bws == null) { return emps; }

			var i = bws.Dimension.Rows;
			for (var rownum = 0; rownum < bws.Dimension.Rows; rownum++)
			{
				if (rownum == 1) { continue; } //increment rownum
				var emp = bws.Cells[rownum, 1].Value.ToString();
				if (!string.IsNullOrWhiteSpace(emp))
				{
					try {
						emps[emp] = new Tuple<string, decimal>(bws.Cells[rownum, 2].Value.ToString(), bws.Cells[rownum, col].GetValue<decimal>());
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
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Value = decimal.Parse(line);

					// insert formulas
					if (IncomeAccounts.Contains(ws.Cells[rownum, 1].Value))
					{
						ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 1].Formula = "=" + ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Address + "-" + ws.Cells[rownum, PROJ_BUDGET_COLUMN].Address;
						ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 2].Formula = "=" + ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Address + "/" + ws.Cells[rownum, PROJ_BUDGET_COLUMN].Address;
					}
					else
					{
						ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 1].Formula = "=" + ws.Cells[rownum, PROJ_BUDGET_COLUMN].Address + "-" + ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Address;
						ws.Cells[rownum, PROJ_ACTUAL_COLUMN + 2].Formula = "=" + ws.Cells[rownum, PROJ_BUDGET_COLUMN].Address + "/" + ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Address;
					}

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
