using PM = Xero.Api.Payroll.Australia.Model;
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

		static List<string> IncomeAccounts = new List<string>();
		static List<string> HideAccounts = new List<string>();

		static AustralianPayroll ap = null;
		static Core c = null;

		static Dictionary<int, Tuple<string, decimal>> posIDmap = new Dictionary<int, Tuple<string, decimal>>();
		static Dictionary<string, Tuple<int, decimal>> posmap = new Dictionary<string, Tuple<int, decimal>>();
		static Dictionary<string, Dictionary<string, decimal>> projPosHours = new Dictionary<string, Dictionary<string,decimal>>();
		static Dictionary<string, string> projCollation = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		static List<string> overheadProjs = new List<string>();

		static Dictionary<string, string> mergeAccounts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // alias -> rename

		static BackgroundWorker worker = null;
		static StreamWriter logfile = null;

		static int BUDGETCOLUMN = 3;
		static int ACTUALCOLUMN = 4;
		static int OVERALLROWSTART = 4;
		static int PROJ_BUDGET_COLUMN = BUDGETCOLUMN;
		static int PROJ_ACTUAL_COLUMN = ACTUALCOLUMN;
		static int PROJ_HOURS_ROW = 4;
		static int PROJ_COST_ROW = 6;
		static int PROJ_OTHERS_ROW = 8;
		static int TIMESHEET_HOURS_TOTAL = 38; // AL total column
		static int TIMESHEET_PROJ_COL = 5;
		static List<string> TIMESHEET_BAD_PROJS = new List<string>() { "PROJECT", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER", "JANUARY", "FEBUARY", "MARCH", "APRIL", "MAY"};

		// static int BUDGETTIMEROWSTART = 3; Progmatically find
		//static int BUDGETTIMECOLSEARCH = 2;
		//static string BUDGETTIMESSEARCH = "Total Days YTD";
		static int BUDGETTIMECOLUMNSTART = 3;
		//_-\$* #,##0_-;-\$* #,##0_-;_-\$* " -"_-;_-@_-
		static string CURRENCY_FORMAT = "_-\\$* #,##0_-;[Red]_\\$* (#,##0)_-;_-\\$* \" -\"_-;_-@_-";
		//static string CURRENCY_FORMAT = "$#,##0";
		static string NUMBER_FORMAT = "0;[Red]-0;[Black] \\- ; \\- ";
		static string PERCENT_FORMAT = "0%;[Red]-0%;[Black]0%; \\- ";
		static string DATE_FORMAT = "d/mm/yyyy;@";
		//static string DATE_FORMAT = "dd/mm/yyyy;@";

		static string TOTAL_STAFF_DAYS = "Total staff days";
		static string TOTAL_STAFF_COST = "Total staff cost";
		static string FULL_YEAR = "FY{0}/{1} Budget";
		static string GNET_PROFIT = "Net Profit";
		static string NET_PROFIT = "Net Profit";
		static string NET_COST_TITLE = "Net Profit (inc. Staff Cost)";

		static Guid ProjectTCID;

		internal static void LogBox(string msg)
		{
			if (worker != null)
				worker.ReportProgress(0, msg+"\n");
			if (logfile != null)
			{
				logfile.WriteLine(msg);
				logfile.Flush();
				logfile.BaseStream.Flush();
			}
		}

		private static bool initXero()
		{
			// setup log file
			if (logfile == null)
				logfile = new StreamWriter(@"log.txt", true);

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

		private static void initOthers()
		{
			projPosHours = new Dictionary<string, Dictionary<string, decimal>>();

			// init collation groups
			projCollation = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			projCollation["Unassigned"] = "Overheads";
			using (StringReader reader = new StringReader(Options.Default.Collation))
			{
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					// split line up for collated projects
					line = line.Trim();
					if (string.IsNullOrWhiteSpace(line)) { continue; }

					var ls = line.Split(',');
					if (ls.Count() < 2)
					{
						LogBox(string.Format("Not enough items for collation ({0})", line));
					}
					else
					{
						foreach (var item in ls)
						{
							projCollation[item.Trim()] = ls[ls.Count()-1].Trim();
						}
					}
				}
			}
			// init position tables
			posIDmap = new Dictionary<int, Tuple<string, decimal>>();
			posmap = new Dictionary<string, Tuple<int, decimal>>();
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
							posmap[pos] = new Tuple<int, decimal>(id, cost);
						}
					}
				}
			}
			// init overhead projects
			overheadProjs = new List<string>();
			using (StringReader reader = new StringReader(Options.Default.OverheadProjs))
			{
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					line = line.Trim();
					if(!overheadProjs.Contains(line)) { overheadProjs.Add(line); }
				}
			}
			// init hide accounts
			HideAccounts = new List<string>();
			using (StringReader reader = new StringReader(Options.Default.HideAccounts))
			{
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					line = line.Trim();
					if (!HideAccounts.Contains(line)) { HideAccounts.Add(line); }
				}
			}

			// init collation accounts
			mergeAccounts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			if (!string.IsNullOrWhiteSpace(Options.Default.NetProfitLabel))
			{
				mergeAccounts["Net Profit"] = Options.Default.NetProfitLabel;
				mergeAccounts[Options.Default.NetProfitLabel] = Options.Default.NetProfitLabel;
				NET_PROFIT = Options.Default.NetProfitLabel;
				NET_COST_TITLE = NET_PROFIT + " (inc. Staff Cost)";
			}
			else
			{
				NET_PROFIT = "Net Profit";
				NET_COST_TITLE = "Net Profit (inc. Staff Cost)";
			}

			using (StringReader reader = new StringReader(Options.Default.MergeAccounts))
			{
				string line;
				while ((line = reader.ReadLine()) != null)
				{
					// split line up for merged accounts
					line = line.Trim();
					if (string.IsNullOrWhiteSpace(line)) { continue; }

					var ls = line.Split(',');
					if (ls.Count() < 2)
					{
						LogBox(string.Format("Not enough items for account merge ({0})", line));
					}
					else
					{
						foreach (var item in ls)
						{
							mergeAccounts[item.Trim()] = ls[ls.Count() - 1].Trim();
						}
					}
				}
			}
		}

		private static ExcelWorksheet setupWorkbook(ExcelWorkbook wb, DateTime from, DateTime to)
		{
			var ws = wb.Worksheets.Add("Overall PL");
			ws.Cells["A1"].Value = "Profit and loss"; ws.Cells["C1"].Value = "Begining"; ws.Cells["D1"].Value = "Ending";
			ws.Cells["A2"].Value = Options.Default.OrganisationName; ws.Cells["C2"].Value = from; ws.Cells["D2"].Value = to;

			ws.Cells[OVERALLROWSTART, 1].Value = "Accounts"; ws.Cells[OVERALLROWSTART, BUDGETCOLUMN-1].Value = string.Format(FULL_YEAR, from.Year % 100, (from.Year % 100) + 1);
			ws.Cells[OVERALLROWSTART, BUDGETCOLUMN].Value = "Budget YTD"; ws.Cells[OVERALLROWSTART, ACTUALCOLUMN].Value = "Actual";
			ws.Cells[OVERALLROWSTART, ACTUALCOLUMN+1].Value = "Var $"; ws.Cells[OVERALLROWSTART, ACTUALCOLUMN+2].Value = "Var %";
			ws.Cells[OVERALLROWSTART, ACTUALCOLUMN + 3].Value = "Notes";

			// Format
			ws.Cells["A1:D1"].Style.Font.Bold = true; ws.Cells["A2"].Style.Font.Bold = true; ws.Cells["4:4"].Style.Font.Bold = true;
			ws.Cells["C2"].Style.Numberformat.Format = DATE_FORMAT; ws.Cells["D2"].Style.Numberformat.Format = DATE_FORMAT;
			return ws;
		}

		private static void addOverallData(ExcelWorksheet ws, DateTime start, DateTime end)
		{
			// Overall profit and loss report
			var plReport = c.Reports.ProfitAndLoss(start, from: start, to: end, standardLayout: true);
			LogBox("Getting Overall PL\n");

			var rowNum = OVERALLROWSTART + 1;

			var incomelist = new List<string>();
			var incomeitems = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
			var expenselist = new List<string>();
			var expenseitems = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

			Tuple<string, decimal> netprofit = null;

			// split items in to income and expense
			foreach (var row in plReport.Rows)
			{
				if (row.Cells == null)
				{
					foreach (var trow in row.Rows)
					{
						if (trow.Cells != null)
						{
							// get name and value, cell[0] and [1]
							var acct = trow.Cells[0].Value;
							if (HideAccounts.Contains(acct)) { continue; }
							var value = decimal.Parse(trow.Cells[1].Value);

							// check if merging accounts
							if (mergeAccounts.ContainsKey(acct))
							{
								acct = mergeAccounts[acct];
							}
							decimal item = 0;
							// income or expense?
							if (acct.Equals(NET_PROFIT, StringComparison.OrdinalIgnoreCase) || acct.Equals(GNET_PROFIT, StringComparison.OrdinalIgnoreCase))
							{
								netprofit = new Tuple<string, decimal>(acct, value);
							}
							else if (IncomeAccounts.Contains(acct))
							{
								incomeitems.TryGetValue(acct, out item);
								incomeitems[acct] = item + value;
								if (!incomelist.Contains(acct))
								{
									if (acct.StartsWith("Total", StringComparison.OrdinalIgnoreCase))
										incomelist.Add(acct);
									else
										incomelist.Insert(0, acct);
								}
							}
							else
							{
								expenseitems.TryGetValue(acct, out item);
								expenseitems[acct] = item + value;
								if (!expenselist.Contains(acct))
								{
									if (acct.StartsWith("Total", StringComparison.OrdinalIgnoreCase))
										expenselist.Add(acct);
									else
										expenselist.Insert(0, acct);
								}
							}
						}
					}
				}
			}

			foreach (var acct in incomelist)
			{
				ws.Cells[rowNum, 1].Value = acct;
				ws.Cells[rowNum, ACTUALCOLUMN].Value = incomeitems[acct];
				ws.Cells[rowNum, BUDGETCOLUMN - 1, rowNum, BUDGETCOLUMN].Value = 0; // fill in missing actual

				// insert formulas
				addFormula(ws, acct, rowNum, ACTUALCOLUMN, BUDGETCOLUMN);
				rowNum++;
			}
			ws.Cells[rowNum, 1].Value = "Operating Expenses";
			rowNum++;
			foreach (var acct in expenselist)
			{
				ws.Cells[rowNum, 1].Value = acct;
				ws.Cells[rowNum, ACTUALCOLUMN].Value = expenseitems[acct];
				ws.Cells[rowNum, BUDGETCOLUMN - 1, rowNum, BUDGETCOLUMN].Value = 0; // fill in missing actual

				// insert formulas
				addFormula(ws, acct, rowNum, ACTUALCOLUMN, BUDGETCOLUMN);
				rowNum++;
			}
			if (netprofit != null)
			{
				ws.Cells[rowNum, 1].Value = netprofit.Item1;
				ws.Cells[rowNum, ACTUALCOLUMN].Value = netprofit.Item2;
				ws.Cells[rowNum, BUDGETCOLUMN - 1, rowNum, BUDGETCOLUMN].Value = 0; // fill in missing actual
				addFormula(ws, null, rowNum, ACTUALCOLUMN, BUDGETCOLUMN, inc: true);
				rowNum++;
			}
		}

		private static bool IsEmptyCell(string data)
		{
			return string.IsNullOrWhiteSpace(data);
		}

		private static int searchText(ExcelWorksheet ws, int rowi, int coli, string text, bool reverse = false, bool skipempty = false, bool startswith = false)
		{
			var row = rowi;
			var i = 0;
			string data = ws.Cells[row, coli].GetValue<string>();
			bool notbail;
			// loop condition
			if (skipempty == true)
				notbail = i < 300;
			else
				notbail = !IsEmptyCell(data);
			while (notbail)
			{
				if (!startswith)
				{
					if (data != null && data.Equals(text, StringComparison.OrdinalIgnoreCase))
						return row;
				}
				else
				{
					if (data != null && data.StartsWith(text, StringComparison.OrdinalIgnoreCase))
						return row;
				}

				if (reverse)
					row--;
				else
					row++;
				data = ws.Cells[row, coli].GetValue<string>();
				// loop condition
				if (skipempty == true)
					notbail = i < 300;
				else
					notbail = !IsEmptyCell(data);
				i++;
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
			int budgetrow;
			try {
				budgetrow = searchText(budgetSheet, 1, Options.Default.CostBudgetAccCol, Options.Default.CostBudgetACHeader, skipempty: true) + 1;
			} catch (ArgumentException)
			{
				throw new ArgumentException("Couldn't find Account name header text for overall budget");
			}

			var mergedValue = new Dictionary<int, Tuple<decimal,decimal>>(); // row -> budgetfull, budgetytd

			string budgetAccName = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetAccCol].GetValue<string>();
			var missingActuals = new Dictionary<string, Tuple<decimal, decimal>>(StringComparer.OrdinalIgnoreCase);
			int irow;

			while (!IsEmptyCell(budgetAccName))
			{
				var merge = false;
				if (mergeAccounts.ContainsKey(budgetAccName))
				{
					budgetAccName = mergeAccounts[budgetAccName];
					merge = true;
				}
				irow = -1;
				try
				{
					irow = searchText(ws, rownum, 1, budgetAccName);
				} catch (ArgumentException)
				{
					// Missing actual in budget
					// LogBox(string.Format("MISSING FROM ACTUAL ({0}). Merging ({1})", budgetAccName, merge));
				}
				var budgetytd = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetYearCol + 1, budgetrow, Options.Default.CostBudgetYearCol + months].Sum(cell => { return getDecimal(cell); });
				var budgetfull = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetYearCol].GetValue<decimal>();
				if (irow > 0) {
					//LogBox("Summing: " + (Options.Default.CostBudgetYearCol + 1) + " - " + (Options.Default.CostBudgetYearCol + months));
					if (merge)
					{
						if (mergedValue.ContainsKey(irow))  // sum values
							mergedValue[irow] = new Tuple<decimal, decimal>(mergedValue[irow].Item1 + budgetfull, mergedValue[irow].Item2 + budgetytd);
						else  // new values
							mergedValue[irow] = new Tuple<decimal, decimal>(budgetfull, budgetytd);
					}
					else {
						ws.Cells[irow, BUDGETCOLUMN].Value = budgetytd;
						ws.Cells[irow, BUDGETCOLUMN - 1].Value = budgetfull;
					}
				} else
				{
					// build list of budgets with missing actuals to prepend at top of report
					if (missingActuals.ContainsKey(budgetAccName))  // sum values
						missingActuals[budgetAccName] = new Tuple<decimal, decimal>(missingActuals[budgetAccName].Item1 + budgetfull, missingActuals[budgetAccName].Item2 + budgetytd);
					else  // new values
						missingActuals[budgetAccName] = new Tuple<decimal, decimal>(budgetfull, budgetytd);
				}
				budgetrow = budgetrow + 1;
				budgetAccName = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetAccCol].GetValue<string>();
			}
			// process merged accounts
			foreach (var kvpair in mergedValue)
			{
				rownum = kvpair.Key;
				var budgetfull = kvpair.Value.Item1;
				var budgetytd = kvpair.Value.Item2;
				ws.Cells[rownum, BUDGETCOLUMN].Value = budgetytd;
				ws.Cells[rownum, BUDGETCOLUMN - 1].Value = budgetfull;
			}
			// process budgets with missing actuals split between income and expense
			rownum = OVERALLROWSTART + 1;
			foreach (var kvpair in missingActuals)
			{
				if (IncomeAccounts.Contains(kvpair.Key))
					irow = rownum;
				else
					irow = searchText(ws, rownum, 1, "Operating Expenses") + 1;
				ws.InsertRow(irow, 1);
				ws.Cells[irow, 1].Value = kvpair.Key;
				ws.Cells[irow, BUDGETCOLUMN].Value = kvpair.Value.Item2;
				ws.Cells[irow, BUDGETCOLUMN - 1].Value = kvpair.Value.Item1;
				ws.Cells[irow, ACTUALCOLUMN].Value = 0; // fill in missing actual
				addFormula(ws, kvpair.Key, irow, ACTUALCOLUMN, BUDGETCOLUMN);
			}
			//Format cells
			ws.Cells[OVERALLROWSTART + 1, BUDGETCOLUMN - 1, ws.Dimension.Rows, ACTUALCOLUMN + 1].Style.Numberformat.Format = CURRENCY_FORMAT;
			ws.Cells[OVERALLROWSTART + 1, ACTUALCOLUMN +2, ws.Dimension.Rows, ACTUALCOLUMN + 2].Style.Numberformat.Format = PERCENT_FORMAT;
		}

		private static void formatAccCells(ExcelWorksheet ws, int row = 0)
		{
			if (row > 0)
			{
				ws.Cells[row, PROJ_BUDGET_COLUMN - 1, row, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = CURRENCY_FORMAT;
				ws.Cells[row, PROJ_ACTUAL_COLUMN + 2, row, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = PERCENT_FORMAT;
			}
			else {
				ws.Cells[PROJ_OTHERS_ROW + 1, PROJ_BUDGET_COLUMN - 1, ws.Dimension.Rows, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = CURRENCY_FORMAT;
				ws.Cells[PROJ_OTHERS_ROW + 1, PROJ_ACTUAL_COLUMN + 2, ws.Dimension.Rows, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = PERCENT_FORMAT;
			}
		}

		private static void processProjectsCostBudget(ExcelWorkbook wb, ExcelWorkbook budget, int months, DateTime from, DateTime to)
		{
			LogBox("Processing Project Budgets - Cost\n");
			var doneProjs = new List<string>();

			var sheets = wb.Worksheets.Select(x => x.Name).ToList();
			int index = 0;
			bool passone = true;

			while (index < sheets.Count)
			{
				var ws = wb.Worksheets[sheets[index]];
				if (ws.Name.Equals("Overall PL", StringComparison.OrdinalIgnoreCase)) { doneProjs.Add(ws.Name); index++; continue; }
				// get budget sheet
				var budgetSheet = budget.Worksheets[ws.Name];
				if (budgetSheet == null) { LogBox(string.Format("No Cost Budget found for ({0})", ws.Name)); index++; formatAccCells(ws); continue; }

				var wsrow = PROJ_OTHERS_ROW + 1;
				int budgetrow;
				try
				{
					budgetrow = searchText(budgetSheet, 1, Options.Default.CostBudgetAccCol, Options.Default.CostBudgetACHeader, skipempty: true) + 1;
				}
				catch (ArgumentException)
				{
					throw new ArgumentException(string.Format("Couldn't find Account name header text for project: {0}", ws.Name));
				}

				// Check sheet
				if (!budgetSheet.Cells[budgetrow-1, Options.Default.CostBudgetAccCol].GetValue<string>().Equals("Account Name", StringComparison.OrdinalIgnoreCase))
				{
					throw new ArgumentException(string.Format("ERROR: Didn't find \"Account Name\" in expected cell in cost budget for ({0}))", ws.Name));
				}

				var mergedValue = new Dictionary<int, Tuple<decimal, decimal>>(); // row -> budgetfull, budgetytd
				var missingActuals = new Dictionary<string, Tuple<decimal, decimal>>(StringComparer.OrdinalIgnoreCase);
				Tuple<string, decimal, decimal> netprofit = null;

				string budgetAccName = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetAccCol].GetValue<string>();
				int irow;
				while (!IsEmptyCell(budgetAccName))
				{
					//LogBox("Premerge: (" + budgetAccName + ")");
					var merge = false;
					if (mergeAccounts.ContainsKey(budgetAccName))
					{
						budgetAccName = mergeAccounts[budgetAccName];
						merge = true;
					}
					//LogBox("[" + budgetSheet.Name + "]" + "Processing Budget acc: (" + budgetAccName + ")" + " merge: " + merge);
					irow = -1;
					try
					{
						irow = searchText(ws, wsrow, 1, budgetAccName);
					}
					catch (ArgumentException)
					{
						// Missing from actual
						//LogBox(string.Format("MISSING FROM COST ACTUAL ({0}) from ({1}).", budgetAccName, ws.Name));
					}
					var budgetfull = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetYearCol].GetValue<decimal>();
					var budgetytd = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetYearCol + 1, budgetrow, Options.Default.CostBudgetYearCol + months].Sum(cell => { return getDecimal(cell); });
					//LogBox("Full: " + budgetfull + " YTD: " + budgetytd);

					//LogBox(string.Format("Proj ({0}), Acc ({1}), full ({2})", ws.Name, budgetAccName, budgetfull));
					if (irow > 0)
					{
						if (merge)
						{
							if (mergedValue.ContainsKey(irow))  // sum values
								mergedValue[irow] = new Tuple<decimal, decimal>(mergedValue[irow].Item1 + budgetfull, mergedValue[irow].Item2 + budgetytd);
							else  // new values
								mergedValue[irow] = new Tuple<decimal, decimal>(budgetfull, budgetytd);
						}
						else {
							// if not merging insert to found item
							//LogBox("irow: " + irow);
							ws.Cells[irow, PROJ_BUDGET_COLUMN].Value = budgetytd;
							ws.Cells[irow, PROJ_BUDGET_COLUMN - 1].Value = budgetfull;
						}
					} else
					{
						// build list of budgets with missing actuals to prepend at top of report
						if (budgetAccName.Equals(NET_PROFIT, StringComparison.OrdinalIgnoreCase) || budgetAccName.Equals(GNET_PROFIT, StringComparison.OrdinalIgnoreCase))
							netprofit = new Tuple<string, decimal, decimal>(budgetAccName, budgetfull, budgetytd);
						else if (missingActuals.ContainsKey(budgetAccName))  // sum values
							missingActuals[budgetAccName] = new Tuple<decimal, decimal>(missingActuals[budgetAccName].Item1 + budgetfull, missingActuals[budgetAccName].Item2 + budgetytd);
						else  // new values
							missingActuals[budgetAccName] = new Tuple<decimal, decimal>(budgetfull, budgetytd);
						
					}
					budgetrow = budgetrow + 1;
					budgetAccName = budgetSheet.Cells[budgetrow, Options.Default.CostBudgetAccCol].GetValue<string>();
				}

				// process merged accounts
				foreach (var kvpair in mergedValue)
				{
					wsrow = kvpair.Key;
					var budgetfull = kvpair.Value.Item1;
					var budgetytd = kvpair.Value.Item2;
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Value = budgetytd;
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN - 1].Value = budgetfull;
				}
				wsrow = PROJ_OTHERS_ROW + 1;
				// process budgets with missing actuals
				foreach (var kvpair in missingActuals)
				{
					//LogBox("dumping: " + kvpair.Key);
					if (IncomeAccounts.Contains(kvpair.Key))
					{
						irow = wsrow;
						if (kvpair.Key.StartsWith("Total"))
						{
							try { irow = searchText(ws, wsrow, 1, "Operating Expenses") + 1; }
							catch (ArgumentException)
							{
								ws.Cells[ws.Dimension.Rows + 1, 1].Value = "Operating Expenses";
								irow = searchText(ws, wsrow, 1, "Operating Expenses") + 1;
							}
							irow = irow - 1;
						}
					}
					else
					{
						try { irow = searchText(ws, wsrow, 1, "Operating Expenses") + 1; }
						catch (ArgumentException)
						{
							ws.Cells[ws.Dimension.Rows + 1, 1].Value = "Operating Expenses";
							irow = searchText(ws, wsrow, 1, "Operating Expenses") + 1;
						}
						if (kvpair.Key.StartsWith("Total"))
						{
							irow = ws.Dimension.Rows;
						}
					}
					//LogBox("dumping: " + kvpair.Key + " in: " + ws.Cells[irow, 1].Address);
					ws.InsertRow(irow, 1);
					ws.Cells[irow, 1].Value = kvpair.Key;
					ws.Cells[irow, PROJ_BUDGET_COLUMN].Value = kvpair.Value.Item2;
					ws.Cells[irow, PROJ_BUDGET_COLUMN - 1].Value = kvpair.Value.Item1;
					ws.Cells[irow, PROJ_ACTUAL_COLUMN].Value = 0; // fill in missing actual
					addFormula(ws, kvpair.Key, irow, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);
				}
				if (netprofit != null)
				{
					irow = ws.Dimension.Rows + 1;
					ws.Cells[irow, 1].Value = netprofit.Item1;
					ws.Cells[irow, PROJ_ACTUAL_COLUMN].Value = netprofit.Item2;
					ws.Cells[irow, PROJ_BUDGET_COLUMN - 1, irow, PROJ_BUDGET_COLUMN].Value = 0; // fill in missing actual
					addFormula(ws, netprofit.Item1, irow, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);
					irow++;
				}
				//Format cells
				formatAccCells(ws);

				doneProjs.Add(ws.Name);
				if (passone && index == sheets.Count-1)
				{
					// add budget sheets
					foreach (var bs in budget.Worksheets)
					{
						if (doneProjs.Contains(bs.Name) || Options.Default.IgnoreBudgetSheets.Contains(bs.Name)) { continue; }
						ws = wb.Worksheets.Add(bs.Name);
						setupProjectSheet(ws, bs.Name, from, to);
						sheets.Add(bs.Name);
						LogBox(string.Format("WARNING: Have Budget for project ({0}) but no actuals.", bs.Name));
					}

					passone = false;
				}
				index++;
			}
		}

		private static void processTimesheet(ExcelWorkbook wb, ExcelWorkbook timesheetwb, int months)
		{
			LogBox("Processing Timesheet\n");
			var igsheets = Options.Default.IgnoreSheets.Replace("\r\n", "\n").Split('\n').ToList();
			var posSheet = timesheetwb.Worksheets["Positions"];

			foreach (var ws in timesheetwb.Worksheets)
			{
				if (igsheets.Contains(ws.Name)) { continue; }
				// people sheets collate project hours
				// get person name
				var name = ws.Cells["F3"].GetValue<string>();
				if (string.IsNullOrEmpty(name))
				{
					LogBox("ERROR: Missing name in timesheet for ({0}). Expect name in F3.");
				}
                // get position column
                ExcelRangeBase poscell;


                try
                {
                    poscell = (from cell in posSheet.Cells["1:1"]
                               where cell.GetValue<string>().Equals(name)
                               select cell).Single();
                }
                catch (Exception e)
                {
                    LogBox(string.Format("Cannot find staff member in Positions. Sheet: ({0}), Name: ({1})" +
                        "\n {2}", ws.Name, name, e));
                    throw;
                }
				var irow = 11;
				for (int i = 1; i <= months; i++)
				{
					// get position for month
					var pos = poscell.Offset(i, 0).GetValue<int>();
					//LogBox("Pos: (" + poscell.First().Address + ") Name: (" + name + ") pos: (" + pos + ") ");
					var position = posIDmap.ContainsKey(pos) ? posIDmap[pos].Item1 : "";
					if (position.Equals(""))
					{
						LogBox(string.Format("Missing position data for month ({0}) of ({1}). Check Positions sheet in Timesheet, and make sure there is an entry in the program in the Timesheets tab for ({2}).", i, ws.Name, pos));
						position = "N/A";
					}
					// check for projects
					for (int x = 0; x < 15; x++)
					{
						//LogBox("I: " + i + " X: " + x + " ROW: " + irow);
						var project = ws.Cells[irow, TIMESHEET_PROJ_COL].GetValue<string>();
						if (!string.IsNullOrEmpty(project))
						{
							// sanity check
							if (TIMESHEET_BAD_PROJS.Contains(project))
							{
								LogBox(string.Format("ERROR: Cannot sync to Timesheet rows. \"Rows between months\" is likely set incorrectly or sheet has wrong layout." +
                                    "\nSheet: {0}, Name: {1}, Data: {2}", ws.Name, name, project));
								return;
							}
							if (overheadProjs.Contains(project))
								project = "Overheads";
							else if (projCollation.ContainsKey(project))
								project = projCollation[project];

							if (!projPosHours.ContainsKey(project)) { projPosHours[project] = new Dictionary<string, decimal>(); }
							if (!projPosHours[project].ContainsKey(position)) { projPosHours[project][position] = 0; } // LogBox(string.Format("ADDING POSITION {0} TO {1}", position, project));
							projPosHours[project][position] = projPosHours[project][position] + 
								ws.Cells[irow, TIMESHEET_HOURS_TOTAL].GetValue<decimal>();
						}
						irow++;
					}
					// jump to next
					irow = irow + Options.Default.MonthRows;
				}
			}

		}

		private static void processProjectsTime(ExcelWorkbook wb, ExcelWorkbook budget, int months)
		{
			LogBox("Processing Project - Time Actuals and Budget\n");
			var hoursinday = Options.Default.HoursDay;

			var doneProjs = new List<string>();

			var worksheets = wb.Worksheets.ToList();
			//place Overheads last
			worksheets.Remove(wb.Worksheets["Overheads"]);
			worksheets.Add(wb.Worksheets["Overheads"]);

			foreach (var ws in worksheets)
			{
				if (ws == null) { continue; }
				if (ws.Name.Equals("Overall PL", StringComparison.OrdinalIgnoreCase)) { continue; }
				//prepare for overheads:
				if (ws.Name.Equals("Overheads"))
				{
					// Warn for projects not processed:
					foreach (var proj in projPosHours.Keys)
					{
						if (proj.Equals("Overheads", StringComparison.OrdinalIgnoreCase)) { continue; }
						if (!doneProjs.Contains(proj))
						{
							LogBox(string.Format("WARNING: Project time from timesheet ({0}) not allocated as Overheads but also not a project in Xero/budget OR Collated. Will be assigned to Overheads.", proj));
							var overheads = projPosHours["Overheads"];
							foreach (var x in projPosHours[proj])
							{
								if (overheads.ContainsKey(x.Key))  // sum
								{
									overheads[x.Key] = overheads[x.Key] + x.Value;
								} else
								{
									overheads[x.Key] = x.Value;
								}
							}
						}
					}
				}
				
				// get budget sheet
				var budgetSheet = budget.Worksheets[ws.Name];

				Dictionary<string, decimal> projHours = null;
				if (projPosHours.ContainsKey(ws.Name))
				{
					projHours = projPosHours[ws.Name];
					doneProjs.Add(ws.Name);
				}
				var posdone = new List<string>();
				// dictionary of position : ID, position, actualhours, budgetfullhours, budgetYTD
				Dictionary<string, Tuple<int, string, decimal, decimal, decimal>> mapping = new Dictionary<string, Tuple<int, string, decimal, decimal, decimal>>();

				var wsrow = PROJ_HOURS_ROW + 1;

				var budgetskip = false;
				if (budgetSheet == null) { LogBox(string.Format("No Time Budget found for ({0})", ws.Name)); budgetskip = true; }
				if (!budgetskip)
				{
					int budgetrow;
					int budgetrowStop;
					try
					{
						budgetrow = searchText(budgetSheet, 1, Options.Default.TBSearchCol, Options.Default.TBStartRowText, skipempty: true) + 1;
					}
					catch (ArgumentException)
					{
						throw new ArgumentException(string.Format("Couldn't find Time Budget Start Row Search text for project: {0}", ws.Name));
					}
					try
					{
						budgetrowStop = searchText(budgetSheet, 1, Options.Default.TBSearchCol, Options.Default.TBEndRowText, skipempty: true);
					}
					catch (ArgumentException)
					{
						throw new ArgumentException(string.Format("Couldn't find Time Budget End Row Search text for project: {0}", ws.Name));
					}

					// Check sheet
					var test = budgetSheet.Cells[budgetrow - 2, Options.Default.TimeBudgetPosCol].GetValue<string>();
					if (test == null || !test.Trim().Equals("Expected days", StringComparison.OrdinalIgnoreCase))
					{
						if (test != null)
							LogBox(string.Format("Didn't find \"Expected days\" in expected cell in time budget for ({0}) found ({1}))", ws.Name, test.Trim()));
						else
							LogBox(string.Format("Didn't find \"Expected days\" in expected cell in time budget for ({0})", ws.Name));
						budgetskip = true;
					}

					if (!budgetskip)
					{
						string position;

						for (; budgetrow < budgetrowStop; budgetrow++)
						{
							position = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetPosCol].GetValue<string>();
							var sumToDate = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetYearCol, budgetrow, Options.Default.TimeBudgetYearCol + months - 1].Sum(cell => { return getDecimal(cell); });
							var sumYear = budgetSheet.Cells[budgetrow, Options.Default.TimeBudgetYearCol, budgetrow, Options.Default.TimeBudgetYearCol + 11].Sum(cell => { return getDecimal(cell); });
							//ws.Cells[budgetrow, PROJ_BUDGET_COLUMN].Value
							if (sumToDate != 0 || sumYear != 0)
							{
								if (position == null)
								{
									throw new ArgumentException(string.Format("ERROR: Empty position name in Time Budget with allocated hours. Project ({0}).", ws.Name));
								}
								// insert position if values
								decimal actual = 0;
								if (projHours != null && projHours.ContainsKey(position))
								{
									actual = projHours[position];
									posdone.Add(position);
								}
								// actualhours, budgetfullhours, budgetYTD
								Tuple<int, decimal> positem = null;
								var got = posmap.TryGetValue(position, out positem);
								if (!got) {
									positem = new Tuple<int, decimal>(-100, 0);
									LogBox(string.Format("WARNING: Position ({0}) not entered in Program -> Timesheets -> Positions. Maybe a typo?", position));
								}
								mapping[position] = new Tuple<int, string, decimal, decimal, decimal>(positem.Item1, position, actual, sumYear, sumToDate);
								 
							}
							else
							{
								//LogBox(string.Format("No Time Budget for position ({0}) for project ({1})", position, ws.Name));
							}
						}
					}
				}
				// add missing actual positions from budget:
				if (projHours != null)
				{
					foreach (var pos in projHours)
					{
						if (posdone.Contains(pos.Key)) { continue; }
						if (pos.Key.Equals("N/A") && !(pos.Value > 0)) { continue; } // skip "N/A" if zero value
						if (pos.Key.Equals("N/A"))
						{
							mapping[pos.Key] = new Tuple<int, string, decimal, decimal, decimal>(-999, pos.Key, pos.Value, 0, 0);
						} else
						{
							Tuple<int, decimal> positem = null;
							var got = posmap.TryGetValue(pos.Key, out positem);
							if (!got)
							{
								positem = new Tuple<int, decimal>(-100, 0);
								LogBox(string.Format("WARNING: Position ({0}) not entered in Program -> Timesheets -> Positions. Maybe a typo?", pos.Key));
							}
							mapping[pos.Key] = new Tuple<int, string, decimal, decimal, decimal>(positem.Item1, pos.Key, pos.Value, 0, 0);
						}
					}
				}
				//insert sorted position values
				var values = mapping.Values.ToList();
				values.Sort((x, y) => y.Item1.CompareTo(x.Item1));
				foreach (var item in values)
				{
					// dictionary of position : ID, position, actualhours, budgetfullhours, budgetYTD
					ws.InsertRow(wsrow, 1);
					ws.Cells[wsrow, 1].Value = item.Item2;
					// DIVIDE BY HOURS IN DAY
					ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Value = item.Item3 / hoursinday;
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN - 1].Value = item.Item4;
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Value = item.Item5;

					// INSERT Variance formulas, budget - actual, variance / budget
					addFormula(ws, null, wsrow, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);

					wsrow = wsrow + 1;
				}

				// Totals for staff days
				ws.InsertRow(wsrow, 1);
				bool entries = PROJ_HOURS_ROW + 1 < wsrow - 1;
				if (entries)
				{
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN - 1].Formula = "=SUM(" + ws.Cells[PROJ_HOURS_ROW + 1, PROJ_BUDGET_COLUMN - 1].Address + ":" + ws.Cells[wsrow - 1, PROJ_BUDGET_COLUMN - 1].Address + ")";
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Formula = "=SUM(" + ws.Cells[PROJ_HOURS_ROW + 1, PROJ_BUDGET_COLUMN].Address + ":" + ws.Cells[wsrow - 1, PROJ_BUDGET_COLUMN].Address + ")";
					ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Formula = "=SUM(" + ws.Cells[PROJ_HOURS_ROW + 1, PROJ_ACTUAL_COLUMN].Address + ":" + ws.Cells[wsrow - 1, PROJ_ACTUAL_COLUMN].Address + ")";
				}
				ws.Cells[wsrow, 1].Value = TOTAL_STAFF_DAYS;
				addFormula(ws, null, wsrow, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);

				//Format cells
				ws.Cells[PROJ_HOURS_ROW + 1, PROJ_BUDGET_COLUMN - 1, wsrow + 1, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = NUMBER_FORMAT;
				ws.Cells[PROJ_HOURS_ROW + 1, PROJ_ACTUAL_COLUMN + 2, wsrow + 1, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = PERCENT_FORMAT;

				var stoprow = wsrow;
				// Insert Staff Cost section
				wsrow = wsrow + 3;

				for (int row = PROJ_HOURS_ROW + 1; row < stoprow; row++)
				{
					ws.InsertRow(wsrow, 1);
					// insert cost and then formulas
					var position = ws.Cells[row, 1].GetValue<string>();
					var rateCell = ws.Cells[wsrow, ACTUALCOLUMN + 3];
					if (posmap.ContainsKey(position))
					{
						rateCell.Value = posmap[position].Item2;
					}
					else {
						rateCell.Value = "N/A";
					}
					ws.Cells[wsrow, 1].Value = position;
					// formulas ws.Cells[rowNum, ACTUALCOLUMN + 1].Formula = "=" + acell.Address + "-" + ws.Cells[rowNum, BUDGETCOLUMN].Address;
					//ws.Cells[rownum, actcol + 2].Formula = "=IFERROR(" + ws.Cells[rownum, actcol + 1].Address + "/" + ws.Cells[rownum, actcol].Address + ", \"\")";
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN-1].Formula = "=IFERROR(" + ws.Cells[row, PROJ_BUDGET_COLUMN-1].Address + "*" + rateCell.Address + ", \"\")";
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Formula = "=IFERROR(" + ws.Cells[row, PROJ_BUDGET_COLUMN].Address + "*" + rateCell.Address + ", \"\")";
					ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Formula = "=IFERROR(" + ws.Cells[row, PROJ_ACTUAL_COLUMN].Address + "*" + rateCell.Address + ", \"\")";
					// var
					addFormula(ws, null, wsrow, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);
					
					// increment things
					wsrow = wsrow + 1;
				}
				// Totals for staff cost
				ws.InsertRow(wsrow, 1);
				if (entries)
				{
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN - 1].Formula = "=SUM(" + ws.Cells[stoprow + 3, PROJ_BUDGET_COLUMN - 1].Address + ":" + ws.Cells[wsrow - 1, PROJ_BUDGET_COLUMN - 1].Address + ")";
					ws.Cells[wsrow, PROJ_BUDGET_COLUMN].Formula = "=SUM(" + ws.Cells[stoprow + 3, PROJ_BUDGET_COLUMN].Address + ":" + ws.Cells[wsrow - 1, PROJ_BUDGET_COLUMN].Address + ")";
					ws.Cells[wsrow, PROJ_ACTUAL_COLUMN].Formula = "=SUM(" + ws.Cells[stoprow + 3, PROJ_ACTUAL_COLUMN].Address + ":" + ws.Cells[wsrow - 1, PROJ_ACTUAL_COLUMN].Address + ")";
				}
				ws.Cells[wsrow, 1].Value = TOTAL_STAFF_COST;
				addFormula(ws, null, wsrow, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);

				// format cells
				ws.Cells[stoprow + 3, PROJ_BUDGET_COLUMN - 1, wsrow, PROJ_ACTUAL_COLUMN + 1].Style.Numberformat.Format = CURRENCY_FORMAT;
				ws.Cells[stoprow + 3, PROJ_ACTUAL_COLUMN + 2, wsrow, PROJ_ACTUAL_COLUMN + 2].Style.Numberformat.Format = PERCENT_FORMAT;
				// Rate column
				ws.Cells[stoprow + 3, PROJ_ACTUAL_COLUMN + 3, wsrow, PROJ_ACTUAL_COLUMN + 3].Style.Numberformat.Format = CURRENCY_FORMAT;
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
			initOthers();

			object[] args = (object[])e.Argument;
			//string orgName = (string)args[0];
			//string budgetfname = (string)args[1];
			string savefname = Options.Default.OutputFile;
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
			if (File.Exists(Options.Default.OutputFile))
			{
				var result = System.Windows.MessageBox.Show(string.Format("Warning: Report exists. Overwrite?:\n{0}", Options.Default.OutputFile),
				"Warning: Delete existing report", System.Windows.MessageBoxButton.OKCancel, System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.Cancel);
				if (result != System.Windows.MessageBoxResult.OK)
				{
					LogBox("User Cancelled.");
					return;
				}
				File.Delete(Options.Default.OutputFile);
			}
			
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
			processProjectsCostBudget(wb, costBudgetwb, months, start, end);

			// Do project time budget
			processProjectsTime(wb, timeBudgetwb, months);

			addNetCostMargin(wb);
			// autosize before save
			foreach (var ws in wb.Worksheets)
			{
				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				for (var ncol = 1; ncol <= ws.Dimension.Columns; ncol++)
				{
					ws.Column(ncol).Width = ws.Column(ncol).Width + 5;
				}
			}
			// Hide nill budget+actual rows (after autosize because autosize might unhide?)
			hideNilAccounts(wb);

			FormatExpense(wb);

			UnderlineTotals(wb);
			// move overheads to first index
			wb.Worksheets.MoveAfter("Overheads", "Overall PL");

			System.Drawing.Image logo = null;
			if (!Options.Default.LogoFname.Equals("")) { logo = System.Drawing.Image.FromFile(Options.Default.LogoFname); }

			// Add logo to all pages if logo (AFTER copying sheets because weird things happen when in reverse)
			foreach (var ws in wb.Worksheets) {
				if (logo != null)
				{
					AddLogo(ws, logo);
				}
				ws.PrinterSettings.Orientation = eOrientation.Landscape;
				ws.PrinterSettings.FitToPage = true;
			}
			if (logo != null)
			{
				logo.Dispose();
			}
			
			pkg.Save();

			LogBox("Done Monthly dump.\n");

		}

		internal static void SplitSheets(object sender, DoWorkEventArgs e)
		{
			worker = sender as BackgroundWorker;

			if (!initXero()) { return; }
			initOthers();

			object[] args = (object[])e.Argument;
			string fName = (string)args[0];

			LogBox("Splitting report...");

			var pkg = new ExcelPackage(new FileInfo(fName));
			var wb = pkg.Workbook;

			var dir = new FileInfo(fName).Directory;
			dir = dir.CreateSubdirectory("Reports");
			var dirname = dir.FullName;

			// Delete old reports
			var files = false;
			foreach (var fn in Directory.EnumerateFiles(dirname, "*.xlsx"))
			{
				files = true; break;
			}
			if (files)
			{
				var result = System.Windows.MessageBox.Show(string.Format("Warning: Will delete all Excel files in Report folder:\n{0}", dirname),
				"Warning: Delete old files", System.Windows.MessageBoxButton.OKCancel, System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.Cancel);
				if (result != System.Windows.MessageBoxResult.OK)
				{
					LogBox("User Cancelled.");
					return;
				}
				foreach (var fn in Directory.EnumerateFiles(dirname, "*.xlsx"))
				{
					LogBox("Deleting... " + fn);
					File.Delete(fn);
				}
			}

			// Save individual sheets
			foreach (var ws in wb.Worksheets)
			{
				string sheetfname = Path.Combine(dirname, string.Format("{0}.xlsx", ws.Name));
				var sheetpkg = new ExcelPackage(new FileInfo(sheetfname));
				var nws = sheetpkg.Workbook.Worksheets.Add(ws.Name, ws);

				nws.PrinterSettings.Orientation = eOrientation.Landscape;
				nws.PrinterSettings.FitToPage = true;
				sheetpkg.Save();
			}
			LogBox("Split done.");
		}

		private static void addNetCostMargin(ExcelWorkbook wb)
		{
			foreach (var ws in wb.Worksheets)
			{
				int drow = ws.Dimension.Rows;
				int prow = -1; int irow = -1;
				try { prow = searchText(ws, ws.Dimension.Rows, 1, NET_PROFIT, reverse: true, skipempty: true); }
				catch (ArgumentException) { }
				try { irow = searchText(ws, ws.Dimension.Rows, 1, "Total Income", reverse: true, skipempty: true); }
				catch (ArgumentException) { }
				if (prow > 0 && irow > 0)
				{
					//ws.InsertRow(prow+1, 3);
					ws.Cells[drow + 2, 1].Value = "% Margin";
					ws.Cells[drow + 2, 2].Formula = "=B" + prow + "/B" + irow;
					ws.Cells[drow + 2, 3].Formula = "=C" + prow + "/C" + irow;
					ws.Cells[drow + 2, 4].Formula = "=D" + prow + "/D" + irow;
					ws.Cells[drow + 2, 2, drow + 2, 4].Style.Numberformat.Format = PERCENT_FORMAT;
				}

				if (ws.Name.Equals("Overall PL")) { continue; }

				int row = -1;
				try { row = searchText(ws, ws.Dimension.Rows, 1, NET_PROFIT, reverse: true); }
				catch (ArgumentException) { }
				if (row > 0)
				{
					//ws.InsertRow(row+1, 1);
					ws.Cells[row + 1, 1].Value = NET_COST_TITLE;
					int srow = -1;
					try { srow = searchText(ws, 1, 1, TOTAL_STAFF_COST, skipempty: true); }
					catch (ArgumentException) { }
					if (srow > 0) {
						ws.Cells[row + 1, 2].Formula = "=B" + row + "-B"+srow;
						ws.Cells[row + 1, 3].Formula = "=C" + row + "-C" + srow;
						ws.Cells[row + 1, 4].Formula = "=D" + row + "-D" + srow;
					} else
					{
						ws.Cells[row + 1, 2].Value = 0;
						ws.Cells[row + 1, 3].Value = 0;
						ws.Cells[row + 1, 4].Value = 0;
					}
					// add variance stuff (as income)
					addFormula(ws, null, row+1, ACTUALCOLUMN, BUDGETCOLUMN, inc: true);

					// formatting
					formatAccCells(ws, row+1);
				}
			}
		}

		private static void hideNilAccounts(ExcelWorkbook wb)
		{
			foreach (var ws in wb.Worksheets)
			{
				//if (ws.Name.Equals("Overall PL")) { continue; }
				var row = -1;
				try { row = searchText(ws, ws.Dimension.Rows, 1, "Accounts", true, skipempty: true); }
				catch (ArgumentException) { LogBox("CAN'T FIND Accounts?"); }
				if (row > 0)
				{
					row++;
					string data = ws.Cells[row, 1].GetValue<string>();
					while (!IsEmptyCell(data))
					{
						var budget = ws.Cells[row, PROJ_BUDGET_COLUMN - 1].GetValue<decimal>();
						var budgetytd = ws.Cells[row, PROJ_BUDGET_COLUMN].GetValue<decimal>();
						var actual = ws.Cells[row, PROJ_ACTUAL_COLUMN].GetValue<decimal>();
						if (!data.Equals("Operating Expenses") && !data.Contains("Total") && !data.Equals(NET_COST_TITLE))
						{
							if (budget == 0 && budgetytd == 0 && actual == 0)
							{
								if (Options.Default.DelEmptyAccounts)
									ws.DeleteRow(row);
								else
									ws.Row(row).Hidden = true;
							}
						}
						row++;
						data = ws.Cells[row, 1].GetValue<string>();
					}
				}
			}
		}

		private static int getLastUsedColumn(ExcelWorksheet ws, int row)
		{
			int col = 1;
			for (int x = 1; x < 10; x++)
			{
				if (!IsEmptyCell(ws.Cells[row, x].GetValue<string>()) || !IsEmptyCell(ws.Cells[row, x].Formula))
					col = x;
			}
			return col;
		}

		private static void SearchAndUnderline(ExcelWorksheet ws, string text)
		{
			int row = -1;
			try { row = searchText(ws, 4, 1, text, skipempty: true, startswith: true); }
			catch (ArgumentException) { row = -1; }
			if (row > 0)
			{
				// find last used column
				//var col = getLastUsedColumn(ws, row);
				// Just kidding, we want all standard columns "totalled" up to column 6
				var col = 6;
				ws.Cells[row, 1, row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ws.Cells[row, 1, row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
				ws.Cells[row, 1, row, col].Style.Font.Bold = true;
				ws.InsertRow(row, 1);
			}
		}

		private static void UnderlineTotals(ExcelWorkbook wb)
		{
			foreach (var ws in wb.Worksheets)
			{
				//if (ws.Name.Equals("Overall PL")) { continue; }

				int row;
				try { row = searchText(ws, 1, 1, "Total", skipempty: true, startswith: true); }
				catch (ArgumentException) { row = -1; }
				while (row > 0)
				{
					// find last used column
					//var col = getLastUsedColumn(ws, row);
					// Just kidding, we want all standard columns "totalled" up to column 6
					var col = 6;
					ws.Cells[row, 1, row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

					try { row = searchText(ws, row + 1, 1, "Total", skipempty: true, startswith: true); }
					catch (ArgumentException) { row = -1; }
				}

				// do net profit
				SearchAndUnderline(ws, NET_PROFIT);
				SearchAndUnderline(ws, NET_COST_TITLE);
			}
		}
		private static void FormatExpense(ExcelWorkbook wb)
		{
			foreach (var ws in wb.Worksheets)
			{
				//if (ws.Name.Equals("Overall PL")) { continue; }

				int row = -1;
				try { row = searchText(ws, 1, 1, "Operating Expenses", skipempty: true); }
				catch (ArgumentException) { LogBox("CAN'T FIND Accounts? (For FormatExpense)"); }
				if (row > 0)
				{
					ws.Cells[row, 1].Style.Font.Bold = true;
					ws.InsertRow(row, 1);
				}
			}
		}

		private static void AddLogo(ExcelWorksheet ws, System.Drawing.Image logo)
		{
			ws.InsertRow(1, 1);
			
			var pic = ws.Drawings.AddPicture("logo", logo);
			pic.SetPosition(0, 0);
			pic.SetSize(50);
			ws.Row(1).Height = logo.Height / 2.5D;
		}

		private static void setupProjectSheet(ExcelWorksheet ws, string title, DateTime from, DateTime to)
		{
			ws.Cells["A1"].Value = "Profit and loss"; ws.Cells["C1"].Value = "Begining"; ws.Cells["D1"].Value = "Ending";
			ws.Cells["A2"].Value = "Project";
			ws.Cells["B2"].Value = title; ws.Cells["C2"].Value = from; ws.Cells["D2"].Value = to;

			// staff Hours
			ws.Cells[PROJ_HOURS_ROW, 1].Value = "Staff Days"; ws.Cells[PROJ_HOURS_ROW, PROJ_BUDGET_COLUMN-1].Value = string.Format(FULL_YEAR, from.Year % 100, (from.Year % 100) + 1);
			ws.Cells[PROJ_HOURS_ROW, PROJ_BUDGET_COLUMN].Value = "Budget YTD"; ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN].Value = "Actual"; ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN + 1].Value = "Var";
			ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN + 2].Value = "Var %"; ws.Cells[PROJ_HOURS_ROW, PROJ_ACTUAL_COLUMN + 3].Value = "Notes";
			// staff Cost
			ws.Cells[PROJ_COST_ROW, 1].Value = "Staff Cost"; ws.Cells[PROJ_COST_ROW, PROJ_BUDGET_COLUMN - 1].Value = string.Format(FULL_YEAR, from.Year % 100, (from.Year % 100) + 1);
			ws.Cells[PROJ_COST_ROW, PROJ_BUDGET_COLUMN].Value = "Budget YTD"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN].Value = "Actual";
			ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 1].Value = "Var"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 2].Value = "Var %";
			ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 3].Value = "Rate"; ws.Cells[PROJ_COST_ROW, PROJ_ACTUAL_COLUMN + 4].Value = "Notes";

			// Formatting
			ws.Cells["A1:D1"].Style.Font.Bold = true; ws.Cells["A2"].Style.Font.Bold = true; ws.Cells["4:4"].Style.Font.Bold = true; ws.Cells["6:6"].Style.Font.Bold = true;
			ws.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
			ws.Cells["B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
			ws.Cells["C2"].Style.Numberformat.Format = DATE_FORMAT; ws.Cells["D2"].Style.Numberformat.Format = DATE_FORMAT;

			ws.Cells[PROJ_OTHERS_ROW, PROJ_BUDGET_COLUMN - 1].Value = string.Format(FULL_YEAR, from.Year % 100, (from.Year % 100) + 1);
			ws.Cells[PROJ_OTHERS_ROW, PROJ_BUDGET_COLUMN].Value = "Budget YTD"; ws.Cells[PROJ_OTHERS_ROW, PROJ_ACTUAL_COLUMN].Value = "Actual";
			ws.Cells[PROJ_OTHERS_ROW, PROJ_ACTUAL_COLUMN + 1].Value = "Var $"; ws.Cells[PROJ_OTHERS_ROW, PROJ_ACTUAL_COLUMN + 2].Value = "Var %";
			ws.Cells[PROJ_OTHERS_ROW, 1].Value = "Accounts"; ws.Cells[PROJ_OTHERS_ROW, PROJ_ACTUAL_COLUMN + 3].Value = "Notes";
			ws.Cells[string.Format("{0}:{0}", PROJ_OTHERS_ROW)].Style.Font.Bold = true;
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
				// project names
				if (headers == null) { headers = row.Cells.Select(x => x.Value).ToList(); }

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
			//collate projects
			foreach (var k in preport.Keys.ToList())
			{
				//LogBox("Processing Project " + k);
				if (projCollation.ContainsKey(k)) {
					var target = projCollation[k];
					if (target != k)
					{
						if (preport.ContainsKey(target))
						{
							LogBox(string.Format("Collating {0} in to {1}", k, target));
							preport[target] = preport[target].Zip(preport[k], (i1, i2) => (decimal.Parse(i1) + decimal.Parse(i2)).ToString()).ToList();
						} else
						{
							LogBox(string.Format("Moving {0} in to {1}", k, target));
							preport[target] = preport[k];
						}
						preport.Remove(k);
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
					if (k.Equals("Total"))
						continue;
					else
					{
						ws = wb.Worksheets.Add(k);
						setupProjectSheet(ws, k, from, to);
					}
				}
				var rownum = PROJ_OTHERS_ROW;
				rownum++;

				var incomelist = new List<string>();
				var incomeitems = new Dictionary<string, decimal>();
				var expenselist = new List<string>();
				var expenseitems = new Dictionary<string, decimal>();
				Tuple<string, decimal> netprofit = null;

				int i = 0;
				foreach (var line in projectReports[k])
				{
					// get name and value
					var acct = categories[i];
					if (HideAccounts.Contains(acct)) { i++; continue; }
					var value = decimal.Parse(line);

					// check if merging accounts
					if (mergeAccounts.ContainsKey(acct))
					{
						acct = mergeAccounts[acct];
					}
					decimal item = 0;
					// income or expense?
					if (acct.Equals(NET_PROFIT, StringComparison.OrdinalIgnoreCase) || acct.Equals(GNET_PROFIT, StringComparison.OrdinalIgnoreCase))
					{
						netprofit = new Tuple<string, decimal>(acct, value);
					}
					else if (IncomeAccounts.Contains(acct))
					{
						incomeitems.TryGetValue(acct, out item);
						incomeitems[acct] = item + value;
						if (!incomelist.Contains(acct))
						{
							if (acct.StartsWith("Total", StringComparison.OrdinalIgnoreCase))
								incomelist.Add(acct);
							else
								incomelist.Insert(0, acct);
						}
					}
					else
					{
						expenseitems.TryGetValue(acct, out item);
						expenseitems[acct] = item + value;
						if (!expenselist.Contains(acct))
						{
							if (acct.StartsWith("Total", StringComparison.OrdinalIgnoreCase))
								expenselist.Add(acct);
							else
								expenselist.Insert(0, acct);
						}
					}

					i++;
				}

				//income 
				foreach (var acct in incomelist)
				{
					ws.Cells[rownum, 1].Value = acct;
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Value = incomeitems[acct];
					ws.Cells[rownum, PROJ_BUDGET_COLUMN - 1, rownum, PROJ_BUDGET_COLUMN].Value = 0; // fill in missing actual

					// insert formulas
					addFormula(ws, acct, rownum, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);
					rownum++;
				}
				//LogBox(string.Format("Inserting for {0} at: {1}", ws.Name, rownum));
				ws.Cells[rownum, 1].Value = "Operating Expenses";
				rownum++;
				//expense 
				foreach (var acct in expenselist)
				{

					ws.Cells[rownum, 1].Value = acct;
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Value = expenseitems[acct];
					ws.Cells[rownum, PROJ_BUDGET_COLUMN - 1, rownum, PROJ_BUDGET_COLUMN].Value = 0; // fill in missing actual
					// insert formulas
					addFormula(ws, acct, rownum, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN);
					rownum++;
				}
				if (netprofit != null)
				{
					ws.Cells[rownum, 1].Value = netprofit.Item1;
					ws.Cells[rownum, PROJ_ACTUAL_COLUMN].Value = netprofit.Item2;
					ws.Cells[rownum, PROJ_BUDGET_COLUMN - 1, rownum, PROJ_BUDGET_COLUMN].Value = 0; // fill in missing actual
					addFormula(ws, null, rownum, PROJ_ACTUAL_COLUMN, PROJ_BUDGET_COLUMN, inc: true);
					rownum++;
				}
			}
		}

		private static void addFormula(ExcelWorksheet ws, string accname, int rownum, int actcol, int budcol, bool inc = false)
		{
			if (inc || accname != null && IncomeAccounts.Contains(accname))
			{   // actual - budget
				ws.Cells[rownum, actcol + 1].Formula = "=IFERROR(" + ws.Cells[rownum, actcol].Address + "-" + ws.Cells[rownum, budcol].Address + ", \"\")";
				// variance / actual
				ws.Cells[rownum, actcol + 2].Formula = "=IFERROR(" + ws.Cells[rownum, actcol + 1].Address + "/" + ws.Cells[rownum, actcol].Address + ", \"\")";
			}
			else
			{   // budget - actual
				ws.Cells[rownum, actcol + 1].Formula = "=IFERROR(" + ws.Cells[rownum, budcol].Address + "-" + ws.Cells[rownum, actcol].Address + ", \"\")";
				// variance / budget
				ws.Cells[rownum, actcol + 2].Formula = "=IFERROR(" + ws.Cells[rownum, actcol + 1].Address + "/" + ws.Cells[rownum, budcol].Address + ", \"\")";
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
