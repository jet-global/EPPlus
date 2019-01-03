using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelWorksheetTest : TestBase
	{
		#region Tests From Previous EPPlus Contributors
		[TestMethod]
		public void RunWorksheetTests()
		{
			// TODO (Task #8178): Fix grouping/test.
			InsertDeleteTestRows();
			InsertDeleteTestColumns();
			LoadData();
			AutoFilter();
			StyleFill();
			Performance();
			RichTextCells();
			TestComments();
			Hyperlink();
			PictureURL();
			CopyOverwrite();
			HideTest();
			VeryHideTest();
			PrinterSettings();
			Address();
			Merge();
			Encoding();
			LoadText();
			LoadDataReader();
			LoadDataTable();
			LoadFromCollectionTest();
			LoadFromEmptyCollectionTest();
			LoadArray();
			WorksheetCopy();
			DefaultColWidth();
			CopyTable();
			AutoFitColumns();
			CopyRange();
			CopyMergedRange();
			ValueError();
			FormulaOverwrite();
			FormulaError();
			StyleNameTest();
			NamedStyles();
			TableTest();
			DefinedName();
			CreatePivotTable();
			AddChartSheet();
			SetHeaderFooterImage();

			SaveWorksheet("Worksheet.xlsx");

			ReadWorkSheet();
			ReadStreamSaveAsStream();
		}

		private void AutoFilter()
		{
			var ws = _pck.Workbook.Worksheets.Add("Autofilter");
			ws.Cells["A1"].Value = "A1";
			ws.Cells["B1"].Value = "B1";
			ws.Cells["C1"].Value = "C1";
			ws.Cells["D1"].Value = "D1";

			ws.Cells["A2"].Value = 1;
			ws.Cells["B2"].Value = 2;
			ws.Cells["C2"].Value = 3;
			ws.Cells["D2"].Value = 4;

			ws.Cells["A1:D2"].AutoFilter = true;
			ws.Cells["A1:D2"].AutoFilter = false;
			ws.Cells["A1:D2"].AutoFilter = true;
			ws.Cells["A1:D5"].AutoFilter = false;
			ws.Cells["A1:D2"].AutoFilter = true;
		}

		private void AddChartSheet()
		{
			var chart = _pck.Workbook.Worksheets.AddChart("ChartSheet", eChartType.ColumnClustered);
			foreach (var _n in _pck.Workbook.Names)
			{

			}
			//Iterate all collection and make sure no exception is thrown.
			foreach (var worksheet in _pck.Workbook.Worksheets)
			{
				if (!(worksheet is ExcelChartsheet))
				{
					foreach (var d in worksheet.Drawings)
					{

					}
					foreach (var d in worksheet.Tables)
					{

					}
					foreach (var d in worksheet.PivotTables)
					{

					}
					foreach (var d in worksheet.Names)
					{

					}
					foreach (var d in worksheet.Comments)
					{

					}
					foreach (var d in worksheet.ConditionalFormatting)
					{

					}
				}
			}
		}

		//[Ignore]
		//[TestMethod]
		public void ReadWorkSheet()
		{
			FileStream instream = new FileStream(_worksheetPath + @"Worksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);
			using (ExcelPackage pck = new ExcelPackage(instream))
			{
				var ws = pck.Workbook.Worksheets["Perf"];
				Assert.AreEqual(ws.Cells["H6"].Formula, "B5+B6");

				ws = pck.Workbook.Worksheets["Comment"];
				var comment = ws.Cells["B2"].Comment;

				Assert.AreNotEqual(comment, null);
				Assert.AreEqual(comment.Author, "Jan Källman");
				ws = pck.Workbook.Worksheets["Hidden"];
				Assert.AreEqual<eWorkSheetHidden>(ws.Hidden, eWorkSheetHidden.Hidden);

				ws = pck.Workbook.Worksheets["VeryHidden"];
				Assert.AreEqual<eWorkSheetHidden>(ws.Hidden, eWorkSheetHidden.VeryHidden);

				ws = pck.Workbook.Worksheets["RichText"];
				Assert.AreEqual("Room 02 & 03", ws.Cells["G1"].RichText.Text);

				ws = pck.Workbook.Worksheets["HeaderImage"];
				//Assert.AreEqual(ws.HeaderFooter.Pictures.Count, 3);

				ws = pck.Workbook.Worksheets["newsheet"];
				Assert.AreEqual(ws.Cells["F2"].Style.Font.UnderLine, true);
				Assert.AreEqual(ws.Cells["F2"].Style.Font.UnderLineType, ExcelUnderLineType.Double);
				Assert.AreEqual(ws.Cells["F3"].Style.Font.UnderLineType, ExcelUnderLineType.SingleAccounting);
				Assert.AreEqual(ws.Cells["F5"].Style.Font.UnderLineType, ExcelUnderLineType.None);
				Assert.AreEqual(ws.Cells["F5"].Style.Font.UnderLine, false);

				Assert.AreEqual(ws.Cells["T20"].GetValue<string>(), 0.396180555555556d.ToString(CultureInfo.CurrentCulture));
				Assert.AreEqual(ws.Cells["T20"].GetValue<int>(), 0);
				Assert.AreEqual(ws.Cells["T20"].GetValue<int?>(), 0);
				Assert.AreEqual(ws.Cells["T20"].GetValue<double>(), 0.396180555555556d);
				Assert.AreEqual(ws.Cells["T20"].GetValue<double?>(), 0.396180555555556d);
				Assert.AreEqual(ws.Cells["T20"].GetValue<decimal>(), 0.396180555555556m);
				Assert.AreEqual(ws.Cells["T20"].GetValue<decimal?>(), 0.396180555555556m);
				Assert.AreEqual(ws.Cells["T20"].GetValue<bool>(), true);
				Assert.AreEqual(ws.Cells["T20"].GetValue<bool?>(), true);
				Assert.AreEqual(ws.Cells["T20"].GetValue<DateTime>(), new DateTime(1899, 12, 30, 9, 30, 30));
				Assert.AreEqual(ws.Cells["T20"].GetValue<DateTime?>(), new DateTime(1899, 12, 30, 9, 30, 30));
				Assert.AreEqual(ws.Cells["T20"].GetValue<TimeSpan>(), new TimeSpan(693593, 9, 30, 30));
				Assert.AreEqual(ws.Cells["T20"].GetValue<TimeSpan?>(), new TimeSpan(693593, 9, 30, 30));
				Assert.AreEqual(ws.Cells["T20"].Text, "09:30:30");

				Assert.AreEqual(ws.Cells["T24"].GetValue<string>(), 1.39618055555556d.ToString(CultureInfo.CurrentCulture));
				Assert.AreEqual(ws.Cells["T24"].GetValue<int>(), 1);
				Assert.AreEqual(ws.Cells["T24"].GetValue<int?>(), 1);
				Assert.AreEqual(ws.Cells["T24"].GetValue<double>(), 1.39618055555556d);
				Assert.AreEqual(ws.Cells["T24"].GetValue<double?>(), 1.39618055555556d);
				Assert.AreEqual(ws.Cells["T24"].GetValue<decimal>(), 1.39618055555556m);
				Assert.AreEqual(ws.Cells["T24"].GetValue<decimal?>(), 1.39618055555556m);
				Assert.AreEqual(ws.Cells["T24"].GetValue<bool>(), true);
				Assert.AreEqual(ws.Cells["T24"].GetValue<bool?>(), true);
				Assert.AreEqual(ws.Cells["T24"].GetValue<DateTime>(), new DateTime(1899, 12, 31, 9, 30, 30));
				Assert.AreEqual(ws.Cells["T24"].GetValue<DateTime?>(), new DateTime(1899, 12, 31, 9, 30, 30));
				Assert.AreEqual(ws.Cells["T24"].GetValue<TimeSpan>(), new TimeSpan(693593, 33, 30, 30));
				Assert.AreEqual(ws.Cells["T24"].GetValue<TimeSpan?>(), new TimeSpan(693593, 33, 30, 30));
				Assert.AreEqual(ws.Cells["T24"].Text, "09:30:30");

				Assert.AreEqual(ws.Cells["U20"].GetValue<string>(), "40179");
				Assert.AreEqual(ws.Cells["U20"].GetValue<int>(), 40179);
				Assert.AreEqual(ws.Cells["U20"].GetValue<int?>(), 40179);
				Assert.AreEqual(ws.Cells["U20"].GetValue<double>(), 40179d);
				Assert.AreEqual(ws.Cells["U20"].GetValue<double?>(), 40179d);
				Assert.AreEqual(ws.Cells["U20"].GetValue<decimal>(), 40179m);
				Assert.AreEqual(ws.Cells["U20"].GetValue<decimal?>(), 40179m);
				Assert.AreEqual(ws.Cells["U20"].GetValue<bool>(), true);
				Assert.AreEqual(ws.Cells["U20"].GetValue<bool?>(), true);
				Assert.AreEqual(ws.Cells["U20"].GetValue<DateTime>(), new DateTime(2010, 1, 1));
				Assert.AreEqual(ws.Cells["U20"].GetValue<DateTime?>(), new DateTime(2010, 1, 1));
				Assert.AreEqual(ws.Cells["U20"].Text, "2010-01-01");

				Assert.AreEqual(ws.Cells["V20"].GetValue<string>(), "102");
				Assert.AreEqual(ws.Cells["V20"].GetValue<int>(), 102);
				Assert.AreEqual(ws.Cells["V20"].GetValue<int?>(), 102);
				Assert.AreEqual(ws.Cells["V20"].GetValue<double>(), 102d);
				Assert.AreEqual(ws.Cells["V20"].GetValue<double?>(), 102d);
				Assert.AreEqual(ws.Cells["V20"].GetValue<decimal>(), 102m);
				Assert.AreEqual(ws.Cells["V20"].GetValue<decimal?>(), 102m);
				Assert.AreEqual(ws.Cells["V20"].GetValue<bool>(), true);
				Assert.AreEqual(ws.Cells["V20"].GetValue<bool?>(), true);
				Assert.AreEqual(ws.Cells["V20"].GetValue<DateTime>(), new DateTime(1900, 4, 11));
				Assert.AreEqual(ws.Cells["V20"].GetValue<DateTime?>(), new DateTime(1900, 4, 11));
				Assert.AreEqual(ws.Cells["V20"].Text,
					$"$102{CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator}00");

				Assert.AreEqual(ws.Cells["W20"].GetValue<string>(), null);
				Assert.AreEqual(ws.Cells["W20"].GetValue<int>(), 0);
				Assert.AreEqual(ws.Cells["W20"].GetValue<int?>(), null);
				Assert.AreEqual(ws.Cells["W20"].GetValue<double>(), 0d);
				Assert.AreEqual(ws.Cells["W20"].GetValue<double?>(), null);
				Assert.AreEqual(ws.Cells["W20"].GetValue<decimal>(), 0m);
				Assert.AreEqual(ws.Cells["W20"].GetValue<decimal?>(), null);
				Assert.AreEqual(ws.Cells["W20"].GetValue<bool>(), false);
				Assert.AreEqual(ws.Cells["W20"].GetValue<bool?>(), null);
				Assert.AreEqual(ws.Cells["W20"].GetValue<DateTime>(), DateTime.MinValue);
				Assert.AreEqual(ws.Cells["W20"].GetValue<DateTime?>(), null);
				Assert.AreEqual(ws.Cells["W20"].GetValue<TimeSpan>(), TimeSpan.Zero);
				Assert.AreEqual(ws.Cells["W20"].GetValue<TimeSpan?>(), null);
				Assert.AreEqual(ws.Cells["W20"].Text, string.Empty);

				Assert.AreEqual(ws.Cells["Y20"].GetValue<string>(), "True");
				Assert.AreEqual(ws.Cells["Y20"].GetValue<int>(), 0);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<int?>(), null);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<double>(), 0d);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<double?>(), null);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<decimal>(), 0m);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<decimal?>(), null);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<bool>(), true);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<bool?>(), true);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<DateTime>(), DateTime.MinValue);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<DateTime?>(), null);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<TimeSpan>(), TimeSpan.Zero);
				Assert.AreEqual(ws.Cells["Y20"].GetValue<TimeSpan?>(), null);
				Assert.AreEqual(ws.Cells["Y20"].Text, "1");

				Assert.AreEqual(ws.Cells["Z20"].GetValue<string>(), "Text2");
				Assert.AreEqual(ws.Cells["Z20"].GetValue<int>(), 0);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<int?>(), null);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<double>(), 0d);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<double?>(), null);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<decimal>(), 0m);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<decimal?>(), null);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<bool>(), false);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<bool?>(), null);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<DateTime>(), DateTime.MinValue);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<DateTime?>(), null);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<TimeSpan>(), TimeSpan.Zero);
				Assert.AreEqual(ws.Cells["Z20"].GetValue<TimeSpan?>(), null);
				Assert.AreEqual(ws.Cells["Z20"].Text, "Text2");
			}
			instream.Close();
		}

		[Ignore]
		[TestMethod]
		public void ReadStreamWithTemplateWorkSheet()
		{
			FileStream instream = new FileStream(_worksheetPath + @"\Worksheet.xlsx", FileMode.Open, FileAccess.Read);
			MemoryStream stream = new MemoryStream();
			using (ExcelPackage pck = new ExcelPackage(stream, instream))
			{
				var ws = pck.Workbook.Worksheets["Perf"];
				Assert.AreEqual(ws.Cells["H6"].Formula, "B5+B6");

				ws = pck.Workbook.Worksheets["newsheet"];
				Assert.AreEqual(ws.GetValue<DateTime>(20, 21), new DateTime(2010, 1, 1));

				ws = pck.Workbook.Worksheets["Loaded DataTable"];
				Assert.AreEqual(ws.GetValue<string>(2, 1), "Row1");
				Assert.AreEqual(ws.GetValue<int>(2, 2), 1);
				Assert.AreEqual(ws.GetValue<bool>(2, 3), true);
				Assert.AreEqual(ws.GetValue<double>(2, 4), 1.5);

				ws = pck.Workbook.Worksheets["RichText"];

				var r1 = ws.Cells["A1"].RichText[0];
				Assert.AreEqual(r1.Text, "Test");
				Assert.AreEqual(r1.Bold, true);

				ws = pck.Workbook.Worksheets["Pic URL"];
				Assert.AreEqual(((ExcelPicture)ws.Drawings["Pic URI"]).Hyperlink, "http://epplus.codeplex.com");

				Assert.AreEqual(pck.Workbook.Worksheets["Address"].GetValue<string>(40, 1), "\b\t");

				pck.SaveAs(new FileInfo(@"Test\Worksheet2.xlsx"));
			}
			instream.Close();
		}

		//[Ignore]
		//[TestMethod]
		public void ReadStreamSaveAsStream()
		{
			if (!File.Exists(_worksheetPath + @"Worksheet.xlsx"))
			{
				Assert.Inconclusive("Worksheet.xlsx does not exists");
			}
			FileStream instream = new FileStream(_worksheetPath + @"Worksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);
			MemoryStream stream = new MemoryStream();
			using (ExcelPackage pck = new ExcelPackage(instream))
			{
				var ws = pck.Workbook.Worksheets["Names"];

				var address = new ExcelAddress(ws.Names["FullCol"].NameFormula);
				Assert.AreEqual(1, address.Start.Row);
				Assert.AreEqual(ExcelPackage.MaxRows, address.End.Row);
				pck.SaveAs(stream);
			}
			instream.Close();
		}

		//
		// You can use the following additional attributes as you write your tests:
		//
		// Use ClassInitialize to run code before running the first test in the class
		// Use ClassCleanup to run code after all tests in a class have run
		//[Ignore]
		//[TestMethod]
		public void LoadData()
		{
			ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("newsheet");
			ws.Cells["T19"].Value = new TimeSpan(3, 30, 30);
			ws.Cells["T20"].Value = new TimeSpan(9, 30, 30);
			ws.Cells["T21"].Value = new TimeSpan(15, 30, 30);
			ws.Cells["T22"].Value = new TimeSpan(21, 30, 30);
			ws.Cells["T23"].Value = new TimeSpan(27, 30, 30);
			ws.Cells["T24"].Value = new TimeSpan(33, 30, 30);
			ws.Cells["T19:T24"].Style.Numberformat.Format = "hh:mm:ss";

			ws.Cells["U19"].Value = new DateTime(2009, 12, 31);
			ws.Cells["U20"].Value = new DateTime(2010, 1, 1);
			ws.Cells["U21"].Value = new DateTime(2010, 1, 2);
			ws.Cells["U22"].Value = new DateTime(2010, 1, 3);
			ws.Cells["U23"].Value = new DateTime(2010, 1, 4);
			ws.Cells["U24"].Value = new DateTime(2010, 1, 5);
			ws.Cells["U19:U24"].Style.Numberformat.Format = "yyyy-mm-dd";

			ws.Cells["V19"].Value = 100;
			ws.Cells["V20"].Value = 102;
			ws.Cells["V21"].Value = 101;
			ws.Cells["V22"].Value = 103;
			ws.Cells["V23"].Value = 105;
			ws.Cells["V24"].Value = 104;
			ws.Cells["v19:v24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
			ws.Cells["v19:v24"].Style.Numberformat.Format = @"$#,##0.00_);($#,##0.00)";

			ws.Cells["X19"].Value = 210;
			ws.Cells["X20"].Value = 212;
			ws.Cells["X21"].Value = 221;
			ws.Cells["X22"].Value = 123;
			ws.Cells["X23"].Value = 135;
			ws.Cells["X24"].Value = 134;

			ws.Cells["Y19"].Value = true;
			ws.Cells["Y20"].Value = true;
			ws.Cells["Y21"].Value = true;
			ws.Cells["Y22"].Value = false;
			ws.Cells["Y23"].Value = false;
			ws.Cells["Y24"].Value = false;

			ws.Cells["Z19"].Value = "Text1";
			ws.Cells["Z20"].Value = "Text2";
			ws.Cells["Z21"].Value = "Text3";
			ws.Cells["Z22"].Value = "Text4";
			ws.Cells["Z23"].Value = "Text5";
			ws.Cells["Z24"].Value = "Text6";

			// add autofilter
			ws.Cells["U19:X24"].AutoFilter = true;
			ExcelPicture pic = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
			pic.SetPosition(150, 140);

			ws.Cells["A30"].Value = "Text orientation 45";
			ws.Cells["A30"].Style.TextRotation = 45;
			ws.Cells["B30"].Value = "Text orientation 90";
			ws.Cells["B30"].Style.TextRotation = 90;
			ws.Cells["C30"].Value = "Text orientation 180";
			ws.Cells["C30"].Style.TextRotation = 180;
			ws.Cells["D30"].Value = "Text orientation 38";
			ws.Cells["D30"].Style.TextRotation = 38;
			ws.Cells["D30"].Style.Font.Bold = true;
			ws.Cells["D30"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

			//Test vertical align
			ws.Cells["E19"].Value = "Subscript";
			ws.Cells["E19"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
			ws.Cells["E20"].Value = "Subscript";
			ws.Cells["E20"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
			ws.Cells["E21"].Value = "Superscript";
			ws.Cells["E21"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
			ws.Cells["E21"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.None;

			ws.Cells["E22"].Value = "Indent 2";
			ws.Cells["E22"].Style.Indent = 2;
			ws.Cells["E23"].Value = "Shrink to fit";
			ws.Cells["E23"].Style.ShrinkToFit = true;

			ws.Cells["e24"].Value = "ReadingOrder LeftToRight";
			ws.Cells["e24"].Style.ReadingOrder = ExcelReadingOrder.LeftToRight;
			ws.Cells["e25"].Value = "ReadingOrder RightToLeft";
			ws.Cells["e25"].Style.ReadingOrder = ExcelReadingOrder.RightToLeft;
			ws.Cells["e26"].Value = "ReadingOrder Context";
			ws.Cells["e26"].Style.ReadingOrder = ExcelReadingOrder.ContextDependent;
			ws.Cells["e27"].Value = "Default Readingorder";

			//Underline
			ws.Cells["F1:F7"].Value = "Underlined";
			ws.Cells["F1"].Style.Font.UnderLineType = ExcelUnderLineType.Single;
			ws.Cells["F2"].Style.Font.UnderLineType = ExcelUnderLineType.Double;
			ws.Cells["F3"].Style.Font.UnderLineType = ExcelUnderLineType.SingleAccounting;
			ws.Cells["F4"].Style.Font.UnderLineType = ExcelUnderLineType.DoubleAccounting;
			ws.Cells["F5"].Style.Font.UnderLineType = ExcelUnderLineType.None;
			ws.Cells["F6:F7"].Style.Font.UnderLine = true;
			ws.Cells["F7"].Style.Font.UnderLine = false;

			ws.Cells["E24"].Value = 0;
			Assert.AreEqual(ws.Cells["E24"].Text, "0");
			ws.Cells["F7"].Style.Font.UnderLine = false;
			ws.Names.Add("SheetName", ws.Cells["A1:A2"]);
			ws.View.FreezePanes(3, 5);

			foreach (ExcelRangeBase cell in ws.Cells["A1"])
			{
				Assert.Fail("A1 is not set");
			}

			foreach (ExcelRangeBase cell in ws.Cells[ws.Dimension.Address])
			{
				System.Diagnostics.Debug.WriteLine(cell.Address);
			}

			// Linq test
			var res = from c in ws.Cells[ws.Dimension.Address] where c.Value != null && c.Value.ToString() == "Offset test 1" select c;

			foreach (ExcelRangeBase cell in res)
			{
				System.Diagnostics.Debug.WriteLine(cell.Address);
			}

			_pck.Workbook.Properties.Author = "Jan Källman";
			_pck.Workbook.Properties.Category = "Category";
			_pck.Workbook.Properties.Comments = "Comments";
			_pck.Workbook.Properties.Company = "Adventure works";
			_pck.Workbook.Properties.Keywords = "Keywords";
			_pck.Workbook.Properties.Title = "Title";
			_pck.Workbook.Properties.Subject = "Subject";
			_pck.Workbook.Properties.Status = "Status";
			_pck.Workbook.Properties.HyperlinkBase = new Uri("http://serversideexcel.com", UriKind.Absolute);
			_pck.Workbook.Properties.Manager = "Manager";

			_pck.Workbook.Properties.SetCustomPropertyValue("DateTest", new DateTime(2008, 12, 31));
			TestContext.WriteLine(_pck.Workbook.Properties.GetCustomPropertyValue("DateTest").ToString());
			_pck.Workbook.Properties.SetCustomPropertyValue("Author", "Jan Källman");
			_pck.Workbook.Properties.SetCustomPropertyValue("Count", 1);
			_pck.Workbook.Properties.SetCustomPropertyValue("IsTested", false);
			_pck.Workbook.Properties.SetCustomPropertyValue("LargeNo", 123456789123);
			_pck.Workbook.Properties.SetCustomPropertyValue("Author", 3);
		}

		const int PERF_ROWS = 5000;

		//[Ignore]
		//[TestMethod]
		public void Performance()
		{
			ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("Perf");
			TestContext.WriteLine("StartTime {0}", DateTime.Now);

			Random r = new Random();
			for (int i = 1; i <= PERF_ROWS; i++)
			{
				ws.Cells[i, 1].Value = string.Format("Row {0}\n.Test new row\"' öäåü", i);
				ws.Cells[i, 2].Value = i;
				ws.Cells[i, 2].Style.WrapText = true;
				ws.Cells[i, 3].Value = DateTime.Now;
				ws.Cells[i, 4].Value = r.NextDouble() * 100000;
			}
			ws.Cells[1, 2, PERF_ROWS, 2].Style.Numberformat.Format = "#,##0";
			ws.Cells[1, 3, PERF_ROWS, 3].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
			ws.Cells[1, 4, PERF_ROWS, 4].Style.Numberformat.Format = "#,##0.00";
			ws.Cells[PERF_ROWS + 1, 2].Formula = "SUM(B1:B" + PERF_ROWS.ToString() + ")";
			ws.Column(1).Width = 12;
			ws.Column(2).Width = 8;
			ws.Column(3).Width = 20;
			ws.Column(4).Width = 14;

			ws.Cells["A1:C1"].Merge = true;
			ws.Cells["A2:A5"].Merge = true;
			ws.DeleteRow(1, 1);
			ws.InsertRow(1, 1);
			ws.InsertRow(3, 1);

			ws.DeleteRow(1000, 3, true);
			ws.DeleteRow(2000, 1, true);

			ws.InsertRow(2001, 4);

			ws.InsertRow(2010, 1, 2010);

			ws.InsertRow(20000, 2);

			ws.DeleteRow(20005, 4, false);

			//Single formula
			ws.Cells["H3"].Formula = "B2+B3";
			ws.DeleteRow(2, 1, true);

			//Shared formula
			ws.Cells["H5:H30"].Formula = "B4+B5";
			ws.Cells["H5:H30"].Style.Numberformat.Format = "_(\"$\"* # ##0.00_);_(\"$\"* (# ##0.00);_(\"$\"* \"-\"??_);_(@_)";
			ws.InsertRow(7, 3);
			ws.InsertRow(2, 1);
			ws.DeleteRow(30, 3, true);

			ws.DeleteRow(15, 2, true);
			ws.Cells["a1:B100"].Style.Locked = false;
			ws.Cells["a1:B12"].Style.Hidden = true;
			TestContext.WriteLine("EndTime {0}", DateTime.Now);
		}

		//[Ignore]
		//[TestMethod]
		public void InsertDeleteTestRows()
		{
			ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("InsertDelete");
			//ws.Cells.Value = 0;
			ws.Cells["A1:C5"].Value = 1;
			Assert.AreEqual(((object[,])ws.Cells["A1:C5"].Value)[1, 1], 1);
			ws.Cells["A1:B3"].Merge = true;
			ws.Cells["D3"].Formula = "A2+C5";
			ws.InsertRow(2, 1);

			ws.Cells["A10:C15"].Value = 1;
			ws.Cells["A11:B13"].Merge = true;
			ws.DeleteRow(12, 1, true);

			ws.Cells["a1:B100"].Style.Locked = false;
			ws.Cells["a1:B12"].Style.Hidden = true;
			ws.Protection.IsProtected = true;
			ws.Protection.SetPassword("Password");


			var range = ws.Cells["B2:D100"];

			ws.PrinterSettings.PrintArea = null;
			ws.PrinterSettings.PrintArea = ws.Cells["B2:D99"];
			ws.PrinterSettings.PrintArea = null;
			ws.Row(15).PageBreak = true;
			ws.Column(3).PageBreak = true;
			ws.View.ShowHeaders = false;
			ws.View.PageBreakView = true;

			ws.Row(200).Height = 50;
			ws.Workbook.CalcMode = ExcelCalcMode.Automatic;

			Assert.AreEqual(range.Start.Column, 2);
			Assert.AreEqual(range.Start.Row, 2);
			Assert.AreEqual(range.Start.Address, "B2");

			Assert.AreEqual(range.End.Column, 4);
			Assert.AreEqual(range.End.Row, 100);
			Assert.AreEqual(range.End.Address, "D100");

			ExcelAddress addr = new ExcelAddress("B1:D3");

			Assert.AreEqual(addr.Start.Column, 2);
			Assert.AreEqual(addr.Start.Row, 1);
			Assert.AreEqual(addr.End.Column, 4);
			Assert.AreEqual(addr.End.Row, 3);
		}

		//[Ignore]
		//[TestMethod]
		public void InsertDeleteTestColumns()
		{
			ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("InsertDeleteColumns");
			//ws.Cells.Value = 0;
			ws.Cells["A1:C1"].Value = 1;
			ws.Cells["A2:C2"].Value = 2;
			ws.Cells["A3:C3"].Value = 3;
			ws.Cells["A4:C4"].Value = 4;
			ws.Cells["A5:C5"].Value = 5;
			Assert.AreEqual(((object[,])ws.Cells["A1:C5"].Value)[1, 1], 2);
			ws.Cells["A1:B3"].Merge = true;
			ws.Cells["D3"].Formula = "A2+C5";
			ws.InsertColumn(1, 1);

			//ws.DeleteColumn(3, 2);
			ws.Cells["K10:M15"].Value = 1;
			ws.Cells["K11:L13"].Merge = true;
			ws.DeleteColumn(12, 1);

			ws.Cells["X1:Y100"].Style.Locked = false;
			ws.Cells["C1:Y12"].Style.Hidden = true;
			ws.Protection.IsProtected = true;
			ws.Protection.SetPassword("Password");


			var range = ws.Cells["X2:Z100"];

			ws.PrinterSettings.PrintArea = null;
			ws.PrinterSettings.PrintArea = ws.Cells["X2:Z99"];
			ws.PrinterSettings.PrintArea = null;
			ws.Row(15).PageBreak = true;
			ws.Column(3).PageBreak = true;
			ws.View.ShowHeaders = false;
			ws.View.PageBreakView = true;

			ws.Row(200).Height = 50;
			ws.Workbook.CalcMode = ExcelCalcMode.Automatic;

			//Assert.AreEqual(range.Start.Column, 2);
			//Assert.AreEqual(range.Start.Row, 2);
			//Assert.AreEqual(range.Start.Address, "B2");

			//Assert.AreEqual(range.End.Column, 4);
			//Assert.AreEqual(range.End.Row, 100);
			//Assert.AreEqual(range.End.Address, "D100");

			//ExcelAddress addr = new ExcelAddressBase("B1:D3");

			//Assert.AreEqual(addr.Start.Column, 2);
			//Assert.AreEqual(addr.Start.Row, 1);
			//Assert.AreEqual(addr.End.Column, 4);
			//Assert.AreEqual(addr.End.Row, 3);
		}

		//[Ignore]
		//[TestMethod]
		public void RichTextCells()
		{
			ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("RichText");
			var rs = ws.Cells["A1"].RichText;

			var r1 = rs.Add("Test");
			r1.Bold = true;
			r1.Color = Color.Pink;

			var r2 = rs.Add(" of");
			r2.Size = 14;
			r2.Italic = true;

			var r3 = rs.Add(" rich");
			r3.FontName = "Arial";
			r3.Size = 18;
			r3.Italic = true;

			var r4 = rs.Add("text.");
			r4.Size = 8.25f;
			r4.Italic = true;
			r4.UnderLine = true;

			var rIns = rs.Insert(2, " inserted");
			rIns.Bold = true;
			rIns.Color = Color.Green;

			rs = ws.Cells["A3:A4"].RichText;

			var r5 = rs.Add("Double");
			r5.Color = Color.PeachPuff;
			r5.FontName = "times new roman";
			r5.Size = 16;

			var r6 = rs.Add(" cells");
			r6.Color = Color.Red;
			r6.UnderLine = true;


			rs = ws.Cells["C8"].RichText;
			r1 = rs.Add("Blue ");
			r1.Color = Color.Blue;

			r2 = rs.Add("Red");
			r2.Color = Color.Red;

			ws.Cells["G1"].RichText.Add("Room 02 & 03");
			ws.Cells["G2"].RichText.Text = "Room 02 & 03";

			ws = ws = _pck.Workbook.Worksheets.Add("RichText2");
			ws.Cells["A1"].RichText.Text = "Room 02 & 03";
			ws.TabColor = Color.PowderBlue;

			r1 = ws.Cells["G3"].RichText.Add("Test");
			r1.Bold = true;
			ws.Cells["G3"].RichText.Add(" a new t");
			ws.Cells["G3"].RichText[1].Bold = false;
		}

		//[Ignore]
		//[TestMethod]
		public void TestComments()
		{
			var ws = _pck.Workbook.Worksheets.Add("Comment");
			var comment = ws.Comments.Add(ws.Cells["C3"], "Jan Källman\r\nAuthor\r\n", "JK");
			comment.RichText[0].Bold = true;
			comment.RichText[0].PreserveSpace = true;
			var rt = comment.RichText.Add("Test comment");
			comment = ws.Comments.Add(ws.Cells["A2"], "Jan Källman\r\nAuthor\r\n1", "JK");

			comment = ws.Comments.Add(ws.Cells["A1"], "Jan Källman\r\nAuthor\r\n2", "JK");
			comment = ws.Comments.Add(ws.Cells["C2"], "Jan Källman\r\nAuthor\r\n3", "JK");
			comment = ws.Comments.Add(ws.Cells["C1"], "Jan Källman\r\nAuthor\r\n5", "JK");
			comment = ws.Comments.Add(ws.Cells["B1"], "Jan Källman\r\nAuthor\r\n7", "JK");

			ws.Comments.Remove(ws.Cells["A2"].Comment);
			var rt2 = ws.Cells["B2"].AddComment("Range Added Comment test test test test test test test test test test testtesttesttesttesttesttesttesttesttesttest", "Jan Källman");
		}

		public void Address()
		{
			var ws = _pck.Workbook.Worksheets.Add("Address");
			ws.Cells["A1:A4,B5:B7"].Value = "AddressTest";
			ws.Cells["A1:A4,B5:B7"].Style.Font.Color.SetColor(Color.Red);
			ws.Cells["A2:A3,B4:B8"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightUp;
			ws.Cells["A2:A3,B4:B8"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
			ws.Cells["2:2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
			ws.Cells["2:2"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
			ws.Cells["B:B"].Style.Font.Name = "Times New Roman";

			ws.Cells["C4:G4,H8:H30,B15"].FormulaR1C1 = "RC[-1]+R1C[-1]";
			ws.Cells["C4:G4,H8:H30,B15"].Style.Numberformat.Format = "#,##0.000";
			ws.Cells["G1,G3"].Hyperlink = new ExcelHyperLink("Comment!$A$1", "Comment");
			ws.Cells["G1,G3"].Style.Font.Color.SetColor(Color.Blue);
			ws.Cells["G1,G3"].Style.Font.UnderLine = true;

			ws.Cells["A1:G5"].Copy(ws.Cells["A50"]);

			var ws2 = _pck.Workbook.Worksheets.Add("Copy Cells");
			ws.Cells["1:4"].Copy(ws2.Cells["1:1"]);

			ws.Cells["H1:J5"].Merge = true;
			ws.Cells["2:3"].Copy(ws.Cells["50:51"]);

			ws.Cells["A40"].Value = new string(new char[] { (char)8, (char)9 });

			ExcelRange styleRng = ws.Cells["A1"];
			ExcelStyle tempStyle = styleRng.Style;
			var namedStyle = _pck.Workbook.Styles.CreateNamedStyle("HyperLink", tempStyle);
			namedStyle.Style.Font.UnderLineType = ExcelUnderLineType.Single;
			namedStyle.Style.Font.Color.SetColor(Color.Blue);
		}

		public void Encoding()
		{
			var ws = _pck.Workbook.Worksheets.Add("Encoding");
			ws.Cells["A1"].Value = "_x0099_";
			ws.Cells["A2"].Value = " Test \b" + (char)1 + " end\"";
			ws.Cells["A3"].Value = "_x0097_ test_x001D_1234";
			ws.Cells["A4"].Value = "test" + (char)31;   //Bug issue 14689 //Fixed
		}

		public void WorksheetCopy()
		{
			var ws = _pck.Workbook.Worksheets.Add("Copied Address", _pck.Workbook.Worksheets["Address"]);
			var wsCopy = _pck.Workbook.Worksheets.Add("Copied Comment", _pck.Workbook.Worksheets["Comment"]);
			Assert.AreEqual(6, _pck.Workbook.Worksheets["Comment"].Comments.Count);
			Assert.AreEqual(6, wsCopy.Comments.Count);

			ExcelPackage pck2 = new ExcelPackage();
			pck2.Workbook.Worksheets.Add("Copy From other pck", _pck.Workbook.Worksheets["Address"]);
			pck2.SaveAs(new FileInfo(_worksheetPath + "copy.xlsx"));
			pck2 = null;
			Assert.AreEqual(6, wsCopy.Comments.Count);
		}

		[Ignore]
		[TestMethod]
		public void TestDelete()
		{
			string file = _worksheetPath + "test.xlsx";

			if (File.Exists(file))
				File.Delete(file);

			Create(file);

			ExcelPackage pack = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w = pack.Workbook.Worksheets["delete"];
			w.DeleteRow(1, 2);

			pack.Save();
		}

		public void LoadFromCollectionTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("LoadFromCollection");
			List<TestDTO> list = new List<TestDTO>();
			list.Add(new TestDTO() { Id = 1, Name = "Item1", Boolean = false, Date = new DateTime(2011, 1, 1), dto = null, NameVar = "Field 1" });
			list.Add(new TestDTO() { Id = 2, Name = "Item2", Boolean = true, Date = new DateTime(2011, 1, 15), dto = new TestDTO(), NameVar = "Field 2" });
			list.Add(new TestDTO() { Id = 3, Name = "Item3", Boolean = false, Date = new DateTime(2011, 2, 1), dto = null, NameVar = "Field 3" });
			list.Add(new TestDTO() { Id = 4, Name = "Item4", Boolean = true, Date = new DateTime(2011, 4, 19), dto = list[1], NameVar = "Field 4" });
			list.Add(new TestDTO() { Id = 5, Name = "Item5", Boolean = false, Date = new DateTime(2011, 5, 8), dto = null, NameVar = "Field 5" });
			list.Add(new TestDTO() { Id = 6, Name = "Item6", Boolean = true, Date = new DateTime(2010, 3, 27), dto = null, NameVar = "Field 6" });
			list.Add(new TestDTO() { Id = 7, Name = "Item7", Boolean = false, Date = new DateTime(2009, 1, 5), dto = list[3], NameVar = "Field 7" });
			list.Add(new TestDTO() { Id = 8, Name = "Item8", Boolean = true, Date = new DateTime(2018, 12, 31), dto = null, NameVar = "Field 8" });
			list.Add(new TestDTO() { Id = 9, Name = "Item9", Boolean = false, Date = new DateTime(2010, 2, 1), dto = null, NameVar = "Field 9" });

			ws.Cells["A1"].LoadFromCollection(list, true);
			ws.Cells["A30"].LoadFromCollection(list, true, OfficeOpenXml.Table.TableStyles.Medium9, BindingFlags.Instance | BindingFlags.Instance, typeof(TestDTO).GetFields());

			ws.Cells["A45"].LoadFromCollection(list, true, OfficeOpenXml.Table.TableStyles.Light1, BindingFlags.Instance | BindingFlags.Instance, new MemberInfo[] { typeof(TestDTO).GetMethod("GetNameID"), typeof(TestDTO).GetField("NameVar") });
			ws.Cells["J1"].LoadFromCollection(from l in list where l.Boolean orderby l.Date select new { Name = l.Name, Id = l.Id, Date = l.Date, NameVariable = l.NameVar }, true, OfficeOpenXml.Table.TableStyles.Dark4);

			var ints = new int[] { 1, 3, 4, 76, 2, 5 };
			ws.Cells["A15"].Value = ints;
		}

		public void LoadFromEmptyCollectionTest()
		{
			if (_pck == null) _pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("LoadFromEmpyCollection");
			List<TestDTO> listDTO = new List<TestDTO>(0);

			ws.Cells["A1"].LoadFromCollection(listDTO, true);
			ws.Cells["A5"].LoadFromCollection(listDTO, true, OfficeOpenXml.Table.TableStyles.Medium9, BindingFlags.Instance | BindingFlags.Instance, typeof(TestDTO).GetFields());

			ws.Cells["A10"].LoadFromCollection(listDTO, true, OfficeOpenXml.Table.TableStyles.Light1, BindingFlags.Instance | BindingFlags.Instance, new MemberInfo[] { typeof(TestDTO).GetMethod("GetNameID"), typeof(TestDTO).GetField("NameVar") });
			ws.Cells["A15"].LoadFromCollection(from l in listDTO where l.Boolean orderby l.Date select new { Name = l.Name, Id = l.Id, Date = l.Date, NameVariable = l.NameVar }, true, OfficeOpenXml.Table.TableStyles.Dark4);

			ws.Cells["A20"].LoadFromCollection(listDTO, false);
		}

		[TestMethod]
		public void LoadFromOneCollectionTest()
		{
			if (_pck == null) _pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("LoadFromEmpyCollection");
			List<TestDTO> listDTO = new List<TestDTO>(0) { new TestDTO() { Name = "Single" } };

			var r = ws.Cells["A1"].LoadFromCollection(listDTO, true);
			Assert.AreEqual(2, r.Rows);
			var r2 = ws.Cells["A5"].LoadFromCollection(listDTO, false);
			Assert.AreEqual(1, r2.Rows);
		}

		static void Create(string file)
		{
			ExcelPackage pack = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w = pack.Workbook.Worksheets.Add("delete");
			w.Cells[1, 1].Value = "test";
			w.Cells[1, 2].Value = "test";
			w.Cells[2, 1].Value = "to delete";
			w.Cells[2, 2].Value = "to delete";
			w.Cells[3, 1].Value = "3Left";
			w.Cells[3, 2].Value = "3Left";
			w.Cells[4, 1].Formula = "B3+C3";
			w.Cells[4, 2].Value = "C3+D3";
			pack.Save();
		}

		[Ignore]
		[TestMethod]
		public void RowStyle()
		{
			FileInfo newFile = new FileInfo(_worksheetPath + @"sample8.xlsx");
			if (newFile.Exists)
			{
				newFile.Delete();  // ensures we create a new workbook
				//newFile = new FileInfo(dir + @"sample8.xlsx");
			}

			ExcelPackage package = new ExcelPackage();
			//Load the sheet with one string column, one date column and a few random numbers.
			var ws = package.Workbook.Worksheets.Add("First line test");

			ws.Cells[1, 1].Value = "1; 1";
			ws.Cells[2, 1].Value = "2; 1";
			ws.Cells[1, 2].Value = "1; 2";
			ws.Cells[2, 2].Value = "2; 2";

			ws.Row(1).Style.Font.Bold = true;
			ws.Column(1).Style.Font.Bold = true;
			package.SaveAs(newFile);

		}

		public void HideTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("Hidden");
			ws.Cells["A1"].Value = "This workbook is hidden";
			ws.Hidden = eWorkSheetHidden.Hidden;
		}

		public void Hyperlink()
		{
			var ws = _pck.Workbook.Worksheets.Add("HyperLinks");
			var hl = new ExcelHyperLink("G1", "Till G1");
			hl.ToolTip = "Link to cell G1";
			ws.Cells["A1"].Hyperlink = hl;
		}

		public void VeryHideTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("VeryHidden");
			ws.Cells["a1"].Value = "This workbook is hidden";
			ws.Hidden = eWorkSheetHidden.VeryHidden;
		}

		public void PrinterSettings()
		{
			var ws = _pck.Workbook.Worksheets.Add("Sod/Hydroseed");

			ws.Cells[1, 1].Value = "1; 1";
			ws.Cells[2, 1].Value = "2; 1";
			ws.Cells[1, 2].Value = "1; 2";
			ws.Cells[2, 2].Value = "2; 2";
			ws.Cells[1, 1, 1, 2].AutoFilter = true;
			ws.PrinterSettings.BlackAndWhite = true;
			ws.PrinterSettings.ShowGridLines = true;
			ws.PrinterSettings.ShowHeaders = true;
			ws.PrinterSettings.PaperSize = ePaperSize.A4;

			ws.PrinterSettings.RepeatRows = new ExcelAddress("1:1");
			ws.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");

			ws.PrinterSettings.Draft = true;
			var r = ws.Cells["A26"];
			r.Value = "X";
			r.Worksheet.Row(26).PageBreak = true;
			ws.PrinterSettings.PrintArea = ws.Cells["A1:B2"];
			ws.PrinterSettings.HorizontalCentered = true;
			ws.PrinterSettings.VerticalCentered = true;

			ws.Select(new ExcelAddress("3:4,E5:F6"));

			ws = _pck.Workbook.Worksheets["RichText"];
			ws.PrinterSettings.RepeatColumns = ws.Cells["A:B"];
			ws.PrinterSettings.RepeatRows = ws.Cells["1:11"];
			ws.PrinterSettings.TopMargin = 1M;
			ws.PrinterSettings.LeftMargin = 1M;
			ws.PrinterSettings.BottomMargin = 1M;
			ws.PrinterSettings.RightMargin = 1M;
			ws.PrinterSettings.Orientation = eOrientation.Landscape;
			ws.PrinterSettings.PaperSize = ePaperSize.A4;
		}

		public void StyleNameTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("StyleNameTest");

			ws.Cells[1, 1].Value = "R1 C1";
			ws.Cells[1, 2].Value = "R1 C2";
			ws.Cells[1, 3].Value = "R1 C3";
			ws.Cells[2, 1].Value = "R2 C1";
			ws.Cells[2, 2].Value = "R2 C2";
			ws.Cells[2, 3].Value = "R2 C3";
			ws.Cells[3, 1].Value = double.PositiveInfinity;
			ws.Cells[3, 2].Value = double.NegativeInfinity;
			ws.Cells[4, 1].CreateArrayFormula("A1+B1");
			var ns = _pck.Workbook.Styles.CreateNamedStyle("TestStyle");
			ns.Style.Font.Bold = true;

			ws.Cells.Style.Locked = true;
			ws.Cells["A1:C1"].StyleName = "TestStyle";
			ws.DefaultRowHeight = 35;
			ws.Cells["A1:C4"].Style.Locked = false;
			ws.Protection.IsProtected = true;
		}

		public void ValueError()
		{
			var ws = _pck.Workbook.Worksheets.Add("ValueError");

			ws.Cells[1, 1].Value = "Domestic Violence&#xB; and the Professional";
			var rt = ws.Cells[1, 2].RichText.Add("Domestic Violence&#xB; and the Professional 2");
			TestContext.WriteLine(rt.Bold.ToString());
			rt.Bold = true;
			TestContext.WriteLine(rt.Bold.ToString());
		}

		public void FormulaError()
		{
			var ws = _pck.Workbook.Worksheets.Add("FormulaError");

			ws.Cells["D5"].Formula = "COUNTIF(A1:A100,\"Miss\")";
			ws.Cells["A1:K3"].Formula = "A3+A4";
			ws.Cells["A4"].FormulaR1C1 = "+ROUNDUP(RC[1]/10,0)*10";

			ws = _pck.Workbook.Worksheets.Add("Sheet-RC1");
			ws.Cells["A4"].FormulaR1C1 = "+ROUNDUP('Sheet-RC1'!RC[1]/10,0)*10";
		}

		[TestMethod, Ignore]
		public void FormulaArray()
		{
			_pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("FormulaError");

			ws.Cells["E2:E5"].CreateArrayFormula("FREQUENCY(B2:B18,C2:C5)");
			_pck.SaveAs(new FileInfo("c:\\temp\\arrayformula.xlsx"));
		}

		public void PictureURL()
		{
			var ws = _pck.Workbook.Worksheets.Add("Pic URL");

			ExcelHyperLink hl = new ExcelHyperLink("http://epplus.codeplex.com");
			hl.ToolTip = "Screen Tip";

			ws.Drawings.AddPicture("Pic URI", Properties.Resources.Test1, hl);
		}

		[TestMethod]
		public void PivotTableTest()
		{
			_pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("PivotTable");
			ws.Cells["A1"].LoadFromArrays(new object[][] { new[] { "A", "B", "C", "D" } });
			ws.Cells["A2"].LoadFromArrays(new object[][]
			{
				new object [] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 },
				new object [] { 9, 8, 7 ,6, 5, 4, 3, 2, 1, 0 },
				new object [] { 1, 1, 2, 3, 5, 8, 13, 21, 34, 55}
			});
			var table = ws.Tables.Add(ws.Cells["A1:D4"], "PivotData");
			ws.PivotTables.Add(ws.Cells["G1"], ws.Cells["A1:D4"], "PivotTable");
			Assert.AreEqual("PivotStyleMedium9", ws.PivotTables["PivotTable"].StyleName);
		}

		[TestMethod]
		public void AddRowsAndColumnsUpdatesPivotTable()
		{
			FileInfo file = new FileInfo(Path.GetTempFileName());
			if (file.Exists)
			{
				file.Delete();
			}
			try
			{
				using (_pck = new ExcelPackage())
				{
					var ws = _pck.Workbook.Worksheets.Add("PivotTableSheet");
					ws.Cells["A1"].LoadFromArrays(new object[][] { new[] { "A", "B", "C", "D" } });
					ws.Cells["A2"].LoadFromArrays(new object[][]
					{
				new object [] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 },
				new object [] { 9, 8, 7 ,6, 5, 4, 3, 2, 1, 0 },
				new object [] { 1, 1, 2, 3, 5, 8, 13, 21, 34, 55}
					});
					var table = ws.Tables.Add(ws.Cells["A1:D4"], "PivotData");
					var pt = ws.PivotTables.Add(ws.Cells["J10"], ws.Cells["A1:D4"], "PivotTable");
					Assert.AreEqual("PivotStyleMedium9", ws.PivotTables["PivotTable"].StyleName);
					Assert.AreEqual(10, pt.Address.Start.Row);
					Assert.AreEqual(10, pt.Address.Start.Column);
					ws.InsertRow(9, 10);
					Assert.AreEqual(20, pt.Address.Start.Row);
					ws.InsertColumn(9, 10);
					Assert.AreEqual(20, pt.Address.Start.Column);
					_pck.SaveAs(file);
				}
				using (_pck = new ExcelPackage(file))
				{
					var pivotTable = _pck.Workbook.Worksheets["PivotTableSheet"].PivotTables.First();
					Assert.AreEqual(20, pivotTable.Address.Start.Row);
					Assert.AreEqual(20, pivotTable.Address.Start.Column);
				}
			}
			finally
			{
				if (file.Exists)
				{
					file.Delete();
				}
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void TestTableNameCanNotStartsWithNumber()
		{
			var ws = _pck.Workbook.Worksheets.Add("Table");
			var tbl = ws.Tables.Add(ws.Cells["A1"], "5TestTable");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void TestTableNameCanNotContainWhiteSpaces()
		{
			var ws = _pck.Workbook.Worksheets.Add("Table");
			var tbl = ws.Tables.Add(ws.Cells["A1"], "Test Table");
		}

		[TestMethod]
		public void TestTableNameCanStartsWithBackSlash()
		{
			var ws = _pck.Workbook.Worksheets.Add("Table");
			var tbl = ws.Tables.Add(ws.Cells["A1"], "\\TestTable");
		}

		[TestMethod]
		public void TestTableNameCanStartsWithUnderscore()
		{
			var ws = _pck.Workbook.Worksheets.Add("Table");
			var tbl = ws.Tables.Add(ws.Cells["A1"], "_TestTable");
		}

		[TestMethod]
		public void TableTotalsRowFunctionEscapesSpecialCharactersInColumnName()
		{
			using (var p = new ExcelPackage())
			{
				var ws = p.Workbook.Worksheets.Add("TotalsFormulaTest");
				ws.Cells["B1"].Value = "Column1";
				ws.Cells["C1"].Value = "[#'Column'2]";
				var tbl = ws.Tables.Add(ws.Cells["B1:C2"], "TestTable");
				tbl.ShowTotal = true;
				tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
				Assert.AreEqual("SUBTOTAL(109,TestTable['['#''Column''2']])", ws.Cells["C3"].Formula);
			}
		}

		[TestMethod]
		public void InsertRowsSetsOutlineLevel()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Row(15).OutlineLevel = 1;
				sheet1.InsertRow(2, 10, 15);
				for (int i = 2; i < 12; i++)
				{
					Assert.AreEqual(1, sheet1.Row(i).OutlineLevel, $"The outline level of row {i} is not set.");
				}
				Assert.AreEqual(1, sheet1.Row(25).OutlineLevel);
			}
		}

		[TestMethod]
		public void InsertColumnsSetsOutlineLevel()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Column(15).OutlineLevel = 1;
				sheet1.InsertColumn(2, 10, 15);
				for (int i = 2; i < 12; i++)
				{
					Assert.AreEqual(1, sheet1.Column(i).OutlineLevel, $"The outline level of column {i} is not set.");
				}
				Assert.AreEqual(1, sheet1.Column(25).OutlineLevel);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void InsertColumnUpdatesSparklines()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet1.InsertColumn(5, 1);
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:G6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("H6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D7:G7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("H7", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D8:G8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("H8", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!F6:F8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!G6:G8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("G9", sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:G6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("H6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D7:G7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("H7", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D8:G8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("H8", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!F6:F8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!G6:G8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("G9", sparklines[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void InsertRowUpdatesSparklines()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet1.InsertRow(7, 1);
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C13", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D8:F8", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D9:F9", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G9", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D9", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D10", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!E6:E9", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E10", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!F6:F9", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F10", sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C13", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D8:F8", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D9:F9", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G9", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D9", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D10", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!E6:E9", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E10", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!F6:F9", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F10", sparklines[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		public void CopyRowSetsOutlineLevelsCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Row(2).OutlineLevel = 1;
				sheet1.Row(3).OutlineLevel = 1;
				sheet1.Row(4).OutlineLevel = 0;

				// Set outline levels on rows to be copied over.
				sheet1.Row(6).OutlineLevel = 17;
				sheet1.Row(7).OutlineLevel = 25;
				sheet1.Row(8).OutlineLevel = 29;

				sheet1.Cells["2:4"].Copy(sheet1.Cells["A6"]);
				Assert.AreEqual(1, sheet1.Row(2).OutlineLevel);
				Assert.AreEqual(1, sheet1.Row(3).OutlineLevel);
				Assert.AreEqual(0, sheet1.Row(4).OutlineLevel);

				Assert.AreEqual(1, sheet1.Row(6).OutlineLevel);
				Assert.AreEqual(1, sheet1.Row(7).OutlineLevel);
				Assert.AreEqual(0, sheet1.Row(8).OutlineLevel);
			}
		}

		[TestMethod]
		public void CopyRowCrossSheetSetsOutlineLevelsCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Row(2).OutlineLevel = 1;
				sheet1.Row(3).OutlineLevel = 1;
				sheet1.Row(4).OutlineLevel = 0;

				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Set outline levels on rows to be copied over.
				sheet2.Row(6).OutlineLevel = 17;
				sheet2.Row(7).OutlineLevel = 25;
				sheet2.Row(8).OutlineLevel = 29;

				sheet1.Cells["2:4"].Copy(sheet2.Cells["A6"]);
				Assert.AreEqual(1, sheet1.Row(2).OutlineLevel);
				Assert.AreEqual(1, sheet1.Row(3).OutlineLevel);
				Assert.AreEqual(0, sheet1.Row(4).OutlineLevel);

				Assert.AreEqual(1, sheet2.Row(6).OutlineLevel);
				Assert.AreEqual(1, sheet2.Row(7).OutlineLevel);
				Assert.AreEqual(0, sheet2.Row(8).OutlineLevel);
			}
		}

		[TestMethod]
		public void CopyColumnSetsOutlineLevelsCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Column(2).OutlineLevel = 1;
				sheet1.Column(3).OutlineLevel = 1;
				sheet1.Column(4).OutlineLevel = 0;

				// Set outline levels on rows to be copied over.
				sheet1.Column(6).OutlineLevel = 17;
				sheet1.Column(7).OutlineLevel = 25;
				sheet1.Column(8).OutlineLevel = 29;

				sheet1.Cells["B:D"].Copy(sheet1.Cells["F1"]);
				Assert.AreEqual(1, sheet1.Column(2).OutlineLevel);
				Assert.AreEqual(1, sheet1.Column(3).OutlineLevel);
				Assert.AreEqual(0, sheet1.Column(4).OutlineLevel);

				Assert.AreEqual(1, sheet1.Column(6).OutlineLevel);
				Assert.AreEqual(1, sheet1.Column(7).OutlineLevel);
				Assert.AreEqual(0, sheet1.Column(8).OutlineLevel);
			}
		}

		public void TableTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("Table");
			ws.Cells["B1"].Value = 123;
			var tbl = ws.Tables.Add(ws.Cells["B1:P12"], "TestTable");
			tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Custom;

			tbl.ShowFirstColumn = true;
			tbl.ShowTotal = true;
			tbl.ShowHeader = true;
			tbl.ShowLastColumn = true;
			tbl.ShowFilter = false;
			Assert.AreEqual(tbl.ShowFilter, false);
			ws.Cells["K2"].Value = 5;
			ws.Cells["J3"].Value = 4;

			tbl.Columns[8].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
			tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])", tbl.Columns[9].Name);
			tbl.Columns[14].CalculatedColumnFormula = "TestTable[[#This Row],[123]]+TestTable[[#This Row],[Column2]]";
			ws.Cells["B2"].Value = 1;
			ws.Cells["B3"].Value = 2;
			ws.Cells["B4"].Value = 3;
			ws.Cells["B5"].Value = 4;
			ws.Cells["B6"].Value = 5;
			ws.Cells["B7"].Value = 6;
			ws.Cells["B8"].Value = 7;
			ws.Cells["B9"].Value = 8;
			ws.Cells["B10"].Value = 9;
			ws.Cells["B11"].Value = 10;
			ws.Cells["B12"].Value = 11;
			ws.Cells["C7"].Value = "Table test";
			ws.Cells["C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
			ws.Cells["C8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

			tbl = ws.Tables.Add(ws.Cells["a12:a13"], "");

			tbl = ws.Tables.Add(ws.Cells["C16:Y35"], "");
			tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;
			tbl.ShowFirstColumn = true;
			tbl.ShowLastColumn = true;
			tbl.ShowColumnStripes = true;
			Assert.AreEqual(tbl.ShowFilter, true);
			tbl.Columns[2].Name = "Test Column Name";

			ws.Cells["G50"].Value = "Timespan";
			ws.Cells["G51"].Value = new DateTime(new TimeSpan(1, 1, 10).Ticks); //new DateTime(1899, 12, 30, 1, 1, 10);
			ws.Cells["G52"].Value = new DateTime(1899, 12, 30, 2, 3, 10);
			ws.Cells["G53"].Value = new DateTime(1899, 12, 30, 3, 4, 10);
			ws.Cells["G54"].Value = new DateTime(1899, 12, 30, 4, 5, 10);

			ws.Cells["G51:G55"].Style.Numberformat.Format = "HH:MM:SS";
			tbl = ws.Tables.Add(ws.Cells["G50:G54"], "");
			tbl.ShowTotal = true;
			tbl.ShowFilter = false;
			tbl.Columns[0].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
		}

		[TestMethod]
		public void TableDeleteTest()
		{
			using (var pkg = new ExcelPackage())
			{
				var wb = pkg.Workbook;
				var sheets = new[]
				{
					wb.Worksheets.Add("WorkSheet A"),
					wb.Worksheets.Add("WorkSheet B")
				};
				for (int i = 1; i <= 4; i++)
				{
					var cell = sheets[0].Cells[1, i];
					cell.Value = cell.Address + "_";
					cell = sheets[1].Cells[1, i];
					cell.Value = cell.Address + "_";
				}

				for (int i = 6; i <= 11; i++)
				{
					var cell = sheets[0].Cells[3, i];
					cell.Value = cell.Address + "_";
					cell = sheets[1].Cells[3, i];
					cell.Value = cell.Address + "_";
				}
				var tables = new[]
				{
					sheets[1].Tables.Add(sheets[1].Cells["A1:D73"], "Tablea"),
					sheets[0].Tables.Add(sheets[0].Cells["A1:D73"], "Table2"),
					sheets[1].Tables.Add(sheets[1].Cells["F3:K10"], "Tableb"),
					sheets[0].Tables.Add(sheets[0].Cells["F3:K10"], "Table3"),
				};
				Assert.AreEqual(5, wb.NextTableID);
				Assert.AreEqual(1, tables[0].Id);
				Assert.AreEqual(2, tables[1].Id);
				try
				{
					sheets[0].Tables.Delete("Tablea");
					Assert.Fail("ArgumentException should have been thrown.");
				}
				catch (ArgumentOutOfRangeException) { }
				sheets[1].Tables.Delete("Tablea");
				Assert.AreEqual(1, tables[1].Id);
				Assert.AreEqual(2, tables[2].Id);

				try
				{
					sheets[1].Tables.Delete(4);
					Assert.Fail("ArgumentException should have been thrown.");
				}
				catch (ArgumentOutOfRangeException) { }
				var range = sheets[0].Cells[sheets[0].Tables[1].Address.Address];
				sheets[0].Tables.Delete(1, true);
				foreach (var cell in range)
				{
					Assert.IsNull(cell.Value);
				}
			}
		}

		[TestMethod]
		public void TableWithSubtotalsParensInColumnName()
		{
			var ws = _pck.Workbook.Worksheets.Add("Table");
			ws.Cells["B2"].Value = "Header 1";
			ws.Cells["C2"].Value = "Header (2)";
			ws.Cells["B3"].Value = 1;
			ws.Cells["B4"].Value = 2;
			ws.Cells["C3"].Value = 3;
			ws.Cells["C4"].Value = 4;
			var table = ws.Tables.Add(ws.Cells["B2:C4"], "TestTable");
			table.ShowTotal = true;
			table.ShowHeader = true;
			table.Columns[0].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
			table.Columns[1].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
			ws.Cells["B5"].Calculate();
			Assert.AreEqual(3.0, ws.Cells["B5"].Value);
			ws.Cells["C5"].Calculate();
			Assert.AreEqual(7.0, ws.Cells["C5"].Value);
		}

		//[Ignore]
		//[TestMethod]
		public void CopyTable()
		{
			_pck.Workbook.Worksheets.Copy("File4", "Copied table");
		}

		//[Ignore]
		//[TestMethod]
		public void CopyRange()
		{
			var ws = _pck.Workbook.Worksheets.Add("CopyTest");

			ws.Cells["A1"].Value = "Single Cell";
			ws.Cells["A2"].Value = "Merged Cells";
			ws.Cells["A2:D30"].Merge = true;
			ws.Cells["A1"].Style.Font.Bold = true;
			ws.Cells["G4:H5"].Merge = true;
			ws.Cells["B3:C5"].Copy(ws.Cells["G4"]);
		}

		//[Ignore]
		//[TestMethod]
		public void CopyMergedRange()
		{
			var ws = _pck.Workbook.Worksheets.Add("CopyMergedRangeTest");

			ws.Cells["A11:C11"].Merge = true;
			ws.Cells["A12:C12"].Merge = true;

			var source = ws.Cells["A11:C12"];
			var target = ws.Cells["A21"];

			source.Copy(target);

			var a21 = ws.Cells[21, 1];
			var a22 = ws.Cells[22, 1];

			Assert.IsTrue(a21.Merge);
			Assert.IsTrue(a22.Merge);

			//Assert.AreNotEqual(a21.MergeId, a22.MergeId);
		}

		[Ignore]
		[TestMethod]
		public void CopyPivotTable()
		{
			_pck.Workbook.Worksheets.Copy("Pivot-Group Date", "Copied Pivottable 1");
			_pck.Workbook.Worksheets.Copy("Pivot-Group Number", "Copied Pivottable 2");
		}

		[Ignore]
		[TestMethod]
		public void Stylebug()
		{
			ExcelPackage p = new ExcelPackage(new FileInfo(@"c:\temp\FullProjecte.xlsx"));

			var ws = p.Workbook.Worksheets.First();
			ws.Cells[12, 1].Value = 0;
			ws.Cells[12, 2].Value = new DateTime(2010, 9, 14);
			ws.Cells[12, 3].Value = "Federico Lois";
			ws.Cells[12, 4].Value = "Nakami";
			ws.Cells[12, 5].Value = "Hores";
			ws.Cells[12, 7].Value = 120;
			ws.Cells[12, 8].Value = "A definir";
			ws.Cells[12, 9].Value = new DateTime(2010, 9, 14);
			ws.Cells[12, 10].Value = new DateTime(2010, 9, 14);
			ws.Cells[12, 11].Value = "Transferència";

			ws.InsertRow(13, 1, 12);
			ws.Cells[13, 1].Value = 0;
			ws.Cells[13, 2].Value = new DateTime(2010, 9, 14);
			ws.Cells[13, 3].Value = "Federico Lois";
			ws.Cells[13, 4].Value = "Nakami";
			ws.Cells[13, 5].Value = "Hores";
			ws.Cells[13, 7].Value = 120;
			ws.Cells[13, 8].Value = "A definir";
			ws.Cells[13, 9].Value = new DateTime(2010, 9, 14);
			ws.Cells[13, 10].Value = new DateTime(2010, 9, 14);
			ws.Cells[13, 11].Value = "Transferència";

			ws.InsertRow(14, 1, 13);

			ws.InsertRow(19, 1, 19);
			ws.InsertRow(26, 1, 26);
			ws.InsertRow(33, 1, 33);
			p.SaveAs(new FileInfo(@"c:\temp\FullProjecte_new.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void ReadBug()
		{
			using (var package = new ExcelPackage(new FileInfo(@"c:\temp\error.xlsx")))
			{
				var fulla = package.Workbook.Worksheets.FirstOrDefault();
				var r = fulla == null ? null : fulla.Cells["a:a"]
				.Where(t => !string.IsNullOrWhiteSpace(t.Text)).Select(cell => cell.Value.ToString())
				.ToList();
			}
		}

		//[Ignore]
		//[TestMethod]
		public void FormulaOverwrite()
		{
			var ws = _pck.Workbook.Worksheets.Add("FormulaOverwrite");
			//Inside
			ws.Cells["A1:G12"].Formula = "B1+C1";
			ws.Cells["B2:C3"].Formula = "G2+E1";


			//Top bottom overwrite
			ws.Cells["A14:G26"].Formula = "B1+C1+D1";
			ws.Cells["B13:C28"].Formula = "G2+E1";

			//Top bottom overwrite
			ws.Cells["B30:E42"].Formula = "B1+C1+$D$1";
			ws.Cells["A32:H33"].Formula = "G2+E1";

			ws.Cells["A50:A59"].CreateArrayFormula("C50+D50");
			ws.Cells["A1"].Value = "test";
			ws.Cells["A15"].Value = "Värde";
			ws.Cells["C12"].AddComment("Test", "JJOD");
			ws.Cells["D12:I12"].Merge = true;
			ws.Cells["D12:I12"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
			ws.Cells["D12:I12"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
			ws.Cells["D12:I12"].Style.WrapText = true;

			ws.Cells["F1:F3"].Formula = "F2+F3";
			ws.Cells["J1:J3"].Formula = "F2+F3";
			ws.Cells["F1:F3"].Formula = "F5+F6";    //Overwrite same range
		}

		//[Ignore]
		//[TestMethod]
		public void DefinedName()
		{
			var ws = _pck.Workbook.Worksheets.Add("Names");
			ws.Names.Add("RefError", ws.Cells["#REF!"]);

			ws.Cells["A1"].Value = "Test";
			ws.Cells["A1"].Style.Font.Size = 8.5F;

			ws.Names.Add("Address", ws.Cells["A2:A3"]);
			ws.Names["Address"].NameFormula = "1";
			ws.Names.Add("Value", "5");
			ws.Names.Add("FullRow", ws.Cells["2:2"]);
			ws.Names.Add("FullCol", ws.Cells["A:A"]);
			ws.Names.Add("Formula", "Names!A2+Names!A3+Names!Value");
		}

		[Ignore]
		[TestMethod]
		public void URL()
		{
			var p = new ExcelPackage(new FileInfo(@"c:\temp\url.xlsx"));
			foreach (var ws in p.Workbook.Worksheets)
			{

			}
			p.SaveAs(new FileInfo(@"c:\temp\urlsaved.xlsx"));
		}

		//[TestMethod]
		public void LoadDataReader()
		{
			if (_pck == null) _pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("Loaded DataReader");
			ExcelRangeBase range;
			using (var dt = new DataTable())
			{
				dt.Columns.Add("String", typeof(string));
				dt.Columns.Add("Int", typeof(int));
				dt.Columns.Add("Bool", typeof(bool));
				dt.Columns.Add("Double", typeof(double));

				var dr = dt.NewRow();
				dr[0] = "Row1";
				dr[1] = 1;
				dr[2] = true;
				dr[3] = 1.5;
				dt.Rows.Add(dr);

				dr = dt.NewRow();
				dr[0] = "Row2";
				dr[1] = 2;
				dr[2] = false;
				dr[3] = 2.25;
				dt.Rows.Add(dr);

				//dr = dt.NewRow();
				//dr[0] = "Row3";
				//dr[1] = 3;
				//dr[2] = true;
				//dr[3] = 3.125;
				//dt.Rows.Add(dr);

				using (var reader = dt.CreateDataReader())
				{
					range = ws.Cells["A1"].LoadFromDataReader(reader, true, "My_Table",
															  OfficeOpenXml.Table.TableStyles.Medium5);
				}
				Assert.AreEqual(1, range.Start.Column);
				Assert.AreEqual(4, range.End.Column);
				Assert.AreEqual(1, range.Start.Row);
				Assert.AreEqual(3, range.End.Row);

				using (var reader = dt.CreateDataReader())
				{
					range = ws.Cells["A5"].LoadFromDataReader(reader, false, "My_Table2",
															  OfficeOpenXml.Table.TableStyles.Medium5);
				}
				Assert.AreEqual(1, range.Start.Column);
				Assert.AreEqual(4, range.End.Column);
				Assert.AreEqual(5, range.Start.Row);
				Assert.AreEqual(6, range.End.Row);
			}
		}

		//[TestMethod, Ignore]
		public void LoadDataTable()
		{
			if (_pck == null) _pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("Loaded DataTable");

			var dt = new DataTable();
			dt.Columns.Add("String", typeof(string));
			dt.Columns.Add("Int", typeof(int));
			dt.Columns.Add("Bool", typeof(bool));
			dt.Columns.Add("Double", typeof(double));


			var dr = dt.NewRow();
			dr[0] = "Row1";
			dr[1] = 1;
			dr[2] = true;
			dr[3] = 1.5;
			dt.Rows.Add(dr);

			dr = dt.NewRow();
			dr[0] = "Row2";
			dr[1] = 2;
			dr[2] = false;
			dr[3] = 2.25;
			dt.Rows.Add(dr);

			dr = dt.NewRow();
			dr[0] = "Row3";
			dr[1] = 3;
			dr[2] = true;
			dr[3] = 3.125;
			dt.Rows.Add(dr);

			ws.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Medium5);

			//worksheet.Cells[startRow, 7, worksheet.Dimension.End.Row, 7].FormulaR1C1 = "=IF(RC[-2]=0,0,RC[-1]/RC[-2])";

			ws.Tables[0].Columns[1].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
			ws.Tables[0].ShowTotal = true;
		}

		[Ignore]
		[TestMethod]
		public void LoadEmptyDataTable()
		{
			if (_pck == null) _pck = new ExcelPackage();
			var ws = _pck.Workbook.Worksheets.Add("Loaded Empty DataTable");

			var dt = new DataTable();
			dt.Columns.Add(new DataColumn("col1"));
			dt.Columns.Add(new DataColumn("col2"));
			ws.Cells["A1"].LoadFromDataTable(dt, true);

			ws.Cells["D1"].LoadFromDataTable(dt, false);
		}

		[TestMethod]
		public void LoadText_Bug15015()
		{
			var package = new ExcelPackage();
			var ws = package.Workbook.Worksheets.Add("Loaded Text");
			ws.Cells["A1"].LoadFromText("\"text with eol,\r\n in a cell\",\"other value\"", new ExcelTextFormat { TextQualifier = '"', EOL = ",\r\n", Delimiter = ',' });
		}

		[TestMethod]
		public void LoadText_Bug15015_Negative()
		{
			var package = new ExcelPackage();
			var ws = package.Workbook.Worksheets.Add("Loaded Text");
			bool exceptionThrown = false;
			try
			{
				ws.Cells["A1"].LoadFromText("\"text with eol,\r\n",
											new ExcelTextFormat { TextQualifier = '"', EOL = ",\r\n", Delimiter = ',' });
			}
			catch (Exception e)
			{
				Assert.AreEqual("Text delimiter is not closed in line : \"text with eol", e.Message, "Exception message");
				exceptionThrown = true;
			}
			Assert.IsTrue(exceptionThrown, "Exception thrown");
		}

		//[Ignore]
		//[TestMethod]
		public void LoadText()
		{
			var ws = _pck.Workbook.Worksheets.Add("Loaded Text");

			ws.Cells["A1"].LoadFromText("1.2");
			ws.Cells["A2"].LoadFromText("1,\"Test av data\",\"12,2\",\"\"Test\"\"");
			ws.Cells["A3"].LoadFromText("\"1,3\",\"Test av \"\"data\",\"12,2\",\"Test\"\"\"", new ExcelTextFormat() { TextQualifier = '"' });
			ws.Cells["A4"].LoadFromText("\"1,3\",\"\",\"12,2\",\"Test\"\"\"", new ExcelTextFormat() { TextQualifier = '"' });

			ws = _pck.Workbook.Worksheets.Add("File1");
			// ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\et1c1004.csv"), new ExcelTextFormat() {SkipLinesBeginning=3,SkipLinesEnd=1, EOL="\n"});

			ws = _pck.Workbook.Worksheets.Add("File2");
			//ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\etiv2812.csv"), new ExcelTextFormat() { SkipLinesBeginning = 3, SkipLinesEnd = 1, EOL = "\n" });

			//ws = _pck.Workbook.Worksheets.Add("File3");
			//ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\last_gics.txt"), new ExcelTextFormat() { SkipLinesBeginning = 1, Delimiter='|'});

			ws = _pck.Workbook.Worksheets.Add("File4");
			//ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\20060927.custom_open_positions.cdf.SPP"), new ExcelTextFormat() { SkipLinesBeginning = 2, SkipLinesEnd=2, TextQualifier='"', DataTypes=new eDataTypes[] {eDataTypes.Number,eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.String, eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number}},
			//    OfficeOpenXml.Table.TableStyles.Medium27, true);

			ws.Cells["A1"].LoadFromText("1,\"Test\",\"\",\"\"\"\",3", new ExcelTextFormat() { TextQualifier = '\"' });

			var style = _pck.Workbook.Styles.CreateNamedStyle("RedStyle");
			style.Style.Fill.PatternType = ExcelFillStyle.Solid;
			style.Style.Fill.BackgroundColor.SetColor(Color.Red);

			//var tbl = ws.Tables[ws.Tables.Count - 1];
			//tbl.ShowTotal = true;
			//tbl.TotalsRowCellStyle = "RedStyle";
			//tbl.HeaderRowCellStyle = "RedStyle";
		}

		[TestMethod]
		public void TestRepeatRowsAndColumnsTest()
		{
			var p = new ExcelPackage();

			var w = p.Workbook.Worksheets.Add("RepeatRowsAndColumnsTest");

			w.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");
			w.PrinterSettings.RepeatRows = new ExcelAddress("1:1");

			Assert.IsNotNull(w.PrinterSettings.RepeatColumns);
			Assert.IsNotNull(w.PrinterSettings.RepeatRows); // Fails!
		}

		//[Ignore]
		//[TestMethod]
		public void Merge()
		{
			var ws = _pck.Workbook.Worksheets.Add("Merge");
			ws.Cells["A1:A4"].Merge = true;
			ws.Cells["C1:C4,C8:C12"].Merge = true;
			ws.Cells["D13:E18,G5,U32:U45"].Merge = true;
			ws.Cells["D13:E18,G5,U32:U45"].Style.WrapText = true;
			//ws.Cells["50:52"].Merge = true;
			ws.Cells["AA:AC"].Merge = true;
			ws.SetValue(13, 4, "Merged\r\nnew row");
		}

		//[Ignore]
		//[TestMethod]
		public void DefaultColWidth()
		{
			var ws = _pck.Workbook.Worksheets.Add("DefColWidth");
			ws.DefaultColWidth = 45;
		}

		//[Ignore]
		//[TestMethod]
		public void LoadArray()
		{
			var ws = _pck.Workbook.Worksheets.Add("Loaded Array");
			List<object[]> testArray = new List<object[]>() { new object[] { 3, 4, 5, 6 }, new string[] { "Test1", "test", "5", "6" } };
			ws.Cells["A1"].LoadFromArrays(testArray);
		}

		[Ignore]
		[TestMethod]
		public void DefColWidthBug()
		{
			ExcelWorkbook book = _pck.Workbook;
			ExcelWorksheet sheet = book.Worksheets.Add("Gebruikers");

			sheet.DefaultColWidth = 25d;
			//sheet.defaultRowHeight = 15d; // needed to make sure the resulting file is valid!

			// Create the header row
			sheet.Cells[1, 1].Value = "Afdeling code";
			sheet.Cells[1, 2].Value = "Afdeling naam";
			sheet.Cells[1, 3].Value = "Voornaam";
			sheet.Cells[1, 4].Value = "Tussenvoegsel";
			sheet.Cells[1, 5].Value = "Achternaam";
			sheet.Cells[1, 6].Value = "Gebruikersnaam";
			sheet.Cells[1, 7].Value = "E-mail adres";
			ExcelRange headerRow = sheet.Cells[1, 1, 1, 7];
			headerRow.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
			headerRow.Style.Font.Size = 12;
			headerRow.Style.Font.Bold = true;

			//// Create a context for retrieving the users
			//using (PalauDataContext context = new PalauDataContext())
			//{
			//    int currentRow = 2;

			//    // iterate through all users in the export and add their info
			//    // to the worksheet.
			//    foreach (vw_ExportUser user in
			//      context.vw_ExportUsers
			//      .OrderBy(u => u.DepartmentCode)
			//      .ThenBy(u => u.AspNetUserName))
			//    {
			//        sheet.Cells[currentRow, 1].Value = user.DepartmentCode;
			//        sheet.Cells[currentRow, 2].Value = user.DepartmentName;
			//        sheet.Cells[currentRow, 3].Value = user.UserFirstName;
			//        sheet.Cells[currentRow, 4].Value = user.UserInfix;
			//        sheet.Cells[currentRow, 5].Value = user.UserSurname;
			//        sheet.Cells[currentRow, 6].Value = user.AspNetUserName;
			//        sheet.Cells[currentRow, 7].Value = user.AspNetEmail;

			//        currentRow++;
			//    }
			//}

			// return the filled Excel workbook
			//  return pkg

		}

		[Ignore]
		[TestMethod]
		public void CloseProblem()
		{
			ExcelPackage pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Manual Receipts");

			ws.Cells["A1"].Value = " SpaceNeedle Manual Receipt Form";

			using (ExcelRange r = ws.Cells["A1:F1"])
			{
				r.Merge = true;
				r.Style.Font.SetFromFont(new Font("Arial", 18, FontStyle.Italic));
				r.Style.Font.Color.SetColor(Color.DarkRed);
				r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
				//r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				//r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
			}
			//			ws.Column(1).BestFit = true;
			ws.Column(1).Width = 17;
			ws.Column(5).Width = 20;


			ws.Cells["A2"].Value = "Date Produced";

			ws.Cells["A2"].Style.Font.Bold = true;
			ws.Cells["B2"].Value = DateTime.Now.ToShortDateString();
			ws.Cells["D2"].Value = "Quantity";
			ws.Cells["D2"].Style.Font.Bold = true;
			ws.Cells["E2"].Value = "txt";

			ws.Cells["C4"].Value = "Receipt Number";
			ws.Cells["C4"].Style.WrapText = true;
			ws.Cells["C4"].Style.Font.Bold = true;

			int rowNbr = 5;
			for (int entryNbr = 1; entryNbr <= 1; entryNbr += 1)
			{
				ws.Cells["B" + rowNbr].Value = entryNbr;
				ws.Cells["C" + rowNbr].Value = 1 + entryNbr - 1;
				rowNbr += 1;
			}
			pck.SaveAs(new FileInfo(".\\test.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void OpenXlsm()
		{
			ExcelPackage p = new ExcelPackage(new FileInfo("c:\\temp\\cs1.xlsx"));
			int c = p.Workbook.Worksheets.Count;
			p.Save();
		}

		[Ignore]
		[TestMethod]
		public void Mergebug()
		{
			var xlPackage = new ExcelPackage();
			var xlWorkSheet = xlPackage.Workbook.Worksheets.Add("Test Sheet");
			var Cells = xlWorkSheet.Cells;
			var TitleCell = Cells[1, 1, 1, 3];

			TitleCell.Merge = true;
			TitleCell.Value = "Test Spreadsheet";
			Cells[2, 1].Value = "Test Sub Heading\r\ntest" + (char)22;
			for (int i = 0; i < 256; i++)
			{
				Cells[3, i + 1].Value = (char)i;
			}
			Cells[2, 1].Style.WrapText = true;
			xlWorkSheet.Row(1).Height = 50;
			xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void OpenProblem()
		{
			var xlPackage = new ExcelPackage();
			var ws = xlPackage.Workbook.Worksheets.Add("W1");
			xlPackage.Workbook.Worksheets.Add("W2");

			ws.Cells["A1:A10"].Formula = "W2!A1+C1";
			ws.Cells["B1:B10"].FormulaR1C1 = "W2!R1C1+C1";
			xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void ProtectionProblem()
		{
			var xlPackage = new ExcelPackage(new FileInfo("c:\\temp\\CovenantsCheckReportTemplate.xlsx"));
			var ws = xlPackage.Workbook.Worksheets.First();
			ws.Protection.SetPassword("Test");
			xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void Nametest()
		{
			var pck = new ExcelPackage(new FileInfo("c:\\temp\\names.xlsx"));
			var ws = pck.Workbook.Worksheets.First();
			ws.Cells["H37"].Formula = "\"Test\"";
			pck.SaveAs(new FileInfo(@"c:\\temp\\nametest_new.xlsx"));
		}

		//[Ignore]
		//[TestMethod]
		public void CreatePivotTable()
		{
			var wsPivot1 = _pck.Workbook.Worksheets.Add("Rows-Data on columns");
			var wsPivot2 = _pck.Workbook.Worksheets.Add("Rows-Data on rows");
			var wsPivot3 = _pck.Workbook.Worksheets.Add("Columns-Data on columns");
			var wsPivot4 = _pck.Workbook.Worksheets.Add("Columns-Data on rows");
			var wsPivot5 = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on columns");
			var wsPivot6 = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on rows");
			var wsPivot7 = _pck.Workbook.Worksheets.Add("Rows/Page-Data on Columns");
			var wsPivot8 = _pck.Workbook.Worksheets.Add("Pivot-Group Date");
			var wsPivot9 = _pck.Workbook.Worksheets.Add("Pivot-Group Number");
			var wsPivot10 = _pck.Workbook.Worksheets.Add("Pivot-Many RowFields");

			var ws = _pck.Workbook.Worksheets.Add("Data");
			ws.Cells["K1"].Value = "Item";
			ws.Cells["L1"].Value = "Category";
			ws.Cells["M1"].Value = "Stock";
			ws.Cells["N1"].Value = "Price";
			ws.Cells["O1"].Value = "Date for grouping";

			ws.Cells["K2"].Value = "Crowbar";
			ws.Cells["L2"].Value = "Hardware";
			ws.Cells["M2"].Value = 12;
			ws.Cells["N2"].Value = 85.2;
			ws.Cells["O2"].Value = new DateTime(2010, 1, 31);

			ws.Cells["K3"].Value = "Crowbar";
			ws.Cells["L3"].Value = "Hardware";
			ws.Cells["M3"].Value = 15;
			ws.Cells["N3"].Value = 12.2;
			ws.Cells["O3"].Value = new DateTime(2010, 2, 28);

			ws.Cells["K4"].Value = "Hammer";
			ws.Cells["L4"].Value = "Hardware";
			ws.Cells["M4"].Value = 550;
			ws.Cells["N4"].Value = 72.7;
			ws.Cells["O4"].Value = new DateTime(2010, 3, 31);

			ws.Cells["K5"].Value = "Hammer";
			ws.Cells["L5"].Value = "Hardware";
			ws.Cells["M5"].Value = 120;
			ws.Cells["N5"].Value = 11.3;
			ws.Cells["O5"].Value = new DateTime(2010, 4, 30);

			ws.Cells["K6"].Value = "Crowbar";
			ws.Cells["L6"].Value = "Hardware";
			ws.Cells["M6"].Value = 120;
			ws.Cells["N6"].Value = 173.2;
			ws.Cells["O6"].Value = new DateTime(2010, 5, 31);

			ws.Cells["K7"].Value = "Hammer";
			ws.Cells["L7"].Value = "Hardware";
			ws.Cells["M7"].Value = 1;
			ws.Cells["N7"].Value = 4.2;
			ws.Cells["O7"].Value = new DateTime(2010, 6, 30);

			ws.Cells["K8"].Value = "Saw";
			ws.Cells["L8"].Value = "Hardware";
			ws.Cells["M8"].Value = 4;
			ws.Cells["N8"].Value = 33.12;
			ws.Cells["O8"].Value = new DateTime(2010, 6, 28);

			ws.Cells["K9"].Value = "Screwdriver";
			ws.Cells["L9"].Value = "Hardware";
			ws.Cells["M9"].Value = 1200;
			ws.Cells["N9"].Value = 45.2;
			ws.Cells["O9"].Value = new DateTime(2010, 8, 31);

			ws.Cells["K10"].Value = "Apple";
			ws.Cells["L10"].Value = "Groceries";
			ws.Cells["M10"].Value = 807;
			ws.Cells["N10"].Value = 1.2;
			ws.Cells["O10"].Value = new DateTime(2010, 9, 30);

			ws.Cells["K11"].Value = "Butter";
			ws.Cells["L11"].Value = "Groceries";
			ws.Cells["M11"].Value = 52;
			ws.Cells["N11"].Value = 7.2;
			ws.Cells["O11"].Value = new DateTime(2010, 10, 31);
			ws.Cells["O2:O11"].Style.Numberformat.Format = "yyyy-MM-dd";

			var pt = wsPivot1.PivotTables.Add(wsPivot1.Cells["A1"], ws.Cells["K1:N11"], "Pivottable1");
			pt.GrandTotalCaption = "Total amount";
			pt.RowFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataFields[0].Function = DataFieldFunctions.Product;
			pt.DataOnRows = false;

			pt = wsPivot2.PivotTables.Add(wsPivot2.Cells["A1"], ws.Cells["K1:N11"], "Pivottable2");
			pt.RowFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataFields[0].Function = DataFieldFunctions.Average;
			pt.DataOnRows = true;

			pt = wsPivot3.PivotTables.Add(wsPivot3.Cells["A1"], ws.Cells["K1:N11"], "Pivottable3");
			pt.ColumnFields.Add(pt.Fields[1]);
			pt.ColumnFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = false;

			pt = wsPivot4.PivotTables.Add(wsPivot4.Cells["A1"], ws.Cells["K1:N11"], "Pivottable4");
			pt.ColumnFields.Add(pt.Fields[1]);
			pt.ColumnFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = true;

			pt = wsPivot5.PivotTables.Add(wsPivot5.Cells["A1"], ws.Cells["K1:N11"], "Pivottable5");
			pt.ColumnFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = false;

			pt = wsPivot6.PivotTables.Add(wsPivot6.Cells["A1"], ws.Cells["K1:N11"], "Pivottable6");
			pt.ColumnFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = true;
			wsPivot6.Drawings.AddChart("Pivotchart6", OfficeOpenXml.Drawing.Chart.eChartType.BarStacked3D, pt);

			pt = wsPivot7.PivotTables.Add(wsPivot7.Cells["A3"], ws.Cells["K1:N11"], "Pivottable7");
			// Commented out because this is not how page fields work.
			//pt.PageFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[0]);
			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = false;

			pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Max;
			Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Max);

			pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP;
			Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP);

			pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.None;
			Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.None);

			pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Default;
			Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Default);

			pt.Fields[0].Sort = eSortType.Descending;
			pt.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;

			pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["A3"], ws.Cells["K1:O11"], "Pivottable8");
			pt.RowFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[4]);
			pt.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days | eDateGroupBy.Quarters, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
			pt.RowHeaderCaption = "År";
			pt.Fields[4].Name = "Dag";
			pt.Fields[5].Name = "Månad";
			pt.Fields[6].Name = "Kvartal";
			pt.GrandTotalCaption = "Totalt";

			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = true;

			pt = wsPivot9.PivotTables.Add(wsPivot9.Cells["A3"], ws.Cells["K1:N11"], "Pivottable9");
			// Commented out because this is not how page fields work.
			//pt.PageFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[3]);
			pt.RowFields[0].AddNumericGrouping(-3.3, 5.5, 4.0);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = false;
			pt.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;

			pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["H3"], ws.Cells["K1:O11"], "Pivottable10");
			pt.RowFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[4]);
			pt.Fields[4].AddDateGrouping(7, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
			pt.RowHeaderCaption = "Veckor";
			pt.GrandTotalCaption = "Totalt";

			pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["A60"], ws.Cells["K1:O11"], "Pivottable11");
			pt.RowFields.Add(pt.Fields["Category"]);
			pt.RowFields.Add(pt.Fields["Item"]);
			pt.RowFields.Add(pt.Fields["Date for grouping"]);

			pt.DataFields.Add(pt.Fields[3]);
			pt.DataFields.Add(pt.Fields[2]);
			pt.DataOnRows = true;

			pt = wsPivot10.PivotTables.Add(wsPivot10.Cells["A1"], ws.Cells["K1:O11"], "Pivottable10");
			pt.ColumnFields.Add(pt.Fields[1]);
			pt.RowFields.Add(pt.Fields[0]);
			pt.RowFields.Add(pt.Fields[3]);
			pt.RowFields.Add(pt.Fields[2]);
			pt.RowFields.Add(pt.Fields[4]);
			pt.DataOnRows = true;
			//wsPivot10.Drawings.AddChart("Pivotchart10", OfficeOpenXml.Drawing.Chart.eChartType.BarStacked3D, pt);

		}

		[Ignore]
		[TestMethod]
		public void ReadPivotTable()
		{
			ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\pivot\pivotforread.xlsx"));

			var pivot1 = pck.Workbook.Worksheets[2].PivotTables[0];

			Assert.AreEqual(pivot1.Fields.Count, 24);
			Assert.AreEqual(pivot1.RowFields.Count, 3);
			Assert.AreEqual(pivot1.DataFields.Count, 7);
			Assert.AreEqual(pivot1.ColumnFields.Count, 0);

			Assert.AreEqual(pivot1.DataFields[1].Name, "Sum of n3");
			Assert.AreEqual(pivot1.Fields[2].Sort, eSortType.Ascending);

			Assert.AreEqual(pivot1.DataOnRows, false);

			var pivot2 = pck.Workbook.Worksheets[2].PivotTables[0];
			var pivot3 = pck.Workbook.Worksheets[3].PivotTables[0];

			var pivot4 = pck.Workbook.Worksheets[4].PivotTables[0];
			var pivot5 = pck.Workbook.Worksheets[5].PivotTables[0];
			pivot5.CacheDefinition.SetSourceRangeAddress(pck.Workbook.Worksheets[1], pck.Workbook.Worksheets[1].Cells["Q1:X300"]);

			var pivot6 = pck.Workbook.Worksheets[6].PivotTables[0];

			pck.Workbook.Worksheets[6].Drawings.AddChart("chart1", OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked3D, pivot6);

			pck.SaveAs(new FileInfo(@"c:\temp\pivot\pivotforread_new.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void CreatePivotMultData()
		{
			FileInfo fi = new FileInfo(@"c:\temp\test.xlsx");
			ExcelPackage pck = new ExcelPackage(fi);

			var ws = pck.Workbook.Worksheets.Add("Data");
			var pv = pck.Workbook.Worksheets.Add("Pivot");

			ws.Cells["A1"].Value = "Data1";
			ws.Cells["B1"].Value = "Data2";

			ws.Cells["A2"].Value = "1";
			ws.Cells["B2"].Value = "2";

			ws.Cells["A3"].Value = "3";
			ws.Cells["B3"].Value = "4";

			ws.Select("A1:B3");

			var pt = pv.PivotTables.Add(pv.SelectedRange, ws.SelectedRange, "Pivot");

			pt.RowFields.Add(pt.Fields["Data2"]);

			var df = pt.DataFields.Add(pt.Fields["Data1"]);
			df.Function = DataFieldFunctions.Count;

			df = pt.DataFields.Add(pt.Fields["Data1"]);
			df.Function = DataFieldFunctions.Sum;

			df = pt.DataFields.Add(pt.Fields["Data1"]);
			df.Function = DataFieldFunctions.StdDev;
			df.Name = "DatA1_2";

			pck.Save();
		}

		//[Ignore]
		[TestMethod]
		public void SetBackground()
		{
			var ws = _pck.Workbook.Worksheets.Add("backimg");

			ws.BackgroundImage.Image = Properties.Resources.Test1;
			ws = _pck.Workbook.Worksheets.Add("backimg2");
			ws.BackgroundImage.SetFromFile(new FileInfo(Path.Combine(_clipartPath, "Vector Drawing.wmf")));
		}

		//[Ignore]
		[TestMethod]
		public void SetHeaderFooterImage()
		{

			var ws = _pck.Workbook.Worksheets.Add("HeaderImage");
			ws.HeaderFooter.OddHeader.CenteredText = "Before ";
			var img = ws.HeaderFooter.OddHeader.InsertPicture(Properties.Resources.Test1, PictureAlignment.Centered);
			img.Title = "Renamed Image";
			//img.GrayScale = true;
			//img.BiLevel = true;
			//img.Gain = .5;
			//img.Gamma = .35;

			Assert.AreEqual(img.Width, 426);
			img.Width /= 4;
			Assert.AreEqual(img.Height, 49.5);
			img.Height /= 4;
			Assert.AreEqual(img.Left, 0);
			Assert.AreEqual(img.Top, 0);
			ws.HeaderFooter.OddHeader.CenteredText += " After";


			//img = ws.HeaderFooter.EvenFooter.InsertPicture(new FileInfo(Path.Combine(_clipartPath,"Vector Drawing.wmf")), PictureAlignment.Left);
			//img.Title = "DiskFile";

			//img = ws.HeaderFooter.FirstHeader.InsertPicture(new FileInfo(Path.Combine(_clipartPath, "Vector Drawing2.WMF")), PictureAlignment.Right);
			//img.Title = "DiskFile2";
			ws.Cells["A1:A400"].Value = 1;

			_pck.Workbook.Worksheets.Copy(ws.Name, "Copied HeaderImage");
		}

		//[Ignore]
		//[TestMethod]
		public void NamedStyles()
		{
			var wsSheet = _pck.Workbook.Worksheets.Add("NamedStyles");

			var firstNamedStyle =
				_pck.Workbook.Styles.CreateNamedStyle("templateFirst");

			var s = firstNamedStyle.Style;

			s.Fill.PatternType = ExcelFillStyle.Solid;
			s.Fill.BackgroundColor.SetColor(Color.LightGreen);
			s.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
			s.VerticalAlignment = ExcelVerticalAlignment.Center;

			var secondNamedStyle = _pck.Workbook.Styles.CreateNamedStyle("first", firstNamedStyle.Style).Style;
			secondNamedStyle.Font.Bold = true;
			secondNamedStyle.Font.SetFromFont(new Font("Arial Black", 8));
			secondNamedStyle.Border.Bottom.Style = ExcelBorderStyle.Medium;
			secondNamedStyle.Border.Left.Style = ExcelBorderStyle.Medium;

			wsSheet.Cells["B2"].Value = "Text Center";
			wsSheet.Cells["B2"].StyleName = "first";
			_pck.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";

			var rowStyle = _pck.Workbook.Styles.CreateNamedStyle("RowStyle", firstNamedStyle.Style).Style;
			rowStyle.Fill.BackgroundColor.SetColor(Color.Pink);
			wsSheet.Cells.StyleName = "templateFirst";
			wsSheet.Cells["C5:H15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
			wsSheet.Cells["C5:H15"].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);

			wsSheet.Cells["30:35"].StyleName = "RowStyle";
			var colStyle = _pck.Workbook.Styles.CreateNamedStyle("columnStyle", firstNamedStyle.Style).Style;
			colStyle.Fill.BackgroundColor.SetColor(Color.CadetBlue);

			wsSheet.Cells["D:E"].StyleName = "ColumnStyle";
		}

		//[Ignore]
		//[TestMethod]
		public void StyleFill()
		{
			var ws = _pck.Workbook.Worksheets.Add("Fills");
			ws.Cells["A1:C3"].Style.Fill.Gradient.Type = ExcelFillGradientType.Linear;
			ws.Cells["A1:C3"].Style.Fill.Gradient.Color1.SetColor(Color.Red);
			ws.Cells["A1:C3"].Style.Fill.Gradient.Color2.SetColor(Color.Blue);

			ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.MediumGray;
			ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
			var r = ws.Cells["A2:A3"];
			r.Style.Fill.Gradient.Type = ExcelFillGradientType.Path;
			r.Style.Fill.Gradient.Left = 0.7;
			r.Style.Fill.Gradient.Right = 0.7;
			r.Style.Fill.Gradient.Top = 0.7;
			r.Style.Fill.Gradient.Bottom = 0.7;

			ws.Cells[4, 1, 4, 360].Style.Fill.Gradient.Type = ExcelFillGradientType.Path;

			for (double col = 1; col < 360; col++)
			{
				r = ws.Cells[4, Convert.ToInt32(col)];
				r.Style.Fill.Gradient.Degree = col;
				r.Style.Fill.Gradient.Left = col / 360;
				r.Style.Fill.Gradient.Right = col / 360;
				r.Style.Fill.Gradient.Top = col / 360;
				r.Style.Fill.Gradient.Bottom = col / 360;
			}
			r = ws.Cells["A5"];
			r.Style.Fill.Gradient.Left = .50;

			ws = _pck.Workbook.Worksheets.Add("FullFills");
			ws.Cells.Style.Fill.Gradient.Left = 0.25;
			ws.Cells["A1"].Value = "test";
			ws.Cells["A1"].RichText.Add("Test rt");
			ws.Cells.AutoFilter = true;
			Assert.AreNotEqual(ws.Cells["A1:D5"].Value, null);
		}

		[Ignore]
		[TestMethod]
		public void BuildInStyles()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Default");
			ws.Cells.Style.Font.Name = "Arial";
			ws.Cells.Style.Font.Size = 15;
			ws.Cells.Style.Border.Bottom.Style = ExcelBorderStyle.MediumDashed;
			var n = pck.Workbook.Styles.NamedStyles[0];
			n.Style.Numberformat.Format = "yyyy";
			n.Style.Font.Name = "Arial";
			n.Style.Font.Size = 15;
			n.Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
			n.Style.Border.Bottom.Color.SetColor(Color.Red);
			n.Style.Fill.PatternType = ExcelFillStyle.Solid;
			n.Style.Fill.BackgroundColor.SetColor(Color.Blue);
			n.Style.Border.Bottom.Color.SetColor(Color.Red);
			n.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			n.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			n.Style.TextRotation = 90;
			ws.Cells["a1:c3"].StyleName = "Normal";
			//  n.CustomBuildin = true;
			pck.SaveAs(new FileInfo(@"c:\temp\style.xlsx"));
		}

		//[Ignore]
		//[TestMethod]
		public void AutoFitColumns()
		{
			var ws = _pck.Workbook.Worksheets.Add("Autofit");
			ws.Cells["A1:H1"].Value = "Auto fit column that is veeery long...";
			ws.Cells["B1"].Style.TextRotation = 30;
			ws.Cells["C1"].Style.TextRotation = 45;
			ws.Cells["D1"].Style.TextRotation = 75;
			ws.Cells["E1"].Style.TextRotation = 90;
			ws.Cells["F1"].Style.TextRotation = 120;
			ws.Cells["G1"].Style.TextRotation = 135;
			ws.Cells["H1"].Style.TextRotation = 180;
			ws.Cells["A1:H1"].AutoFitColumns(0);

			ws.Column(40).AutoFit();
		}

		[TestMethod, Ignore]
		public void Moveissue()
		{
			_pck = new ExcelPackage(new FileInfo(@"C:\temp\bug\FormulaIssue\PreDelete.xlsx"));
			_pck.Workbook.Worksheets[1].DeleteRow(2, 4);
			_pck.SaveAs(new FileInfo(@"c:\temp\move.xlsx"));
		}

		[TestMethod, Ignore]
		public void DelCol()
		{
			_pck = new ExcelPackage(new FileInfo(@"C:\temp\bug\FormulaIssue\PreDeleteCol.xlsx"));
			_pck.Workbook.Worksheets[1].DeleteColumn(5, 1);
			_pck.SaveAs(new FileInfo(@"c:\temp\move.xlsx"));
		}

		[TestMethod, Ignore]
		public void InsCol()
		{
			_pck = new ExcelPackage(new FileInfo(@"C:\temp\bug\FormulaIssue\PreDeleteCol.xlsx"));
			_pck.Workbook.Worksheets[1].InsertColumn(4, 5);
			_pck.SaveAs(new FileInfo(@"c:\temp\move.xlsx"));
		}

		[Ignore]
		[TestMethod]
		public void FileLockedProblem()
		{
			using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\url.xlsx")))
			{
				pck.Workbook.Worksheets[1].DeleteRow(1, 1);
				pck.Save();
				pck.Dispose();
			}
		}
		//[Ignore]
		//[TestMethod]
		public void CopyOverwrite()
		{
			var ws = _pck.Workbook.Worksheets.Add("CopyOverwrite");

			for (int col = 1; col < 15; col++)
			{
				for (int row = 1; row < 30; row++)
				{
					ws.SetValue(row, col, "cell " + ExcelAddress.GetAddress(row, col));
				}
			}
			ws.Cells["A1:P30"].Copy(ws.Cells["B1"]);
		}
		[Ignore]
		[TestMethod]
		public void RunSample0()
		{
			FileInfo newFile = new FileInfo(@"c:\temp\bug\sample0.xlsx");
			using (ExcelPackage package = new ExcelPackage(newFile))
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
				worksheet.InsertColumn(1, 1);

				ExcelColumn entireColumn = worksheet.Column(1);

				var last = worksheet.Column(6);
				last.Style.Fill.PatternType = ExcelFillStyle.Solid;
				last.Style.Fill.BackgroundColor.SetColor(Color.Blue);
				last.ColumnMax = 7;
				worksheet.InsertColumn(7, 1);

				//save our new workbook and we are done!
				package.Save();
			}
		}
		[Ignore]
		[TestMethod]
		public void Deletews()
		{
			FileInfo newFile = new FileInfo(@"c:\temp\bug\worksheet error.xlsx");
			using (ExcelPackage package = new ExcelPackage(newFile))
			{
				var ws1 = package.Workbook.Worksheets.Add("sheet1");
				var ws2 = package.Workbook.Worksheets.Add("sheet2");
				var ws3 = package.Workbook.Worksheets.Add("sheet3");

				package.Workbook.Worksheets.MoveToStart(ws3.Name);
				//save our new workbook and we are done!
				package.Save();
			}
			using (ExcelPackage package = new ExcelPackage(newFile))
			{
				package.Workbook.Worksheets.Delete(1);
				var ws3 = package.Workbook.Worksheets.Add("sheet3");
				package.SaveAs(new FileInfo(@"c:\temp\bug\worksheet error_save.xlsx"));
			}
		}

		[TestMethod]
		public void DeleteWorksheetWithImageWithHyperlink()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var ws1 = package.Workbook.Worksheets.Add("sheet1");
				var ws2 = package.Workbook.Worksheets.Add("sheet2");
				Bitmap image = new Bitmap(2, 2);
				image.SetPixel(1, 1, Color.Black);
				ws1.Drawings.AddPicture("Test", image, new Uri("http://wwww.jetreports.com"));
				package.Save();
				Assert.AreEqual(2, package.Workbook.Worksheets.Count);
				package.Workbook.Worksheets.Delete(ws1);
				Assert.AreEqual(1, package.Workbook.Worksheets.Count);
			}
		}

		[TestMethod, Ignore]
		public void Issue15207()
		{
			using (ExcelPackage ep = new ExcelPackage(new FileInfo(@"c:\temp\bug\worksheet error.xlsx")))
			{
				ExcelWorkbook wb = ep.Workbook;

				if (wb != null)
				{
					ExcelWorksheet ws = null;

					ws = wb.Worksheets[1];

					if (ws != null)
					{
						//do something with the worksheet
						ws.Dispose();
					}

					wb.Dispose();

				} //if wb != null

				wb = null;

				//do some other things

				//running through this next line now throws the null reference exception
				//so the inbuilt dispose method doesn't work properly.
			} //using (ExcelPackage ep = new ExcelPackage(new FileInfo(some_file))
		}
		#endregion

		#region Cross-Sheet Reference Update Tests
		[TestMethod]
		public void InsertRowsUpdatesReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells[2, 2].Formula = "C3";
				sheet1.Cells[3, 3].Value = "Hello, world!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
				sheet1.InsertRow(3, 10);
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[13, 3].Value);
				Assert.AreEqual("C13", sheet1.Cells[2, 2].Formula);
				Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void CrossSheetInsertRowsUpdatesReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells[2, 2].Formula = "Sheet2!C3";
				sheet2.Cells[3, 3].Value = "Hello, world!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
				sheet2.InsertRow(3, 10);
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet2.Cells[13, 3].Value);
				Assert.AreEqual("'Sheet2'!C13", sheet1.Cells[2, 2].Formula, true);
				Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void CrossSheetInsertColumnsUpdatesReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells[2, 2].Formula = "'Sheet2'!C3";
				sheet2.Cells[3, 3].Value = "Hello, world!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
				sheet2.InsertColumn(3, 10);
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet2.Cells[3, 13].Value);
				Assert.AreEqual("'Sheet2'!M3", sheet1.Cells[2, 2].Formula);
				Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void CrossSheetInsertRowAfterReferencesHasNoEffect()
		{
			FileInfo file = new FileInfo("report.xlsx");
			using (ExcelPackage package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets.Add("New Sheet");
				var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
				sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
				otherSheet.Cells[3, 3].Formula = "45";
				otherSheet.InsertRow(5, 1);
				Assert.AreEqual("'Other Sheet'!C3", sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void CrossSheetInsertColumnAfterReferencesHasNoEffect()
		{
			FileInfo file = new FileInfo("report.xlsx");
			using (ExcelPackage package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets.Add("New Sheet");
				var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
				sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
				otherSheet.Cells[3, 3].Formula = "45";
				otherSheet.InsertColumn(5, 1);
				Assert.AreEqual("'Other Sheet'!C3", sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void CrossSheetReferenceIsUpdatedWhenSheetIsRenamed()
		{
			FileInfo file = new FileInfo("report.xlsx");
			using (ExcelPackage package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets.Add("New Sheet");
				var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
				sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
				otherSheet.Cells[3, 3].Formula = "45";
				otherSheet.Name = "New Name";
				Assert.AreEqual("'New Name'!C3", sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void CopyCellUpdatesRelativeCrossSheetReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells[3, 3].Formula = "Sheet2!C3";
				sheet2.Cells[3, 3].Value = "Hello, world!";
				sheet2.Cells[3, 4].Value = "Hello, WORLD!";
				sheet2.Cells[4, 3].Value = "Goodbye, world!";
				sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
				package.Workbook.Calculate();
				Assert.AreEqual("'Sheet2'!D4", sheet1.Cells[4, 4].Formula, true);
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				Assert.AreEqual("Goodbye, WORLD!", sheet1.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void CopyCellUpdatesAbsoluteCrossSheetReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells[3, 3].Formula = "Sheet2!$C$3";
				sheet2.Cells[3, 3].Value = "Hello, world!";
				sheet2.Cells[3, 4].Value = "Hello, WORLD!";
				sheet2.Cells[4, 3].Value = "Goodbye, world!";
				sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
				package.Workbook.Calculate();
				Assert.AreEqual("'Sheet2'!$C$3", sheet1.Cells[4, 4].Formula, true);
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				Assert.AreEqual("Hello, world!", sheet1.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void CopyCellUpdatesRowAbsoluteCrossSheetReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells[3, 3].Formula = "Sheet2!C$3";
				sheet2.Cells[3, 3].Value = "Hello, world!";
				sheet2.Cells[3, 4].Value = "Hello, WORLD!";
				sheet2.Cells[4, 3].Value = "Goodbye, world!";
				sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
				package.Workbook.Calculate();
				Assert.AreEqual("'Sheet2'!D$3", sheet1.Cells[4, 4].Formula, true);
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				Assert.AreEqual("Hello, WORLD!", sheet1.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void CopyCellUpdatesColumnAbsoluteCrossSheetReferencesCorrectly()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells[3, 3].Formula = "Sheet2!$C3";
				sheet2.Cells[3, 3].Value = "Hello, world!";
				sheet2.Cells[3, 4].Value = "Hello, WORLD!";
				sheet2.Cells[4, 3].Value = "Goodbye, world!";
				sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
				package.Workbook.Calculate();
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
				package.Workbook.Calculate();
				Assert.AreEqual("'Sheet2'!$C4", sheet1.Cells[4, 4].Formula, true);
				Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
				Assert.AreEqual("Goodbye, world!", sheet1.Cells[4, 4].Value);
			}
		}
		#endregion

		#region Date1904 Test Cases
		[TestMethod]
		public void TestDate1904WithoutSetting()
		{
			string file = "test1904.xlsx";
			DateTime dateTest1 = new DateTime(2008, 2, 29);
			DateTime dateTest2 = new DateTime(1950, 11, 30);

			if (File.Exists(file))
				File.Delete(file);

			ExcelPackage pack = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w = pack.Workbook.Worksheets.Add("test");
			w.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
			w.Cells[1, 1].Value = dateTest1;
			w.Cells[2, 1].Value = dateTest2;
			pack.Save();


			ExcelPackage pack2 = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w2 = pack2.Workbook.Worksheets["test"];

			Assert.AreEqual(dateTest1, w2.Cells[1, 1].Value);
			Assert.AreEqual(dateTest2, w2.Cells[2, 1].Value);
		}

		[TestMethod]
		public void TestDate1904WithSetting()
		{
			string file = "test1904.xlsx";
			DateTime dateTest1 = new DateTime(2008, 2, 29);
			DateTime dateTest2 = new DateTime(1950, 11, 30);

			if (File.Exists(file))
				File.Delete(file);

			ExcelPackage pack = new ExcelPackage(new FileInfo(file));
			pack.Workbook.Date1904 = true;

			ExcelWorksheet w = pack.Workbook.Worksheets.Add("test");
			w.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
			w.Cells[1, 1].Value = dateTest1;
			w.Cells[2, 1].Value = dateTest2;
			pack.Save();


			ExcelPackage pack2 = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w2 = pack2.Workbook.Worksheets["test"];

			Assert.AreEqual(dateTest1, w2.Cells[1, 1].Value);
			Assert.AreEqual(dateTest2, w2.Cells[2, 1].Value);
		}

		[TestMethod]
		public void TestDate1904SetAndRemoveSetting()
		{
			string file = "test1904.xlsx";
			DateTime dateTest1 = new DateTime(2008, 2, 29);
			DateTime dateTest2 = new DateTime(1950, 11, 30);

			if (File.Exists(file))
				File.Delete(file);

			ExcelPackage pack = new ExcelPackage(new FileInfo(file));
			pack.Workbook.Date1904 = true;

			ExcelWorksheet w = pack.Workbook.Worksheets.Add("test");
			w.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
			w.Cells[1, 1].Value = dateTest1;
			w.Cells[2, 1].Value = dateTest2;
			pack.Save();


			ExcelPackage pack2 = new ExcelPackage(new FileInfo(file));
			pack2.Workbook.Date1904 = false;
			pack2.Save();


			ExcelPackage pack3 = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w3 = pack3.Workbook.Worksheets["test"];

			Assert.AreEqual(dateTest1.AddDays(365.5 * -4), w3.Cells[1, 1].Value);
			Assert.AreEqual(dateTest2.AddDays(365.5 * -4), w3.Cells[2, 1].Value);
		}

		[TestMethod]
		public void TestDate1904SetAndSetSetting()
		{
			string file = "test1904.xlsx";
			DateTime dateTest1 = new DateTime(2008, 2, 29);
			DateTime dateTest2 = new DateTime(1950, 11, 30);

			if (File.Exists(file))
				File.Delete(file);

			ExcelPackage pack = new ExcelPackage(new FileInfo(file));
			pack.Workbook.Date1904 = true;

			ExcelWorksheet w = pack.Workbook.Worksheets.Add("test");
			w.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
			w.Cells[1, 1].Value = dateTest1;
			w.Cells[2, 1].Value = dateTest2;
			pack.Save();


			ExcelPackage pack2 = new ExcelPackage(new FileInfo(file));
			pack2.Workbook.Date1904 = true;  // Only the cells must be updated when this change, if set the same nothing must change
			pack2.Save();


			ExcelPackage pack3 = new ExcelPackage(new FileInfo(file));
			ExcelWorksheet w3 = pack3.Workbook.Worksheets["test"];

			Assert.AreEqual(dateTest1, w3.Cells[1, 1].Value);
			Assert.AreEqual(dateTest2, w3.Cells[2, 1].Value);
		}
		[TestMethod]
		public void ValueText()
		{
			ExcelPackage pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("TestFormat");
			ws.Cells[1, 1].Value = 25.96;
			ws.Cells[1, 1].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

			var s = ws.Cells[1, 1].Text;
		}
		[TestMethod, Ignore]
		public void SaveToStream()
		{
			var stream = new MemoryStream(File.ReadAllBytes(@"c:\temp\book1.xlsx"));
			var excelPackage = new ExcelPackage(stream);
			excelPackage.Workbook.Worksheets.Add("test");
			excelPackage.Save();
			var s = stream.ToArray();
		}
		[TestMethod, Ignore]
		public void ColumnsTest()
		{
			var excelPackage = new ExcelPackage();
			var ws = excelPackage.Workbook.Worksheets.Add("ColumnTest");


			for (var c = 4; c <= 20; c += 4)
			{
				var col = ws.Column(c);
				col.ColumnMax = c + 3;
			}

			ws.Column(3).Hidden = true;
			ws.Column(6).Hidden = true;
			ws.Column(9).Hidden = true;
			ws.Column(15).Hidden = true;
			ws.Cells["a1:Z1"].Value = "Test";
			ws.Cells["a1:FF33"].AutoFitColumns(0);
			ws.Column(26).ColumnMax = ExcelPackage.MaxColumns;
			excelPackage.SaveAs(new FileInfo(@"c:\temp\autofit.xlsx"));
		}

		[TestMethod]
		public void Comment()
		{
			InitBase();
			var pck = new ExcelPackage();
			var ws1 = pck.Workbook.Worksheets.Add("Comment1");
			ws1.Cells[1, 1].AddComment("Testing", "test1");

			pck.SaveAs(new FileInfo(_worksheetPath + "comment.xlsx"));

			pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
			var ws2 = pck.Workbook.Worksheets[1];
			ws2.Cells[1, 2].AddComment("Testing", "test1");
			pck.Save();
		}

		[TestMethod]
		public void CommentShiftsWithRowInserts()
		{
			InitBase();
			var pck = new ExcelPackage();
			var ws1 = pck.Workbook.Worksheets.Add("Comment1");
			ws1.Cells[3, 3].AddComment("Testing comment 1", "test1");
			ws1.Cells[4, 3].AddComment("Testing comment 2", "test2");
			var fileInfo = new FileInfo(_worksheetPath + "comment.xlsx");
			pck.SaveAs(fileInfo);
			pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
			ws1 = pck.Workbook.Worksheets[1];
			// Ensure the comments were saved in the correct location.
			Assert.AreEqual("Testing comment 1", ws1.Cells[3, 3].Comment.Text);
			Assert.AreEqual("test1", ws1.Cells[3, 3].Comment.Author);
			Assert.AreEqual("Testing comment 2", ws1.Cells[4, 3].Comment.Text);
			Assert.AreEqual("test2", ws1.Cells[4, 3].Comment.Author);
			// Ensure they get shifted.
			ws1.InsertRow(4, 4);
			Assert.AreEqual("Testing comment 1", ws1.Cells[3, 3].Comment.Text);
			Assert.AreEqual("test1", ws1.Cells[3, 3].Comment.Author);
			Assert.AreEqual("Testing comment 2", ws1.Cells[8, 3].Comment.Text);
			Assert.AreEqual("test2", ws1.Cells[8, 3].Comment.Author);
			pck.Save();
			pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
			ws1 = pck.Workbook.Worksheets[1];
			// Ensure the shifted index is preserved.
			Assert.AreEqual("Testing comment 1", ws1.Cells[3, 3].Comment.Text);
			Assert.AreEqual("test1", ws1.Cells[3, 3].Comment.Author);
			Assert.AreEqual("Testing comment 2", ws1.Cells[8, 3].Comment.Text);
			Assert.AreEqual("test2", ws1.Cells[8, 3].Comment.Author);
		}

		[TestMethod]
		public void CommentShiftsWithColumnInserts()
		{
			InitBase();
			var pck = new ExcelPackage();
			var ws1 = pck.Workbook.Worksheets.Add("Comment1");
			ws1.Cells[3, 3].AddComment("Testing comment 1", "test1");
			ws1.Cells[3, 4].AddComment("Testing comment 2", "test2");
			var fileInfo = new FileInfo(_worksheetPath + "comment.xlsx");
			pck.SaveAs(fileInfo);
			pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
			ws1 = pck.Workbook.Worksheets[1];
			// Ensure the comments were saved in the correct location.
			Assert.AreEqual("Testing comment 1", ws1.Cells[3, 3].Comment.Text);
			Assert.AreEqual("test1", ws1.Cells[3, 3].Comment.Author);
			Assert.AreEqual("Testing comment 2", ws1.Cells[3, 4].Comment.Text);
			Assert.AreEqual("test2", ws1.Cells[3, 4].Comment.Author);
			// Ensure they get shifted.
			ws1.InsertColumn(4, 4);
			Assert.AreEqual("Testing comment 1", ws1.Cells[3, 3].Comment.Text);
			Assert.AreEqual("test1", ws1.Cells[3, 3].Comment.Author);
			Assert.AreEqual("Testing comment 2", ws1.Cells[3, 8].Comment.Text);
			Assert.AreEqual("test2", ws1.Cells[3, 8].Comment.Author);
			pck.Save();
			pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
			ws1 = pck.Workbook.Worksheets[1];
			// Ensure the shifted index is preserved.
			Assert.AreEqual("Testing comment 1", ws1.Cells[3, 3].Comment.Text);
			Assert.AreEqual("test1", ws1.Cells[3, 3].Comment.Author);
			Assert.AreEqual("Testing comment 2", ws1.Cells[3, 8].Comment.Text);
			Assert.AreEqual("test2", ws1.Cells[3, 8].Comment.Author);
		}
		#endregion

		#region CopySheet Tests
		[TestMethod]
		public void CopySheetWithSharedFormula()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var workbook = package.Workbook;
				var sheet1 = workbook.Worksheets.Add("Sheet1");
				sheet1.Cells[2, 2, 5, 2].Value = new object[,] { { 1 }, { 2 }, { 3 }, { 4 } };
				// Creates a shared formula.
				sheet1.Cells["D2:D5"].Formula = "SUM(B2:C2)";
				var sheet2 = workbook.Worksheets.Copy(sheet1.Name, "Sheet2");
				sheet1.InsertColumn(3, 1);
				// Inserting a column on sheet1 should modify the shared formula on sheet1, but not sheet2.
				Assert.AreEqual("SUM(B2:D2)", sheet1.Cells["E2"].Formula);
				Assert.AreEqual("SUM(B2:C2)", sheet2.Cells["D2"].Formula);
			}
		}

		[TestMethod]
		public void CopySheetWithSparklineContainingNullFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var originalSparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, originalSparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, originalSparklineGroups.SparklineGroups[0].Sparklines.Count);
					var copiedSheet = package.Workbook.Worksheets.Add("CopiedSheet", worksheet);
					Assert.AreEqual(1, originalSparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, originalSparklineGroups.SparklineGroups[0].Sparklines.Count);
					var copiedSparklineGroups = copiedSheet.SparklineGroups;
					Assert.AreEqual(1, copiedSparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, copiedSparklineGroups.SparklineGroups[0].Sparklines.Count);
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}
		#endregion

		#region InsertRows Tests
		[TestMethod]
		public void InsertRowsUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertRowsUpdatesPieChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Pie) as ExcelPieChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$8", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$8", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertRowsUpdatesBarChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$8", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$8", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertRowsUpdatesBubbleChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[2, 4].Value = 1;
				worksheet.Cells[3, 4].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "$D$2:$D$3");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$2:$D$6", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		public void InsertRowsUpdatesCommaSeparatedSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[2, 4].Value = 1;
				worksheet.Cells[3, 4].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("Sheet1!$C$2:$C$3,Sheet1!$E$4:$G$5", "Sheet1!$B$2:$B$3,Sheet1!$F$5:$H$6", "Sheet1!$D$2:$D$3,Sheet1!$H$7:$K$9");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6,'Sheet1'!$F$8:$H$9", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6,'Sheet1'!$E$7:$G$8", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$2:$D$6,'Sheet1'!$H$10:$K$12", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ComboFromExcel.xlsx")]
		public void InsertRowsUpdatesComboChart()
		{
			var file = new FileInfo("ComboFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.PlotArea.ChartTypes[2].Series[0].Series);
				sheet.InsertRow(22, 20);
				Assert.AreEqual("'Sheet1'!$C$20:$C$62", drawing.Series[1].Series);
				Assert.AreEqual("'Sheet1'!$B$20:$B$62", drawing.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$20:$D$62", drawing.PlotArea.ChartTypes[2].Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertRowsUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));

				worksheet1.InsertRow(10, 10);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$36", address);
				ExcelRangeBase.SplitAddress(range1.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$Z$36", address);
				address = null;
				ExcelRangeBase.SplitAddress(range2.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$Z$26", address);
			}
		}

		[TestMethod]
		public void InsertRowCrossSheetDoesNotChangeChartPosition()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var otherSheet = package.Workbook.Worksheets.Add("Unrelated Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 2;
				chart.To.Column = 10;
				chart.To.Row = 12;
				otherSheet.InsertRow(1, 10);
				otherSheet.InsertColumn(1, 10);
				Assert.AreEqual(4, chart.From.Column);
				Assert.AreEqual(2, chart.From.Row);
				Assert.AreEqual(10, chart.To.Column);
				Assert.AreEqual(12, chart.To.Row);
			}
		}

		[TestMethod]
		public void InsertRowAndColumnsAboveChartUpdatesChartPosition()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 2;
				chart.To.Column = 10;
				chart.To.Row = 12;
				sheet.InsertRow(1, 10);
				sheet.InsertColumn(1, 6);
				Assert.AreEqual(10, chart.From.Column);
				Assert.AreEqual(12, chart.From.Row);
				Assert.AreEqual(16, chart.To.Column);
				Assert.AreEqual(22, chart.To.Row);
			}
		}

		[TestMethod]
		public void InsertRowAndColumnsInsideChartUpdatesChartTo()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 4;
				chart.To.Column = 10;
				chart.To.Row = 10;
				sheet.InsertRow(6, 10);
				sheet.InsertColumn(6, 6);
				Assert.AreEqual(4, chart.From.Column);
				Assert.AreEqual(4, chart.From.Row);
				Assert.AreEqual(16, chart.To.Column);
				Assert.AreEqual(20, chart.To.Row);
			}
		}

		[TestMethod]
		public void InsertRowUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				var validation = sheet.Cells["D5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$B$3";
				validation = sheet.Cells["D6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$B$3";
				validation = sheet2.Cells["B2"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$B$3";

				//expand the range
				sheet.InsertRow(3, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!$B$2:$B$5", validationRange.Formula.ExcelFormula);

				//validate that the implicitly addressed Data Validation range has also expanded
				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$B$2:$B$5", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void InsertRowUpdatesDataValidationAddressesAndFormulas()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var anyValidation = sheet.Cells["E5:E10"].DataValidation.AddAnyDataValidation();
				var customValidation = sheet.Cells["F1:F6"].DataValidation.AddCustomDataValidation();
				customValidation.Formula.ExcelFormula = "=Sheet!$B$2:$B$5";
				var listValidation = sheet.Cells["G1:G10"].DataValidation.AddListDataValidation();
				listValidation.Formula.ExcelFormula = "A1:A8";
				var timeValidation = sheet.Cells["J1:K10"].DataValidation.AddTimeDataValidation();
				timeValidation.Formula.ExcelFormula = "A10:H100";
				timeValidation.Formula2.ExcelFormula = "I1:I99";
				var dateTimeValidation = sheet.Cells["L1:L10"].DataValidation.AddDateTimeDataValidation();
				dateTimeValidation.Formula.ExcelFormula = "A10:H100";
				dateTimeValidation.Formula2.ExcelFormula = "I1:I99";
				var integerValidation = sheet.Cells["D5:D8"].DataValidation.AddIntegerDataValidation();
				integerValidation.Formula.ExcelFormula = "=Sheet!$B$2:$B$4";
				integerValidation.Formula2.ExcelFormula = "=Sheet!$B$1:$B$5";
				var decimalValidation = sheet.Cells["Z1:AA100"].DataValidation.AddDecimalDataValidation();
				decimalValidation.Formula.ExcelFormula = "=Sheet!$B$2:$B$4";
				decimalValidation.Formula2.ExcelFormula = "=Sheet!$B$1:$B$5";

				sheet.InsertRow(3, 2);

				var updatedAnyValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.Any) as ExcelDataValidationAny;
				Assert.AreEqual("E7:E12", updatedAnyValidation.Address.Address);
				var updatedCustomValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.Custom) as ExcelDataValidationCustom;
				Assert.AreEqual("F1:F8", updatedCustomValidation.Address.Address);
				Assert.AreEqual("='Sheet'!$B$2:$B$7", updatedCustomValidation.Formula.ExcelFormula);
				var updatedListValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.List) as ExcelDataValidationList;
				Assert.AreEqual("G1:G12", updatedListValidation.Address.Address);
				Assert.AreEqual("A1:A10", updatedListValidation.Formula.ExcelFormula);
				var updatedTimeValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.Time) as ExcelDataValidationTime;
				Assert.AreEqual("J1:K12", updatedTimeValidation.Address.Address);
				Assert.AreEqual("A12:H102", updatedTimeValidation.Formula.ExcelFormula);
				Assert.AreEqual("I1:I101", updatedTimeValidation.Formula2.ExcelFormula);
				var updatedDateTimeValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.DateTime) as ExcelDataValidationDateTime;
				Assert.AreEqual("L1:L12", updatedDateTimeValidation.Address.Address);
				Assert.AreEqual("A12:H102", updatedDateTimeValidation.Formula.ExcelFormula);
				Assert.AreEqual("I1:I101", updatedDateTimeValidation.Formula2.ExcelFormula);
				var updatedIntegerValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.Whole) as ExcelDataValidationInt;
				Assert.AreEqual("D7:D10", updatedIntegerValidation.Address.Address);
				Assert.AreEqual("='Sheet'!$B$2:$B$6", updatedIntegerValidation.Formula.ExcelFormula);
				Assert.AreEqual("='Sheet'!$B$1:$B$7", updatedIntegerValidation.Formula2.ExcelFormula);
				var updatedDecimalValidation = sheet.DataValidations.Single(v => v.ValidationType.Type == eDataValidationType.Decimal) as ExcelDataValidationDecimal;
				Assert.AreEqual("Z1:AA102", updatedDecimalValidation.Address.Address);
				Assert.AreEqual("='Sheet'!$B$2:$B$6", updatedDecimalValidation.Formula.ExcelFormula);
				Assert.AreEqual("='Sheet'!$B$1:$B$7", updatedDecimalValidation.Formula2.ExcelFormula);
			}
		}

		[TestMethod]
		public void InsertColumnUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["D5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$C$2";
				validation = sheet.Cells["D6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$C$2";

				//expand the range
				sheet.InsertColumn(3, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!$B$2:$E$2", validationRange.Formula.ExcelFormula);

				//validate implicitly addressed Data Validation range has also expanded
				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$B$2:$E$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsRowsUpdatesDataValidationRangeAcrossSheets()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";

				//expand the range
				sheetTarget.InsertColumn(3, 2);
				sheetTarget.InsertRow(3, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$2:$G$7", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsRowsRetainsDataValidationRangeOtherSheet()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";
				validation = sheetValidations.DataValidations.AddListValidation(sheetValidations.Cells["D6"].Address);
				validation.Formula.ExcelFormula = "=$A$1:$B$2";

				//expand the range
				sheetValidations.InsertColumn(3, 2);
				sheetValidations.InsertRow(3, 2);

				//validate that the Data Validation range has not expanded
				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$2:$E$5", validationRange.Formula.ExcelFormula);

				//validate that the implicitly addressed Data Validation range has not expanded
				validationRange = sheetValidations.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$A$1:$B$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void InsertRowUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:E7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.InsertRow(1, 1);

				Assert.AreEqual("'sheet1'!C4:D5", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.InsertRow(1, 1);

				Assert.AreEqual("'sheet1'!C4:D5", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E6:E8", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void InsertRowUpdatesPivotTableSourceRangeHandlesWorksheetDataSources()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I10:J16", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);

				worksheet.InsertRow(1, 1);

				Assert.AreEqual("I11:J17", pivotTable.Address.Address);
				Assert.AreEqual("C4:F7", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void InsertRowBetweenUpdatesPivotTableSourceRangeHandlesWorksheetDataSources()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I10:J16", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);

				worksheet.InsertRow(8, 1);

				Assert.AreEqual("I11:J17", pivotTable.Address.Address);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableWithExternalSource.xlsx")]
		public void InsertRowUpdatesPivotTableSourceRangeHandlesExternalDataSources()
		{
			var file = new FileInfo("PivotTableWithExternalSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I2:L5", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.External, pivotTable.CacheDefinition.CacheSource);

				worksheet.InsertRow(1, 1);

				Assert.AreEqual("I3:L6", pivotTable.Address.Address);
			}
		}

		[TestMethod]
		public void InsertRowWithSparklineContainingNullFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var sparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("C3", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
					worksheet.InsertRow(1, 1);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("C4", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void InsertRowUpdatePivotTableSourceRange()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				worksheet.InsertRow(1, 1);
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				Assert.AreEqual("A2:G9", cacheDefinition.GetSourceRangeAddress().ToString());

				worksheet = package.Workbook.Worksheets["RowItems"];
				worksheet.InsertColumn(1, 1);
				Assert.AreEqual("A2:G9", cacheDefinition.GetSourceRangeAddress().ToString());
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void InsertRowTooManyTotalRowsThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[ExcelPackage.MaxRows - 1, 1].Value = "Close to the end";
				sheet.InsertRow(ExcelPackage.MaxRows - 2, 2);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void InsertRowThatWouldPushFormulasOffTheSheetThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[1, 1].Value = "Not close to the end of the worksheet";
				sheet.Cells[ExcelPackage.MaxRows, 1].Formula = "3 + 4";
				sheet.InsertRow(2, 1);
			}
		}
		#endregion

		#region InsertColumns Tests
		[TestMethod]
		public void InsertColumnsUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_an_xlchart.0", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));

				worksheet2.InsertColumn(10, 10);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$26", address);
				ExcelRangeBase.SplitAddress(range1.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$AJ$26", address);
				address = null;
				ExcelRangeBase.SplitAddress(range2.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$AJ$26", address);
			}
		}

		[TestMethod]
		public void InsertColumnsUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$B$3:$C$3", "$B$2:$C$2", "");
				worksheet.InsertColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$F$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$F$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertColumnsUpdatesBarChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$B$3:$C$3", "$B$2:$C$2", "");
				worksheet.InsertColumn(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$H$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$H$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertColumnsUpdatesPieChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Pie) as ExcelPieChart;
				chart.Series.AddSeries("$B$3:$C$3", "$B$2:$C$2", "");
				worksheet.InsertColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$F$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$F$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void InsertColumnsUpdatesBubbleChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[4, 2].Value = 1;
				worksheet.Cells[4, 3].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("$B$3:$C$3", "$B$2:$C$2", "$B$4:$C$4");
				worksheet.InsertColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$F$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$F$3", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$B$4:$F$4", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		public void InsertColumnUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:E7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.InsertColumn(1, 1);

				Assert.AreEqual("'sheet1'!D3:E4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.InsertColumn(1, 1);

				Assert.AreEqual("'sheet1'!D3:E4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!F5:F7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void InsertColumnUpdatesPivotTableSourceRangeHandlesWorksheetDataSources()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I10:J16", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);

				worksheet.InsertColumn(1, 1);

				Assert.AreEqual("J10:K16", pivotTable.Address.Address);
				Assert.AreEqual("D3:G6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableWithExternalSource.xlsx")]
		public void InsertColumnUpdatesPivotTableSourceRangeHandlesExternalDataSources()
		{
			var file = new FileInfo("PivotTableWithExternalSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I2:L5", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.External, pivotTable.CacheDefinition.CacheSource);

				worksheet.InsertColumn(1, 1);

				Assert.AreEqual("J2:M5", pivotTable.Address.Address);
			}
		}

		[TestMethod]
		public void InsertColumnWithSparklineContainingNullFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var sparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("C3", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
					worksheet.InsertColumn(1, 1);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("D3", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void InsertColumnsTooManyTotalColumnsThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[1, ExcelPackage.MaxColumns - 1].Value = "Close to the end";
				sheet.InsertColumn(ExcelPackage.MaxColumns - 2, 2);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void InsertColumnThatWouldPushFormulasOffTheSheetThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[1, 1].Value = "Not close to the end of the worksheet";
				sheet.Cells[2, ExcelPackage.MaxColumns - 1].Formula = "3 + 4";
				sheet.InsertColumn(2, 2);
			}
		}
		#endregion

		#region DeleteRow Tests
		#region Delete Rows in the middle of a range. 
		[TestMethod]
		public void DeleteRowsUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				worksheet.DeleteRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsUpdatesPieChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Pie) as ExcelPieChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$8", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$8", chart.Series[0].Series);
				worksheet.DeleteRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsUpdatesBarChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$8", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$8", chart.Series[0].Series);
				worksheet.DeleteRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsUpdatesBubbleChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[2, 4].Value = 1;
				worksheet.Cells[3, 4].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "$D$2:$D$3");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$2:$D$6", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
				worksheet.DeleteRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$3", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$2:$D$3", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ComboFromExcel.xlsx")]
		public void DeleteRowsUpdatesComboChart()
		{
			var file = new FileInfo("ComboFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.PlotArea.ChartTypes[2].Series[0].Series);
				sheet.DeleteRow(22, 20);
				Assert.AreEqual("'Sheet1'!$C$20:$C$22", drawing.Series[1].Series);
				Assert.AreEqual("'Sheet1'!$B$20:$B$22", drawing.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$20:$D$22", drawing.PlotArea.ChartTypes[2].Series[0].Series);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\OneCellAnchorChart.xlsx")]
		public void DeleteRowsFromMiddleOfChartIgnoresOneCellAnchor()
		{
			var file = new FileInfo("OneCellAnchorChart.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$H$21:$H$30", drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$I$21:$I$30", drawing.Series[0].Series);
				Assert.AreEqual(4, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(eEditAs.OneCell, drawing.EditAs);
				var originalColumnOffset = drawing.From.ColumnOff;
				var originalRowOffset = drawing.From.RowOff;
				sheet.DeleteRow(8, 10);
				Assert.AreEqual("'Sheet1'!$H$11:$H$20", drawing.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$I$11:$I$20", drawing.Series[0].Series);
				Assert.AreEqual(4, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(originalColumnOffset, drawing.From.ColumnOff);
				Assert.AreEqual(originalRowOffset, drawing.From.RowOff);
				Assert.AreEqual(eEditAs.OneCell, drawing.EditAs);
			}
		}

		[TestMethod]
		public void DeleteRowsUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));

				worksheet1.DeleteRow(10, 10);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$16", address);
				ExcelRangeBase.SplitAddress(range1.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$Z$16", address);
				address = null;
				ExcelRangeBase.SplitAddress(range2.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$Z$26", address);
			}
		}

		[TestMethod]
		public void DeleteRowCrossSheetDoesNotChangeChartPosition()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var otherSheet = package.Workbook.Worksheets.Add("Unrelated Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 2;
				chart.To.Column = 10;
				chart.To.Row = 12;
				otherSheet.DeleteRow(1, 10);
				otherSheet.DeleteColumn(1, 10);
				Assert.AreEqual(4, chart.From.Column);
				Assert.AreEqual(2, chart.From.Row);
				Assert.AreEqual(10, chart.To.Column);
				Assert.AreEqual(12, chart.To.Row);
			}
		}

		[TestMethod]
		public void DeleteRowAndColumnsAboveChartUpdatesChartPosition()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 12;
				chart.To.Column = 10;
				chart.To.Row = 22;
				sheet.DeleteRow(1, 10);
				Assert.AreEqual(2, chart.From.Row);
				Assert.AreEqual(12, chart.To.Row);
			}
		}

		[TestMethod]
		public void DeleteRowAndColumnsInsideChartUpdatesChartTo()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 4;
				chart.To.Column = 10;
				chart.To.Row = 20;
				sheet.DeleteRow(6, 10);
				Assert.AreEqual(4, chart.From.Row);
				Assert.AreEqual(10, chart.To.Row);
			}
		}

		[TestMethod]
		public void DeleteRowUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["D5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$B$6";
				validation = sheet.Cells["D6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$B$6";

				sheet.DeleteRow(3, 2);

				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!$B$2:$B$4", validationRange.Formula.ExcelFormula);

				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$B$2:$B$4", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsUpdatesDataValidationRangeAcrossSheets()
		{
			using (var package = new ExcelPackage())
			{
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";

				sheetTarget.DeleteRow(3, 2);

				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$2:$E$3", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsRetainsDataValidationRangeOtherSheet()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";
				validation = sheetValidations.DataValidations.AddListValidation(sheetValidations.Cells["D6"].Address);
				validation.Formula.ExcelFormula = "=$A$1:$B$2";

				sheetValidations.DeleteRow(3, 2);

				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$2:$E$5", validationRange.Formula.ExcelFormula);

				validationRange = sheetValidations.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$A$1:$B$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteRowUpdatesCrossSheetFunctions()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");

				sheet1.Cells["C3"].Formula = "SUM(sheet2!C3:C5)";

				sheet2.DeleteRow(4, 1);

				Assert.AreEqual("SUM('sheet2'!C3:C4)", sheet1.Cells["C3"].Formula);
			}
		}

		[TestMethod]
		public void DeleteRowUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:E7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.DeleteRow(1, 1);

				Assert.AreEqual("'sheet1'!C2:D3", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.DeleteRow(1, 1);

				Assert.AreEqual("'sheet1'!C2:D3", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E4:E6", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void DeleteRowUpdatesSparklines()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet1.DeleteRow(7);
					Assert.AreEqual(6, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("C11", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:F6", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D7:F7", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D7", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D8", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!E6:E7", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E8", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!F6:F7", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F8", sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(6, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("C11", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:F6", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D7:F7", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D7", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D8", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!E6:E7", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E8", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!F6:F7", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F8", sparklines[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void DeleteRowUpdatesCrossSheetSparklineFormulas()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet2 = package.Workbook.Worksheets["Sheet2"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet2.DeleteRow(1);
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("'Sheet2'!B1:I1", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("'Sheet2'!B1:I1", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void DeleteRowUpdatesPivotTableSourceRangeHandlesWorksheetDataSources()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I10:J16", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);

				worksheet.DeleteRow(1, 1);

				Assert.AreEqual("I9:J15", pivotTable.Address.Address);
				Assert.AreEqual("C2:F5", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void DeleteRowBetweenUpdatesPivotTableSourceRangeHandlesWorksheetDataSources()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I10:J16", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);

				worksheet.DeleteRow(8, 1);

				Assert.AreEqual("I9:J15", pivotTable.Address.Address);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableWithExternalSource.xlsx")]
		public void DeleteRowUpdatesPivotTableSourceRangeHandlesExternalDataSources()
		{
			var file = new FileInfo("PivotTableWithExternalSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I2:L5", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.External, pivotTable.CacheDefinition.CacheSource);

				worksheet.DeleteRow(1, 1);

				Assert.AreEqual("I1:L4", pivotTable.Address.Address);
			}
		}
		#endregion

		#region Delete Rows over the start of a range. 
		[TestMethod]
		public void DeleteRowsAcrossStartOfChartUpdatesChart()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 4;
				chart.To.Column = 10;
				chart.To.Row = 10;
				sheet.DeleteRow(1, 6);
				Assert.AreEqual(1, chart.From.Row);
				Assert.AreEqual(4, chart.To.Row);
			}
		}

		[TestMethod]
		public void DeleteRowsAcrossEntireChartSourceUpdatesChart()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 4;
				chart.To.Column = 10;
				chart.To.Row = 10;
				sheet.DeleteRow(1, 12);
				// In a two-cell-anchor chart, deleting the entire row/column set also deletes the chart.
				Assert.IsFalse(sheet.Drawings.Any(drawing => drawing.Name.Equals("myChart")));
			}
		}

		[TestMethod]
		public void DeleteRowsFromBeginningOfRangeUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				worksheet.DeleteRow(1, 3);
				Assert.AreEqual("'Sheet1'!$B$1:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$1:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsFromBeginningOfRangeUpdatesPieChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Pie) as ExcelPieChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$8", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$8", chart.Series[0].Series);
				worksheet.DeleteRow(1, 5);
				Assert.AreEqual("'Sheet1'!$B$1:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$1:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsFromStartOfRangeUpdatesBarChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$B$8", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$8", chart.Series[0].Series);
				worksheet.DeleteRow(1, 5);
				Assert.AreEqual("'Sheet1'!$B$1:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$1:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsFromStartOfRangeUpdatesBubbleChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[2, 4].Value = 1;
				worksheet.Cells[3, 4].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "$D$2:$D$3");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$2:$D$6", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
				worksheet.DeleteRow(1, 3);
				Assert.AreEqual("'Sheet1'!$B$1:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$1:$C$3", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$1:$D$3", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ComboFromExcel.xlsx")]
		public void DeleteRowsFromStartOfRangeUpdatesComboChart()
		{
			var file = new FileInfo("ComboFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.PlotArea.ChartTypes[2].Series[0].Series);
				sheet.DeleteRow(1, 10);
				Assert.AreEqual("'Sheet1'!$C$10:$C$32", drawing.Series[1].Series);
				Assert.AreEqual("'Sheet1'!$B$10:$B$32", drawing.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$10:$D$32", drawing.PlotArea.ChartTypes[2].Series[0].Series);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\OneCellAnchorChart.xlsx")]
		public void DeleteRowsFromStartOfRangeUpdatesOneCellAnchor()
		{
			var file = new FileInfo("OneCellAnchorChart.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$H$21:$H$30", drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$I$21:$I$30", drawing.Series[0].Series);
				Assert.AreEqual(4, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(eEditAs.OneCell, drawing.EditAs);
				var originalColumnOffset = drawing.From.ColumnOff;
				sheet.DeleteRow(3, 10);
				Assert.AreEqual("'Sheet1'!$H$11:$H$20", drawing.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$I$11:$I$20", drawing.Series[0].Series);
				Assert.AreEqual(3, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(originalColumnOffset, drawing.From.ColumnOff);
				Assert.AreEqual(0, drawing.From.RowOff);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\TwoCellAnchorChart.xlsx")]
		public void DeleteRowsFromStartOfRangeUpdatesTwoCellAnchor()
		{
			var file = new FileInfo("TwoCellAnchorChart.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$H$21:$H$30", drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$I$21:$I$30", drawing.Series[0].Series);
				Assert.AreEqual(4, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(16, drawing.To.Row);
				Assert.AreEqual(8, drawing.To.Column);
				Assert.AreEqual(eEditAs.TwoCell, drawing.EditAs);
				var originalColumnOffset = drawing.From.ColumnOff;
				sheet.DeleteRow(3, 10);
				Assert.AreEqual("'Sheet1'!$H$11:$H$20", drawing.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$I$11:$I$20", drawing.Series[0].Series);
				Assert.AreEqual(3, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(originalColumnOffset, drawing.From.ColumnOff);
				Assert.AreEqual(0, drawing.From.RowOff);
				Assert.AreEqual(6, drawing.To.Row);
				Assert.AreEqual(8, drawing.To.Column);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AbsoluteChart.xlsx")]
		public void DeleteRowsFromStartOfRangeDoesNothingToAbsoluteAnchoredChart()
		{
			var file = new FileInfo("AbsoluteChart.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$H$21:$H$30", drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$I$21:$I$30", drawing.Series[0].Series);
				Assert.AreEqual(4, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(16, drawing.To.Row);
				Assert.AreEqual(8, drawing.To.Column);
				Assert.AreEqual(eEditAs.Absolute, drawing.EditAs);
				var originalColumnOffset = drawing.From.ColumnOff;
				var originalRowOffset = drawing.From.RowOff;
				sheet.DeleteRow(3, 10);
				Assert.AreEqual("'Sheet1'!$H$11:$H$20", drawing.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$I$11:$I$20", drawing.Series[0].Series);
				Assert.AreEqual(4, drawing.From.Row);
				Assert.AreEqual(1, drawing.From.Column);
				Assert.AreEqual(originalColumnOffset, drawing.From.ColumnOff);
				Assert.AreEqual(originalRowOffset, drawing.From.RowOff);
			}
		}

		[TestMethod]
		public void DeleteRowsFromStartOfRangeUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$2:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$9:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$A$4:$Z$26"));

				worksheet1.DeleteRow(1, 10);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$16", address);
				ExcelRangeBase.SplitAddress(range1.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$Z$16", address);
				address = null;
				ExcelRangeBase.SplitAddress(range2.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$4:$Z$26", address);
			}
		}

		[TestMethod]
		public void DeleteRowFromStartOfRangeUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["D5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$B$6";
				validation = sheet.Cells["D6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$B$6";

				sheet.DeleteRow(1, 4);

				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!$B$1:$B$2", validationRange.Formula.ExcelFormula);

				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$B$1:$B$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsFromBeginningOfRangeUpdatesDataValidationRangeAcrossSheets()
		{
			using (var package = new ExcelPackage())
			{
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";

				sheetTarget.DeleteRow(1, 3);

				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$1:$E$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteRowFromBeginningOfRangeUpdatesCrossSheetFunctions()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");

				sheet1.Cells["C3"].Formula = "SUM(sheet2!C3:C5)";

				sheet2.DeleteRow(1, 4);

				Assert.AreEqual("SUM('sheet2'!C1)", sheet1.Cells["C3"].Formula);
			}
		}

		[TestMethod]
		public void DeleteRowFromBeginningOfRangeUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:E7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.DeleteRow(1, 1);

				Assert.AreEqual("'sheet1'!C2:D3", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.DeleteRow(1, 5);

				Assert.AreEqual("'sheet1'!C2:D3", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E1:E2", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}
		#endregion

		#region Delete Rows so that the range references are no longer valid.
		[TestMethod]
		public void DeleteRowsAllRowsUpdatesChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				worksheet.DeleteRow(1, 12);
				Assert.AreEqual("'Sheet1'!#REF!", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!#REF!", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteRowsOverAllDataUpdatesBubbleChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[2, 4].Value = 1;
				worksheet.Cells[3, 4].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "$D$2:$D$3");
				worksheet.InsertRow(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$B$6", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$C$2:$C$6", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!$D$2:$D$6", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
				worksheet.DeleteRow(1, 6);
				Assert.AreEqual("'Sheet1'!#REF!", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!#REF!", chart.Series[0].Series);
				Assert.AreEqual("'Sheet1'!#REF!", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		public void DeleteRowsCoveringAllValuesUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$2:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$9:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$A$4:$Z$26"));

				worksheet1.DeleteRow(1, 26);
				Assert.AreEqual("'Sheet1'!#REF!", range0.NameFormula);
				Assert.AreEqual("'Sheet1'!#REF!", range1.NameFormula);
				Assert.AreEqual("Sheet2!$A$4:$Z$26", range2.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowEntireDataValidationSourceRangeUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["D17"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$B$6";
				validation = sheet.Cells["D18"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$B$6";

				sheet.DeleteRow(1, 6);

				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!#REF!", validationRange.Formula.ExcelFormula);

				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=#REF!", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsDeletesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["Z5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$E$2";
				validation = sheet.Cells["Z6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$E$2";
				Assert.AreEqual(2, sheet.DataValidations.Count);
				sheet.DeleteRow(4, 3);
				Assert.AreEqual(0, sheet.DataValidations.Count);
			}
		}

		[TestMethod]
		public void DeleteRowEntireSourceRangeUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:E7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.DeleteRow(1, 1);

				Assert.AreEqual("'sheet1'!C2:D3", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.DeleteRow(1, 10);

				Assert.AreEqual("'sheet1'!C2:D3", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!#REF!", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}
		#endregion

		#region ConditionalFormatting
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingContainedShouldDeleteConditionalFormattingSingleRow()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.AreEqual(2, worksheet.ConditionalFormatting.Count(f => f.Type == eExcelConditionalFormattingRuleType.Expression));
				worksheet.DeleteRow(87, 1);
				Assert.AreEqual(20, worksheet.ConditionalFormatting.Count);
				Assert.AreEqual(1, worksheet.ConditionalFormatting.Count(f => f.Type == eExcelConditionalFormattingRuleType.Expression));
				var nextRowRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Expression);
				Assert.AreEqual("C87", nextRowRule.Address.ToString());
				var previousRowValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.DuplicateValues);
				Assert.AreEqual("C71:C83", previousRowValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingContainedShouldDeleteConditionalFormattingSingleRowMultiples()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.AreEqual(2, worksheet.ConditionalFormatting.Count(f => f.Type == eExcelConditionalFormattingRuleType.Expression));
				worksheet.DeleteRow(87, 2);
				Assert.AreEqual(19, worksheet.ConditionalFormatting.Count);
				Assert.AreEqual(0, worksheet.ConditionalFormatting.Count(f => f.Type == eExcelConditionalFormattingRuleType.Expression));
				var previousRowValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.DuplicateValues);
				Assert.AreEqual("C71:C83", previousRowValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingContainedShouldDeleteConditionalFormatting()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Top));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.TopPercent));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Bottom));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				worksheet.DeleteRow(39, 12);
				Assert.AreEqual(17, worksheet.ConditionalFormatting.Count);
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Top));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.TopPercent));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Bottom));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				var nextRowRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.AboveAverage);
				Assert.AreEqual("C43:C54", nextRowRule.Address.ToString());
				var previousRowValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("C22:C33", previousRowValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingPartiallyContainedShouldNotDeleteConditionalFormatting()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Top));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.TopPercent));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Bottom));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				worksheet.DeleteRow(39, 11);
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Top));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.TopPercent));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Bottom));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				var nextRowRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.AboveAverage);
				Assert.AreEqual("C44:C55", nextRowRule.Address.ToString());
				var previousRowValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("C22:C33", previousRowValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingEndContainedShouldNotDeleteConditionalFormatting()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Top));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.TopPercent));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Bottom));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				worksheet.DeleteRow(40, 12);
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Top));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.TopPercent));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.Bottom));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				var nextRowRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.AboveAverage);
				Assert.AreEqual("C43:C54", nextRowRule.Address.ToString());
				var previousRowValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("C22:C33", previousRowValue.Address.ToString());
			}
		}

		[TestMethod]
		public void DeleteRowsWithCombinedAddressConditionalFormatting()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet 1");
				var address = new ExcelAddress("B2:D5,F6:G7,Z10,Z11");
				var conditionalFormatting = worksheet.ConditionalFormatting.AddContainsBlanks(address);
				Assert.AreEqual("B2:D5,F6:G7,Z10,Z11", conditionalFormatting.Address.ToString());
				var sqrefValue = conditionalFormatting.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("B2:D5 F6:G7 Z10 Z11", sqrefValue);

				worksheet.DeleteRow(6, 5);
				Assert.AreEqual("B2:D5,Z6", conditionalFormatting.Address.ToString());
				sqrefValue = conditionalFormatting.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("B2:D5 Z6", sqrefValue);
			}
		}

		[TestMethod]
		public void DeleteRowsWithCombinedAddressConditionalFormattingMany()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet 1");
				var address1 = new ExcelAddress("B2:D5,F6:G7,Z10,Z11");
				var address2 = new ExcelAddress("F7,G8,J11:K12");
				var conditionalFormatting1 = worksheet.ConditionalFormatting.AddContainsBlanks(address1);
				var conditionalFormatting2 = worksheet.ConditionalFormatting.AddContainsErrors(address2);
				Assert.AreEqual("B2:D5,F6:G7,Z10,Z11", conditionalFormatting1.Address.ToString());
				var sqrefValue = conditionalFormatting1.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("B2:D5 F6:G7 Z10 Z11", sqrefValue);
				Assert.AreEqual("F7,G8,J11:K12", conditionalFormatting2.Address.ToString());
				var sqrefValue2 = conditionalFormatting2.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("F7 G8 J11:K12", sqrefValue2);

				worksheet.DeleteRow(6, 5);
				Assert.AreEqual("B2:D5,Z6", conditionalFormatting1.Address.ToString());
				sqrefValue = conditionalFormatting1.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("B2:D5 Z6", sqrefValue);
				Assert.AreEqual("J6:K7", conditionalFormatting2.Address.ToString());
				sqrefValue2 = conditionalFormatting2.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("J6:K7", sqrefValue2);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithCombinedAddressConditionalFormattingX14()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var rule = worksheet.X14ConditionalFormatting.X14Rules.First(r => r.TopNode.ChildNodes[0].Attributes["type"].Value == "dataBar");
				Assert.AreEqual("G5:G16,S6:S8", rule.Address);
				Assert.AreEqual("G5:G16 S6:S8", rule.GetXmlNodeString("xm:sqref"));

				worksheet.DeleteRow(6, 1);
				Assert.AreEqual("G5:G15,S6:S7", rule.Address);
				Assert.AreEqual("G5:G15 S6:S7", rule.GetXmlNodeString("xm:sqref"));

				worksheet.DeleteRow(6, 2);
				Assert.AreEqual("G5:G13", rule.Address);
				Assert.AreEqual("G5:G13", rule.GetXmlNodeString("xm:sqref"));
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingX14()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(2, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsTrue(worksheet.X14ConditionalFormatting.X14Rules.Any(f => f.Address == "G5:G16,S6:S8"));
				worksheet.DeleteRow(5, 12);
				Assert.AreEqual(1, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsFalse(worksheet.X14ConditionalFormatting.X14Rules.Any(f => f.Address == "G5:G16,S6:S8"));
				var nextRowRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("C10:C21", nextRowRule.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteRowsWithConditionalFormattingX14PartiallyContainedShouldNotDelete()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(2, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsTrue(worksheet.X14ConditionalFormatting.X14Rules.Any(f => f.Address == "G5:G16,S6:S8"));
				worksheet.DeleteRow(5, 11);
				Assert.AreEqual(2, worksheet.X14ConditionalFormatting.X14Rules.Count);
				var transformedRule = worksheet.X14ConditionalFormatting.X14Rules.First(r => r.Address == "G5");
				Assert.IsNotNull(transformedRule);
				var nextRowRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Between);
				Assert.AreEqual("C11:C22", nextRowRule.Address.ToString());
			}
		}
		#endregion

		#region Sparkline
		[TestMethod]
		public void DeleteRowWithSparklineContainingNullFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var sparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("C3", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
					worksheet.DeleteRow(1, 1);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("C2", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}
		#endregion
		#endregion

		#region DeleteColumn Tests
		#region Delete columns in the middle of a range.
		[TestMethod]
		public void DeleteColumnsAboveChartUpdatesChartPosition()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 10;
				chart.From.Row = 2;
				chart.To.Column = 12;
				chart.To.Row = 12;
				sheet.DeleteColumn(1, 6);
				Assert.AreEqual(4, chart.From.Column);
				Assert.AreEqual(6, chart.To.Column);
			}
		}

		[TestMethod]
		public void DeleteColumnsInsideChartUpdatesChart()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 4;
				chart.To.Column = 10;
				chart.To.Row = 10;
				sheet.DeleteColumn(6, 2);
				Assert.AreEqual(4, chart.From.Column);
				Assert.AreEqual(8, chart.To.Column);
			}
		}

		[TestMethod]
		public void DeleteColumnsUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_an_xlchart.0", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));
				worksheet2.DeleteColumn(10, 10);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$26", address);
				ExcelRangeBase.SplitAddress(range1.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$P$26", address);
				address = null;
				ExcelRangeBase.SplitAddress(range2.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$P$26", address);
			}
		}

		[TestMethod]
		public void DeleteColumnsUpdatesNamedRanges()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = worksheet2.Names.Add("myNamedRange", new ExcelRangeBase(worksheet1, "$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("workbookNamedRange", new ExcelRangeBase(worksheet2, "Sheet2!$A$1:$Z$26"));
				worksheet2.DeleteColumn(10, 10);
				Assert.AreEqual("'Sheet1'!$A$1:$Z$26", range0.NameFormula);
				Assert.AreEqual("'Sheet2'!$A$1:$P$26", range1.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$B$3:$C$3", "$B$2:$C$2", "");
				worksheet.InsertColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$F$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$F$3", chart.Series[0].Series);
				worksheet.DeleteColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$C$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteColumnsUpdatesBarChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$B$3:$H$3", "$B$2:$H$2", "");
				worksheet.DeleteColumn(3, 5);
				Assert.AreEqual("'Sheet1'!$B$2:$C$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteColumnsUpdatesPieChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Pie) as ExcelPieChart;
				chart.Series.AddSeries("$B$3:$F$3", "$B$2:$F$2", "");
				worksheet.DeleteColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$C$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteColumnUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["K5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$E$2";
				validation = sheet.Cells["K6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$E$2";

				sheet.DeleteColumn(3, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!$B$2:$C$2", validationRange.Formula.ExcelFormula);

				//validate implicitly addressed Data Validation range has also expanded
				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$B$2:$C$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsUpdatesDataValidationRangeAcrossSheets()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";

				sheetTarget.DeleteColumn(3, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$2:$C$5", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsRetainsDataValidationRangeOtherSheet()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!K5");
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";
				validation = sheetValidations.DataValidations.AddListValidation("K6");
				validation.Formula.ExcelFormula = "=$A$1:$B$2";

				sheetValidations.DeleteColumn(3, 2);

				//validate that the Data Validation range has not expanded
				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$B$2:$E$5", validationRange.Formula.ExcelFormula);

				//validate that the implicitly addressed Data Validation range has not expanded
				validationRange = sheetValidations.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$A$1:$B$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnUpdatesCrossSheetFunctions()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");

				sheet1.Cells["C3"].Formula = "SUM(sheet2!C3:E3)";

				sheet2.DeleteColumn(4, 1);

				Assert.AreEqual("SUM('sheet2'!C3:D3)", sheet1.Cells["C3"].Formula);
			}
		}

		[TestMethod]
		public void DeleteColumnUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:E7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.DeleteColumn(1, 1);

				Assert.AreEqual("'sheet1'!B3:C4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:E7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.DeleteColumn(1, 1);

				Assert.AreEqual("'sheet1'!B3:C4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!D5:D7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void DeleteColumnUpdatesSparklines()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet1.DeleteColumn(5);
					Assert.AreEqual(6, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:E6", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("F6", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D7:E7", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("F7", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D8:E8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("F8", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!E6:E8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(6, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:E6", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("F6", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D7:E7", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("F7", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D8:E8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("F8", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D6:D8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!E6:E8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sparklines[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void DeleteColumnUpdatesCrossSheetSparklineFormulas()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet2 = package.Workbook.Worksheets["Sheet2"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet2.DeleteColumn(4);
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("'Sheet2'!B2:H2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("'Sheet2'!B2:H2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void DeleteColumnUpdatesPivotTableSourceRangeHandlesWorksheetDataSources()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I10:J16", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("C3:F6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);

				worksheet.DeleteColumn(1, 1);

				Assert.AreEqual("H10:I16", pivotTable.Address.Address);
				Assert.AreEqual("B3:E6", pivotTable.CacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableWithExternalSource.xlsx")]
		public void DeleteColumnUpdatesPivotTableSourceRangeHandlesExternalDataSources()
		{
			var file = new FileInfo("PivotTableWithExternalSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				Assert.AreEqual("I2:L5", pivotTable.Address.Address);
				Assert.AreEqual(eSourceType.External, pivotTable.CacheDefinition.CacheSource);

				worksheet.DeleteColumn(1, 1);

				Assert.AreEqual("H2:K5", pivotTable.Address.Address);
			}
		}
		#endregion

		#region Delete columns from the start of a range.
		[TestMethod]
		public void DeleteColumnsAcrossStartOfChartUpdatesChart()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
				var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
				chart.From.Column = 4;
				chart.From.Row = 4;
				chart.To.Column = 10;
				chart.To.Row = 10;
				sheet.DeleteColumn(1, 6);
				Assert.AreEqual(1, chart.From.Column);
				Assert.AreEqual(4, chart.To.Column);
			}
		}

		[TestMethod]
		public void DeleteColumnsFromStartOfChartSeriesUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_an_xlchart.0", new ExcelRangeBase(worksheet2, "Sheet2!$C$1:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$B$1:$Z$26"));

				worksheet2.DeleteColumn(1, 10);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$26", address);
				ExcelRangeBase.SplitAddress(range1.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$P$26", address);
				address = null;
				ExcelRangeBase.SplitAddress(range2.NameFormula, out workbook, out worksheet, out address);
				Assert.AreEqual("$A$1:$P$26", address);
			}
		}

		[TestMethod]
		public void DeleteColumnsFromStartOfSeriesUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$B$3:$C$3", "$B$2:$C$2", "");
				worksheet.InsertColumn(3, 3);
				Assert.AreEqual("'Sheet1'!$B$2:$F$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$B$3:$F$3", chart.Series[0].Series);
				worksheet.DeleteColumn(1, 3);
				Assert.AreEqual("'Sheet1'!$A$2:$C$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$A$3:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteColumnsFromStartOfRangeUpdatesChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$B$3:$H$3", "$B$2:$H$2", "");
				worksheet.DeleteColumn(1, 5);
				Assert.AreEqual("'Sheet1'!$A$2:$C$2", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!$A$3:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteColumnsOverStartOfDataValidationRangeUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["D5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$E$2";
				validation = sheet.Cells["D6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$E$2";

				sheet.DeleteColumn(1, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!$A$2:$C$2", validationRange.Formula.ExcelFormula);

				//validate implicitly addressed Data Validation range has also expanded
				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=$A$2:$C$2", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsFromStartOfRangeUpdatesDataValidationRangeAcrossSheets()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";

				sheetTarget.DeleteColumn(1, 2);

				//validate that the Data Validation range has also expanded
				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!$A$2:$C$5", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnAcrossStartOfRangeUpdatesCrossSheetFunctions()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");

				sheet1.Cells["C3"].Formula = "SUM(sheet2!C3:E3)";

				sheet2.DeleteColumn(1, 3);

				Assert.AreEqual("SUM('sheet2'!A3:B3)", sheet1.Cells["C3"].Formula);
			}
		}

		[TestMethod]
		public void DeleteColumnAcrossPivotTableSourceUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:G7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:G7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.DeleteColumn(1, 1);

				Assert.AreEqual("'sheet1'!B3:C4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:G7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.DeleteColumn(1, 5);

				Assert.AreEqual("'sheet1'!B3:C4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!A5:B7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}
		#endregion

		#region Delete columns across an entire source range.
		[TestMethod]
		public void DeleteColumnsAcrossEntireChartSourceUpdatesChart()
		{
			var temp = new FileInfo(Path.GetTempFileName());
			temp.Delete();
			try
			{
				using (var package = new ExcelPackage(temp))
				{
					var sheet = package.Workbook.Worksheets.Add("Chart Sheet");
					var chart = sheet.Drawings.AddChart("myChart", eChartType.BarClustered);
					chart.From.Column = 4;
					chart.From.Row = 4;
					chart.To.Column = 10;
					chart.To.Row = 10;
					package.Save();
				}
				using (var package = new ExcelPackage(temp))
				{
					var sheet = package.Workbook.Worksheets["Chart Sheet"];
					Assert.AreEqual(1, sheet.Drawings.Count);
					sheet.DeleteColumn(1, 12);
					// In a two-cell-anchor chart, deleting the entire row/column set also deletes the chart.
					Assert.AreEqual(0, sheet.Drawings.Count);
					package.Save();
				}
				using (var package = new ExcelPackage(temp))
				{
					var sheet = package.Workbook.Worksheets["Chart Sheet"];
					Assert.AreEqual(0, sheet.Drawings.Count);
				}
			}
			finally
			{
				temp.Delete();
			}
		}

		[TestMethod]
		public void DeleteColumnsOverEntireChartSeriesUpdatesExcel2016ChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// Excel 2016 chart series must include a worksheet name and are stored as named ranges of the form "_xlchart.n", where n is a positive integer. 
				var range0 = package.Workbook.Names.Add("_xlchart.0", new ExcelRangeBase(worksheet1, "Sheet1!$A$1:$Z$26"));
				var range1 = package.Workbook.Names.Add("not_an_xlchart.0", new ExcelRangeBase(worksheet2, "Sheet2!$C$1:$Z$26"));
				var range2 = package.Workbook.Names.Add("_xlchart.2", new ExcelRangeBase(worksheet2, "Sheet2!$B$1:$Z$26"));

				worksheet2.DeleteColumn(1, 26);
				ExcelRangeBase.SplitAddress(range0.NameFormula, out string workbook, out string worksheet, out string address);
				Assert.AreEqual("$A$1:$Z$26", address);
				Assert.AreEqual("'Sheet2'!#REF!", range1.NameFormula);
				Assert.AreEqual("'Sheet2'!#REF!", range2.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsOverAllDataUpdatesChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[2, 3].Value = "Trucks";
				worksheet.Cells[3, 2].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$B$3:$H$3", "$B$2:$H$2", "");
				worksheet.DeleteColumn(1, 10);
				Assert.AreEqual("'Sheet1'!#REF!", chart.Series[0].XSeries);
				Assert.AreEqual("'Sheet1'!#REF!", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void DeleteColumnsOverEntireDataValidationRangeUpdatesDataValidationRange()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var validation = sheet.Cells["Z5"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=Sheet!$B$2:$E$2";
				validation = sheet.Cells["Z6"].DataValidation.AddListDataValidation();
				validation.Formula.ExcelFormula = "=$B$2:$E$2";

				sheet.DeleteColumn(1, 12);

				//validate that the Data Validation range has also expanded
				var validationRange = sheet.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='Sheet'!#REF!", validationRange.Formula.ExcelFormula);

				//validate implicitly addressed Data Validation range has also expanded
				validationRange = sheet.DataValidations.Last() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("=#REF!", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsDeletesDataValidationRange()
		{
			var temp = new FileInfo(Path.GetTempFileName());
			temp.Delete();
			try
			{
				using (var package = new ExcelPackage(temp))
				{
					//make a Data Validation range
					var sheet = package.Workbook.Worksheets.Add("Sheet");
					var validation = sheet.Cells["Z5"].DataValidation.AddListDataValidation();
					validation.Formula.ExcelFormula = "=Sheet!$B$2:$E$2";
					validation = sheet.Cells["Z6"].DataValidation.AddListDataValidation();
					validation.Formula.ExcelFormula = "=$B$2:$E$2";
					Assert.AreEqual(2, sheet.DataValidations.Count);
					package.Save();
				}
				using (var package = new ExcelPackage(temp))
				{
					var sheet = package.Workbook.Worksheets.First();
					Assert.AreEqual(2, sheet.DataValidations.Count);
					sheet.DeleteColumn(20, 12);
					Assert.AreEqual(0, sheet.DataValidations.Count);
					package.Save();
				}
				using (var package = new ExcelPackage(temp))
				{
					var sheet = package.Workbook.Worksheets.First();
					Assert.AreEqual(0, sheet.DataValidations.Count);
				}
			}
			finally
			{
				temp.Delete();
			}
		}

		[TestMethod]
		public void DeleteColumnsEntireValidationRangePoundRefsDataValidationRangeAcrossSheets()
		{
			using (var package = new ExcelPackage())
			{
				//make a Data Validation range
				var sheetTarget = package.Workbook.Worksheets.Add("Sheet");
				var sheetValidations = package.Workbook.Worksheets.Add("Data Validation");

				var validation = sheetValidations.DataValidations.AddListValidation(@"'Sheet'!" + sheetTarget.Cells["D5"].Address);
				validation.Formula.ExcelFormula = "='SHEET'!$B$2:$E$5";

				sheetTarget.DeleteColumn(1, 20);

				var validationRange = sheetValidations.DataValidations.First() as OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList;
				Assert.AreEqual("='SHEET'!#REF!", validationRange.Formula.ExcelFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnAcrossEntireCrossSheetRangePoundRefsCrossSheetFunctions()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");

				sheet1.Cells["C3"].Formula = "SUM(sheet2!C3:E3)";

				sheet2.DeleteColumn(1, 23);

				Assert.AreEqual("SUM('sheet2'!#REF!)", sheet1.Cells["C3"].Formula);
			}
		}

		[TestMethod]
		public void DeleteColumnAcrossEntirePivotTableSourceUpdatesPivotTableSourceRangeCrossSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("sheet2");
				var pivotTable = sheet1.PivotTables.Add(sheet1.Cells["C3:D4"], sheet2.Cells["E5:G7"], "PivotTable");
				Assert.AreEqual("'sheet1'!C3:D4", pivotTable.Address.FullAddress);
				Assert.AreEqual(eSourceType.Worksheet, pivotTable.CacheDefinition.CacheSource);
				Assert.AreEqual("'sheet2'!E5:G7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet1.DeleteColumn(1, 1);

				Assert.AreEqual("'sheet1'!B3:C4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!E5:G7", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);

				sheet2.DeleteColumn(1, 16);

				Assert.AreEqual("'sheet1'!B3:C4", pivotTable.Address.FullAddress);
				Assert.AreEqual("'sheet2'!#REF!", pivotTable.CacheDefinition.GetSourceRangeAddress().FullAddress);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void DeleteColumnAcrossSparklineSourceUpdatesSparklines()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet = package.Workbook.Worksheets.First();
					var sparklines = sheet.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet.DeleteColumn(4, 3);
					Assert.AreEqual(4, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("'Sheet1'!#REF!", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("'Sheet1'!#REF!", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("'Sheet1'!#REF!", sparklines[0].Sparklines[0].Formula.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet = package.Workbook.Worksheets.First();
					var sparklines = sheet.SparklineGroups.SparklineGroups;
					Assert.AreEqual(4, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("'Sheet1'!#REF!", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("'Sheet1'!#REF!", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("'Sheet1'!#REF!", sparklines[0].Sparklines[0].Formula.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}
		#endregion

		#region ConditionalFormatting
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteColumnsWithConditionalFormattingContainedShouldDeleteConditionalFormatting()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.FiveIconSet));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.ContainsBlanks));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				worksheet.DeleteColumn(9, 1);
				Assert.AreEqual(18, worksheet.ConditionalFormatting.Count);
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.FiveIconSet));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.ContainsBlanks));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				var nextColumnRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.NotContainsBlanks);
				Assert.AreEqual("J22:J33", nextColumnRule.Address.ToString());
				var previousColumnValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Bottom);
				Assert.AreEqual("G39:G50", previousColumnValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteColumnsWithConditionalFormattingContainedShouldDeleteConditionalFormattingMultipleColumns()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.FiveIconSet));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.ContainsBlanks));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				Assert.IsTrue(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.NotContainsBlanks));
				worksheet.DeleteColumn(9, 3);
				Assert.AreEqual(17, worksheet.ConditionalFormatting.Count);
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.FiveIconSet));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.ContainsBlanks));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.BottomPercent));
				Assert.IsFalse(worksheet.ConditionalFormatting.Any(f => f.Type == eExcelConditionalFormattingRuleType.NotContainsBlanks));
				var nextColumnRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.ContainsErrors);
				Assert.AreEqual("J22:J33", nextColumnRule.Address.ToString());
				var previousColumnValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Bottom);
				Assert.AreEqual("G39:G50", previousColumnValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteColumnsWithConditionalFormattingNotContainedShouldNotDelete()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				var original = worksheet.ConditionalFormatting;
				worksheet.DeleteColumn(10, 1);
				Assert.AreEqual(21, worksheet.ConditionalFormatting.Count);
				Assert.AreEqual(original, worksheet.ConditionalFormatting);
				var nextColumnRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.NotContainsBlanks);
				Assert.AreEqual("J22:J33", nextColumnRule.Address.ToString());
				var previousColumnValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.Bottom);
				Assert.AreEqual("G39:G50", previousColumnValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteColumnsWithConditionalFormattingX14()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(2, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsTrue(worksheet.X14ConditionalFormatting.X14Rules.Any(f => f.Address == "E22:E33"));
				worksheet.DeleteColumn(5, 1);
				Assert.AreEqual(1, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsFalse(worksheet.X14ConditionalFormatting.X14Rules.Any(f => f.Address == "E22:E33"));
				var nextColumnRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.ContainsBlanks);
				Assert.AreEqual("H22:H33", nextColumnRule.Address.ToString());
				var previousColumnValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.TwoColorScale);
				Assert.AreEqual("C5:C16", previousColumnValue.Address.ToString());
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteColumnsWithConditionalFormattingX14Multiples()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				Assert.AreEqual(2, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsTrue(worksheet.X14ConditionalFormatting.X14Rules.Any(f => f.Address == "E22:E33"));
				worksheet.DeleteColumn(19, 1);
				worksheet.DeleteColumn(5, 3);
				Assert.AreEqual(0, worksheet.X14ConditionalFormatting.X14Rules.Count);
				Assert.IsFalse(worksheet.X14ConditionalFormatting.X14Rules.Any());
				var nextColumnRule = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.ContainsBlanks);
				Assert.AreEqual("F22:F33", nextColumnRule.Address.ToString());
				var previousColumnValue = worksheet.ConditionalFormatting.First(f => f.Type == eExcelConditionalFormattingRuleType.TwoColorScale);
				Assert.AreEqual("C5:C16", previousColumnValue.Address.ToString());
			}
		}

		[TestMethod]
		public void DeleteColumnsWithCombinedAddressConditionalFormatting()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet 1");
				var address = new ExcelAddress("B2:D5,F6:G7,Z10,Z11");
				var conditionalFormatting = worksheet.ConditionalFormatting.AddContainsBlanks(address);
				Assert.AreEqual("B2:D5,F6:G7,Z10,Z11", conditionalFormatting.Address.ToString());
				var sqrefValue = conditionalFormatting.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("B2:D5 F6:G7 Z10 Z11", sqrefValue);

				worksheet.DeleteColumn(6, 5);
				Assert.AreEqual("B2:D5,U10,U11", conditionalFormatting.Address.ToString());
				sqrefValue = conditionalFormatting.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				Assert.AreEqual("B2:D5 U10 U11", sqrefValue);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\AllConditionalFormatting.xlsx")]
		public void DeleteColumnsWithCombinedAddressConditionalFormattingX14()
		{
			var file = new FileInfo(@"AllConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var rule = worksheet.X14ConditionalFormatting.X14Rules.First(r => r.TopNode.ChildNodes[0].Attributes["type"].Value == "dataBar");
				Assert.AreEqual("G5:G16,S6:S8", rule.Address);
				Assert.AreEqual("G5:G16 S6:S8", rule.GetXmlNodeString("xm:sqref"));

				worksheet.DeleteColumn(6, 1);
				Assert.AreEqual("F5:F16,R6:R8", rule.Address);
				Assert.AreEqual("F5:F16 R6:R8", rule.GetXmlNodeString("xm:sqref"));

				worksheet.DeleteColumn(6, 1);
				Assert.AreEqual("Q6:Q8", rule.Address);
				Assert.AreEqual("Q6:Q8", rule.GetXmlNodeString("xm:sqref"));
			}
		}

		#endregion

		#region Sparkline
		[TestMethod]
		public void DeleteColumnWithSparklineContainingNullFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var sparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("C3", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
					worksheet.DeleteColumn(1, 1);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual("B3", sparklineGroups.SparklineGroups[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}
		#endregion
		#endregion

		#region Delete Worksheet Tests
		[TestMethod]
		public void DeleteWorksheetUpdatesNamedRangesTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
				excelPackage.Workbook.Names.Add("wbName1", "Sheet1!C5");
				excelPackage.Workbook.Names.Add("wbName2", "Sheet1!C5,Sheet2!B2");
				excelPackage.Workbook.Names.Add("wbName3", "#N/A");
				excelPackage.Workbook.Names.Add("wbName4", "#REF!");
				sheet2.Names.Add("s1Name1", "Sheet1!C5");
				sheet2.Names.Add("s1Name2", "Sheet1!C5,Sheet2!B2");
				sheet2.Names.Add("s1Name3", "#N/A");
				sheet2.Names.Add("s1Name4", "#REF!");
				excelPackage.Workbook.Worksheets.Delete("Sheet1");
				Assert.AreEqual("#REF!C5", excelPackage.Workbook.Names["wbName1"].NameFormula);
				Assert.AreEqual("#REF!C5,'Sheet2'!B2", excelPackage.Workbook.Names["wbName2"].NameFormula);
				Assert.AreEqual("#N/A", excelPackage.Workbook.Names["wbName3"].NameFormula);
				Assert.AreEqual("#REF!", excelPackage.Workbook.Names["wbName4"].NameFormula);

				Assert.AreEqual("#REF!C5",sheet2.Names["s1Name1"].NameFormula);
				Assert.AreEqual("#REF!C5,'Sheet2'!B2", sheet2.Names["s1Name2"].NameFormula);
				Assert.AreEqual("#N/A", sheet2.Names["s1Name3"].NameFormula);
				Assert.AreEqual("#REF!", sheet2.Names["s1Name4"].NameFormula);
			}
		}
		#endregion

		#region PivotTable Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\NAV001 - Top Customer Overview Design Mode.xlsx")]
		[DeploymentItem(@"..\..\Workbooks\NAV001 - Top Customer Overview Report Mode.xlsx")]
		public void InsertAndDeleteOperationsWorkAsExpectedOnWorkbooksWithPivotTablesAndNormalTables()
		{
			var files = new[]
			{
				new FileInfo(@"NAV001 - Top Customer Overview Design Mode.xlsx"),
				new FileInfo(@"NAV001 - Top Customer Overview Report Mode.xlsx"),
			};
			foreach (var file in files)
			{
				Assert.IsTrue(file.Exists);
				var temp = new FileInfo(Path.GetTempFileName());
				temp.Delete();
				file.CopyTo(temp.ToString());

				try
				{
					using (var package = new ExcelPackage(temp))
					{
						var sheet = package.Workbook.Worksheets["Report"];
						sheet.DeleteRow(1);
						sheet.DeleteColumn(1);
						sheet.InsertRow(1, 1);
						sheet.InsertColumn(1, 1);
					}
				}
				finally
				{
					temp.Delete();
				}
			}
		}
		#endregion

		#region RenameWorksheet Tests
		[TestMethod]
		public void RenameWorksheetUpdatesScatterChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.XYScatter) as ExcelScatterChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.Name = "A new name";
				Assert.AreEqual("'A new name'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'A new name'!$C$2:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void RenameWorksheetUpdatesBarChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.BarClustered) as ExcelBarChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "");
				worksheet.Name = "A new name";
				Assert.AreEqual("'A new name'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'A new name'!$C$2:$C$3", chart.Series[0].Series);
			}
		}

		[TestMethod]
		public void RenameWorksheetUpdatesBubbleChartSeries()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "Cars";
				worksheet.Cells[3, 2].Value = "Trucks";
				worksheet.Cells[2, 3].Value = 10;
				worksheet.Cells[3, 3].Value = 4;
				worksheet.Cells[2, 4].Value = 1;
				worksheet.Cells[3, 4].Value = 2;
				var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Bubble) as ExcelBubbleChart;
				chart.Series.AddSeries("$C$2:$C$3", "$B$2:$B$3", "$D$2:$D$3");
				worksheet.Name = "A new name";
				Assert.AreEqual("'A new name'!$B$2:$B$3", chart.Series[0].XSeries);
				Assert.AreEqual("'A new name'!$C$2:$C$3", chart.Series[0].Series);
				Assert.AreEqual("'A new name'!$D$2:$D$3", ((ExcelBubbleChartSerie)chart.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void RenameWorksheetUpdatesSparklines()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			string newSheetName = "a new name";
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet1.Name = newSheetName;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets[newSheetName];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual($"'{newSheetName}'!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void RenameWorksheetUpdatesCrossSheetSparklinesFormula()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			string newSheetName = "a new name";
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet2 = package.Workbook.Worksheets["Sheet2"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet2.Name = newSheetName;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual($"'{newSheetName}'!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual($"'{newSheetName}'!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sparklines[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		public void RenameWorksheetWithSparklineContainingNullFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var sparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					worksheet.Name = "RenamedSheet";
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}

		[TestMethod]
		public void RenameWorksheetUpdatesChartReferences()
		{
			FileInfo file = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.Add("chartsheet");
					var chart = worksheet.Drawings.AddChart("namey", eChartType.Doughnut);
					var serie = chart.Series.Add("'chartsheet'!C2:C10", "'chartsheet'!D2:D20");
					serie.HeaderAddress = new ExcelAddress("'chartsheet'!A7");
					worksheet.Name = "newName";
					Assert.AreEqual("'newName'!C2:C10", serie.Series);
					Assert.AreEqual("'newName'!D2:D20", serie.XSeries);
					Assert.AreEqual("'newName'!A7", serie.HeaderAddress.FullAddress);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["newName"];
					var chart = worksheet.Drawings.First() as ExcelChart;
					var serie = chart.Series[0];
					Assert.AreEqual("'newName'!C2:C10", serie.Series);
					Assert.AreEqual("'newName'!D2:D20", serie.XSeries);
					Assert.AreEqual("'newName'!A7", serie.HeaderAddress.FullAddress);
				}
			}
			finally
			{
				if (file.Exists)
					file.Delete();
			}
		}

		[TestMethod]
		public void RenameWorksheetUpdatesChartCrossSheetReferences()
		{
			FileInfo file = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage(file))
				{
					var chartSheet = package.Workbook.Worksheets.Add("chartsheet");
					var dataSheet = package.Workbook.Worksheets.Add("datasheet");
					var chart = chartSheet.Drawings.AddChart("namey", eChartType.Doughnut);
					var serie = chart.Series.Add("'datasheet'!C2:C10", "'datasheet'!D2:D20");
					serie.HeaderAddress = new ExcelAddress("'datasheet'!A7");
					dataSheet.Name = "newDataSheetName";
					Assert.AreEqual("'newDataSheetName'!C2:C10", serie.Series);
					Assert.AreEqual("'newDataSheetName'!D2:D20", serie.XSeries);
					Assert.AreEqual("'newDataSheetName'!A7", serie.HeaderAddress.FullAddress);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["chartsheet"];
					var chart = worksheet.Drawings.First() as ExcelChart;
					var serie = chart.Series[0];
					Assert.AreEqual("'newDataSheetName'!C2:C10", serie.Series);
					Assert.AreEqual("'newDataSheetName'!D2:D20", serie.XSeries);
					Assert.AreEqual("'newDataSheetName'!A7", serie.HeaderAddress.FullAddress);
				}
			}
			finally
			{
				if (file.Exists)
					file.Delete();
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ExcelWorksheetRenameWithStartApostropheThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Name = "'New Name";
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ExcelWorksheetRenameWithEndApostropheThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Name = "New Name'";
			}
		}
		#endregion

		#region Save Tests
		[TestMethod]
		public void SaveWorksheetDoesNotSetHiddenRowHeightToZero()
		{
			var file = new FileInfo(Path.GetTempFileName());
			if (file.Exists)
				file.Delete();
			try
			{
				using (var p = new ExcelPackage())
				{
					var sheet = p.Workbook.Worksheets.Add("sheet");
					var row = sheet.Row(2);
					row.Hidden = true;
					Assert.AreNotEqual(0, row.Height);
					Assert.AreNotEqual(0, row.CustomHeight);
					p.SaveAs(file);
				}
				using (var p = new ExcelPackage(file))
				{
					var sheet = p.Workbook.Worksheets["sheet"];
					var row = sheet.Row(2);
					Assert.IsTrue(row.Hidden);
					Assert.AreNotEqual(0, row.Height);
					Assert.AreNotEqual(0, row.CustomHeight);
				}
			}
			finally
			{
				if (file.Exists)
					file.Delete();
			}
		}

		[TestMethod]
		public void SaveIncludesOnlyThoseTablesThatAreNotDeleted()
		{
			var file = new FileInfo(Path.GetTempFileName());
			if (file.Exists)
				file.Delete();
			try
			{
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.Add("Table Sheet");
					sheet.Tables.Add(sheet.Cells["B2:J10"], "TopTable");
					sheet.Tables.Add(sheet.Cells["L2:Z10"], "RightTable");
					sheet.Tables.Add(sheet.Cells["B12:Z26"], "BottomTable");
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets["Table Sheet"];
					Assert.AreEqual(3, sheet.Tables.Count);
					sheet.Tables.Delete(1);
					Assert.AreEqual(2, sheet.Tables.Count);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets["Table Sheet"];
					Assert.AreEqual(2, sheet.Tables.Count);
				}
			}
			finally
			{
				if (file.Exists)
					file.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"Workbooks\PivotTableWithReference.xlsx")]
		public void SavePivotTableWithCrossSheetReference()
		{
			var file = new FileInfo("PivotTableWithReference.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				package.Save();
				var pivotTable = package.Workbook.Worksheets.ElementAt(1).PivotTables.First();
				var refAddress = pivotTable.CacheDefinition.GetXmlNodeString(ExcelPivotCacheDefinition.SourceAddressPath);
				Assert.AreEqual($"'Venta diaria'!$I$9:$U$15", refAddress);
			}
		}

		[TestMethod]
		public void AddingAndSavingWorksheetDoesNotCreateVmlDrawings()
		{
			var file = new FileInfo(Path.GetTempFileName());
			file.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
					var sheetCopy = package.Workbook.Worksheets.Add("Sheet2", sheet1);
					package.SaveAs(file);
					Assert.IsFalse(package.Package.TryGetPart(new Uri(@"/xl/drawings/vmlDrawing1.vml", UriKind.Relative), out _));
				}
			}
			finally
			{
				if (file.Exists)
					file.Delete();
			}
		}
		#endregion

		#region AutoFilters Tests
		[TestMethod]
		public void HasAutoFiltersApplied()
		{
			var file = new FileInfo(Path.GetTempFileName());
			file.Delete();
			try
			{
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					sheet.AutoFilterAddress = new ExcelAddress("B2:D4");
					Assert.IsTrue(sheet.HasAutoFilters);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.First();
					Assert.IsTrue(sheet.HasAutoFilters);
					sheet.RemoveAutoFilters();
					Assert.IsFalse(sheet.HasAutoFilters);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.First();
					Assert.IsFalse(sheet.HasAutoFilters);
				}
			}
			finally
			{
				file.Delete();
			}
		}

		[TestMethod]
		public void HasAutoFiltersWithRangeSpecifiedAutoFilter()
		{
			var file = new FileInfo(Path.GetTempFileName());
			file.Delete();
			try
			{
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					sheet.Cells["B2:D4"].AutoFilter = true;
					Assert.IsTrue(sheet.HasAutoFilters);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.First();
					Assert.IsTrue(sheet.HasAutoFilters);
					sheet.RemoveAutoFilters();
					Assert.IsFalse(sheet.HasAutoFilters);
					package.Save();
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets.First();
					Assert.IsFalse(sheet.HasAutoFilters);
				}
			}
			finally
			{
				file.Delete();
			}
		}
		#endregion

		#region Calculate Tests
		[TestMethod]
		public void DateFunctionsWorkWithDifferentCultureDateFormats()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			var us = CultureInfo.CreateSpecificCulture("en-US");
			Thread.CurrentThread.CurrentCulture = us;
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells[2, 2].Value = "1/15/2014";
				ws.Cells[3, 3].Formula = "EOMONTH(C2, 0)";
				ws.Cells[2, 3].Formula = "EDATE(B2, 0)";
				ws.Cells[4, 3].Formula = "B2 + 15";
				ws.Cells[5, 3].Formula = "B2 - 14";
				ws.Calculate();
				Assert.AreEqual(41654.0, ws.Cells[2, 3].Value);
				Assert.AreEqual(41670.0, ws.Cells[3, 3].Value);
				Assert.AreEqual(41669.0, ws.Cells[4, 3].Value);
				Assert.AreEqual(41640.0, ws.Cells[5, 3].Value);
			}
			var gb = CultureInfo.CreateSpecificCulture("en-GB");
			Thread.CurrentThread.CurrentCulture = gb;
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells[2, 2].Value = "15/1/2014";
				ws.Cells[3, 3].Formula = "EOMONTH(C2, 0)";
				ws.Cells[2, 3].Formula = "EDATE(B2, 0)";
				ws.Cells[4, 3].Formula = "B2 + 15";
				ws.Cells[5, 3].Formula = "B2 - 14";
				ws.Calculate();
				Assert.AreEqual(41654.0, ws.Cells[2, 3].Value);
				Assert.AreEqual(41670.0, ws.Cells[3, 3].Value);
				Assert.AreEqual(41669.0, ws.Cells[4, 3].Value);
				Assert.AreEqual(41640.0, ws.Cells[5, 3].Value);
			}
			var de = CultureInfo.CreateSpecificCulture("de-DE");
			Thread.CurrentThread.CurrentCulture = de;
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells[2, 2].Value = "15.1.2014";
				ws.Cells[3, 3].Formula = "EOMONTH(C2, 0)";
				ws.Cells[2, 3].Formula = "EDATE(B2, 0)";
				ws.Cells[4, 3].Formula = "B2 + 15";
				ws.Cells[5, 3].Formula = "B2 - 14";
				ws.Calculate();
				Assert.AreEqual(41654.0, ws.Cells[2, 3].Value);
				Assert.AreEqual(41670.0, ws.Cells[3, 3].Value);
				Assert.AreEqual(41669.0, ws.Cells[4, 3].Value);
				Assert.AreEqual(41640.0, ws.Cells[5, 3].Value);
			}
			Thread.CurrentThread.CurrentCulture = currentCulture;
		}
		#endregion

		#region Named Range Formula Update Tests
		#region Copy/Delete Worksheet Tests
		[TestMethod]
		public void CopyWorksheetWithWorksheetScopedNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B2, Sheet1!$B$2)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B2, Sheet2!$B$2)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B2, Sheet1!$B$2)");
					var sheet1Copy = excelPackage.Workbook.Worksheets.Copy("Sheet1", "Sheet1 copy");
					Assert.AreEqual(0, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(2, sheet1Copy.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(Sheet1!B2, Sheet1!$B$2)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B2, Sheet2!$B$2)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet1!B2, Sheet1!$B$2)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet1 copy'!B2,'Sheet1 copy'!$B$2)", sheet1Copy.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet2'!B2,'Sheet2'!$B$2)", sheet1Copy.Names["name2"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet1Copy = excelPackage.Workbook.Worksheets["Sheet1 copy"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(0, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(2, sheet1Copy.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(Sheet1!B2, Sheet1!$B$2)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B2, Sheet2!$B$2)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet1!B2, Sheet1!$B$2)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet1 copy'!B2,'Sheet1 copy'!$B$2)", sheet1Copy.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet2'!B2,'Sheet2'!$B$2)", sheet1Copy.Names["name2"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void CopyWorksheetWithWorkbookScopedNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					excelPackage.Workbook.Names.Add("name1", "CONCATENATE(Sheet1!B2, Sheet1!$B$2)");
					excelPackage.Workbook.Names.Add("name2", "CONCATENATE(Sheet2!B2, Sheet2!$B$2)");
					var sheet1Copy = excelPackage.Workbook.Worksheets.Copy("Sheet1", "Sheet1 copy");
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(0, sheet1.Names.Count);
					Assert.AreEqual(0, sheet1Copy.Names.Count);
					Assert.AreEqual(0, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(Sheet1!B2, Sheet1!$B$2)", excelPackage.Workbook.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B2, Sheet2!$B$2)", excelPackage.Workbook.Names["name2"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet1Copy = excelPackage.Workbook.Worksheets["Sheet1 copy"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(0, sheet1.Names.Count);
					Assert.AreEqual(0, sheet1Copy.Names.Count);
					Assert.AreEqual(0, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(Sheet1!B2, Sheet1!$B$2)", excelPackage.Workbook.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B2, Sheet2!$B$2)", excelPackage.Workbook.Names["name2"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void DeleteWorksheetWithWorksheetScopedNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B2, Sheet1!$B$2)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B2, Sheet2!$B$2)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B2, Sheet1!$B$2)");
					sheet2.Names.Add("name4", "CONCATENATE(Sheet2!B2, Sheet2!$B$2)");
					excelPackage.Workbook.Worksheets.Delete(sheet1);
					Assert.AreEqual(0, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(2, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(#REF!B2,#REF!$B$2)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet2'!B2,'Sheet2'!$B$2)", sheet2.Names["name4"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(#REF!B2,#REF!$B$2)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet2'!B2,'Sheet2'!$B$2)", sheet2.Names["name4"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void DeleteWorksheetWithWorkbookScopedNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					excelPackage.Workbook.Names.Add("name1", "CONCATENATE(Sheet1!B2, Sheet1!$B$2)");
					excelPackage.Workbook.Names.Add("name2", "CONCATENATE(Sheet2!B2, Sheet2!$B$2)");
					excelPackage.Workbook.Worksheets.Delete(sheet1);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(0, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(#REF!B2,#REF!$B$2)", excelPackage.Workbook.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet2'!B2,'Sheet2'!$B$2)", excelPackage.Workbook.Names["name2"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual(0, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE(#REF!B2,#REF!$B$2)", excelPackage.Workbook.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE('Sheet2'!B2,'Sheet2'!$B$2)", excelPackage.Workbook.Names["name2"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}
		#endregion

		#region Insert Row(s) Tests
		[TestMethod]
		public void InsertRowBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet1.InsertRow(4, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$6)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$6)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$6)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$6)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$6)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$6)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void InsertMultipleRowsBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet1.InsertRow(4, 3);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$8)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$8)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$8)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$8)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$8)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$8)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void InsertRowAfterNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet1.InsertRow(6, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}
		#endregion

		#region Insert Column(s) Tests
		[TestMethod]
		public void InsertColumnBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet1.InsertColumn(2, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$E$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$E$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$E$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$E$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$E$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$E$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void InsertMultipleColumnsBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet1.InsertColumn(2, 3);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$G$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$G$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$G$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$G$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$G$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$G$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void InsertColumnAfterNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet1.InsertColumn(6, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}
		#endregion

		#region Delete Row(s) Tests
		[TestMethod]
		public void DeleteRowBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet1.DeleteRow(3, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$D$4)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$D$4)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$D$4)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$D$4)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$D$4)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$D$4)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void DeleteMultipleRowsBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!F5, Sheet1!$F$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!F5, Sheet2!$F$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!F5, Sheet1!$F$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!F5, Sheet1!$F$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!F5, Sheet2!$F$5)");
					sheet1.DeleteRow(2, 3);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$F$2)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$F$2)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$F$2)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$F$2)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$F$2)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$F$2)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void DeleteRowAfterNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet1.DeleteRow(7, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}
		#endregion

		#region Delete Columnn(s) Tests
		[TestMethod]
		public void DeleteColumnBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!D5, Sheet1!$D$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!D5, Sheet2!$D$5)");
					sheet1.DeleteColumn(2, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$C$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$C$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$C$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$C$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$C$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!D5,'Sheet1'!$C$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!D5,Sheet2!$D$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void DeleteMultipleColumnsBeforeNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!F5, Sheet1!$F$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!F5, Sheet2!$F$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!F5, Sheet1!$F$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!F5, Sheet1!$F$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!F5, Sheet2!$F$5)");
					sheet1.DeleteColumn(2, 3);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$C$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$C$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$C$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$C$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$C$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!F5,'Sheet1'!$C$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!F5,Sheet2!$F$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void DeleteColumnAfterNamedRanges()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var excelPackage = new ExcelPackage())
				{
					var sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
					sheet1.Names.Add("name1", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					sheet1.Names.Add("name2", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet2.Names.Add("name3", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name4", "CONCATENATE(Sheet1!B5, Sheet1!$B$5)");
					excelPackage.Workbook.Names.Add("name5", "CONCATENATE(Sheet2!B5, Sheet2!$B$5)");
					sheet1.InsertColumn(6, 1);
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
					excelPackage.SaveAs(tempFile);
				}
				using (var excelPackage = new ExcelPackage(tempFile))
				{
					var sheet1 = excelPackage.Workbook.Worksheets["Sheet1"];
					var sheet2 = excelPackage.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual(2, sheet1.Names.Count);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet1.Names["name1"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", sheet1.Names["name2"].NameFormula);
					Assert.AreEqual(1, sheet2.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", sheet2.Names["name3"].NameFormula);
					Assert.AreEqual(2, excelPackage.Workbook.Names.Count);
					Assert.AreEqual("CONCATENATE('Sheet1'!B5,'Sheet1'!$B$5)", excelPackage.Workbook.Names["name4"].NameFormula);
					Assert.AreEqual("CONCATENATE(Sheet2!B5,Sheet2!$B$5)", excelPackage.Workbook.Names["name5"].NameFormula);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}
		#endregion
		#endregion
	}
}
