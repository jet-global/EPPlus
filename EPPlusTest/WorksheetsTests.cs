using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest
{
	[TestClass]
	public class WorksheetsTests
	{
		private ExcelPackage package;
		private ExcelWorkbook workbook;

		[TestInitialize]
		public void TestInitialize()
		{
			package = new ExcelPackage();
			workbook = package.Workbook;
			workbook.Worksheets.Add("NEW1");
		}

		[TestMethod]
		public void ConfirmFileStructure()
		{
			Assert.IsNotNull(package, "Package not created");
			Assert.IsNotNull(workbook, "No workbook found");
		}

		[TestMethod]
		public void ShouldBeAbleToDeleteAndThenAdd()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete(1);
			workbook.Worksheets.Add("NEW3");
		}

		[TestMethod]
		public void DeleteByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW2");
		}

		[TestMethod, ExpectedException(typeof(ArgumentException))]
		public void DeleteByNameWhereWorkSheetDoesNotExist()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW3");
		}

		[TestMethod]
		public void MoveBeforeByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveAfterByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveBeforeByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveAfterByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		#region Add Tests
		[TestMethod]
		public void AddWorksheetUpdatesChartSeriesReferences()
		{
			var package = new ExcelPackage();
			var myWorkbook = package.Workbook;
			var firstSheet = myWorkbook.Worksheets.Add("Sheet1");
			var chart1 = firstSheet.Drawings.AddChart("Chart1", eChartType.BarClustered);
			chart1.Series.Add("Sheet1!$B$1:$B$16", "Sheet1!$A$1:$A$16");

			var secondSheet = myWorkbook.Worksheets.Add("Sheet2", firstSheet);

			Assert.AreEqual(1, firstSheet.Drawings.Count);
			chart1 = firstSheet.Drawings[0] as ExcelBarChart;
			string workbook, worksheet, range;
			var serie1 = chart1.Series[0];
			ExcelAddress.SplitAddress(serie1.Series, out workbook, out worksheet, out range);
			Assert.AreEqual("Sheet1", worksheet);
			Assert.AreEqual("$B$1:$B$16", range);
			ExcelAddress.SplitAddress(serie1.XSeries, out workbook, out worksheet, out range);
			Assert.AreEqual("Sheet1", worksheet);
			Assert.AreEqual("$A$1:$A$16", range);

			Assert.AreEqual(1, secondSheet.Drawings.Count);
			var chart2 = secondSheet.Drawings[0] as ExcelBarChart;
			var serie2 = chart2.Series[0];
			ExcelAddress.SplitAddress(serie2.Series, out workbook, out worksheet, out range);
			Assert.AreEqual("Sheet2", worksheet);
			Assert.AreEqual("$B$1:$B$16", range);
			ExcelAddress.SplitAddress(serie2.XSeries, out workbook, out worksheet, out range);
			Assert.AreEqual("Sheet2", worksheet);
			Assert.AreEqual("$A$1:$A$16", range);
		}

		[TestMethod]
		public void AddWorksheetUpdatesChartSeriesReferencesWithoutXSeries()
		{
			var package = new ExcelPackage();
			var myWorkbook = package.Workbook;
			var firstSheet = myWorkbook.Worksheets.Add("Sheet1");
			var chart1 = firstSheet.Drawings.AddChart("Chart1", eChartType.BarClustered);
			chart1.Series.Add("Sheet1!$B$1:$B$16", string.Empty);
			// Completely delete the ser/cat node; it will be partially created by Series.Add().
			var xSeriesNode = chart1.Series.TopNode.SelectSingleNode("c:ser", chart1.NameSpaceManager);
			Assert.IsNotNull(xSeriesNode);
			xSeriesNode.RemoveChild(xSeriesNode.SelectSingleNode("c:cat", chart1.Series.NameSpaceManager));
			var secondSheet = myWorkbook.Worksheets.Add("Sheet2", firstSheet);

			Assert.AreEqual(1, firstSheet.Drawings.Count);
			chart1 = firstSheet.Drawings[0] as ExcelBarChart;
			string workbook, worksheet, range;
			var serie1 = chart1.Series[0];
			ExcelAddress.SplitAddress(serie1.Series, out workbook, out worksheet, out range);
			Assert.AreEqual("Sheet1", worksheet);
			Assert.AreEqual("$B$1:$B$16", range);
			Assert.AreEqual(string.Empty, serie1.XSeries);

			Assert.AreEqual(1, secondSheet.Drawings.Count);
			var chart2 = secondSheet.Drawings[0] as ExcelBarChart;
			var serie2 = chart2.Series[0];
			ExcelAddress.SplitAddress(serie2.Series, out workbook, out worksheet, out range);
			Assert.AreEqual("Sheet2", worksheet);
			Assert.AreEqual("$B$1:$B$16", range);
			Assert.AreEqual(string.Empty, serie1.XSeries);
		}
		#endregion

		#region Delete Column with Save Tests

		private const string OutputDirectory = @"d:\temp\";

		[TestMethod, Ignore]
		public void DeleteFirstColumnInRangeColumnShouldBeDeleted()
		{
			// Arrange
			ExcelPackage pck = new ExcelPackage();
			using (
				 Stream file =
					  Assembly.GetExecutingAssembly()
							.GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
			{
				pck.Load(file);
			}
			var wsData = pck.Workbook.Worksheets[1];

			// Act
			wsData.DeleteColumn(1);
			pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

			// Assert
			Assert.AreEqual("Title", wsData.Cells["A1"].Text);
			Assert.AreEqual("First Name", wsData.Cells["B1"].Text);
			Assert.AreEqual("Family Name", wsData.Cells["C1"].Text);
		}


		[TestMethod, Ignore]
		public void DeleteLastColumnInRangeColumnShouldBeDeleted()
		{
			// Arrange
			ExcelPackage pck = new ExcelPackage();
			using (
				 Stream file =
					  Assembly.GetExecutingAssembly()
							.GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
			{
				pck.Load(file);
			}
			var wsData = pck.Workbook.Worksheets[1];

			// Act
			wsData.DeleteColumn(4);
			pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

			// Assert
			Assert.AreEqual("Id", wsData.Cells["A1"].Text);
			Assert.AreEqual("Title", wsData.Cells["B1"].Text);
			Assert.AreEqual("First Name", wsData.Cells["C1"].Text);
		}

		[TestMethod, Ignore]
		public void DeleteColumnAfterNormalRangeSheetShouldRemainUnchanged()
		{
			// Arrange
			ExcelPackage pck = new ExcelPackage();
			using (
				 Stream file =
					  Assembly.GetExecutingAssembly()
							.GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
			{
				pck.Load(file);
			}
			var wsData = pck.Workbook.Worksheets[1];

			// Act
			wsData.DeleteColumn(5);
			pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

			// Assert
			Assert.AreEqual("Id", wsData.Cells["A1"].Text);
			Assert.AreEqual("Title", wsData.Cells["B1"].Text);
			Assert.AreEqual("First Name", wsData.Cells["C1"].Text);
			Assert.AreEqual("Family Name", wsData.Cells["D1"].Text);

		}

		[TestMethod, Ignore]
		[ExpectedException(typeof(ArgumentException))]
		public void DeleteColumnBeforeRangeMimitThrowsArgumentException()
		{
			// Arrange
			ExcelPackage pck = new ExcelPackage();
			using (
				 Stream file =
					  Assembly.GetExecutingAssembly()
							.GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
			{
				pck.Load(file);
			}
			var wsData = pck.Workbook.Worksheets[1];

			// Act
			wsData.DeleteColumn(0);

			// Assert
			Assert.Fail();

		}

		[TestMethod, Ignore]
		[ExpectedException(typeof(ArgumentException))]
		public void DeleteColumnAfterRangeLimitThrowsArgumentException()
		{
			// Arrange
			ExcelPackage pck = new ExcelPackage();
			using (
				 Stream file =
					  Assembly.GetExecutingAssembly()
							.GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
			{
				pck.Load(file);
			}
			var wsData = pck.Workbook.Worksheets[1];

			// Act
			wsData.DeleteColumn(16385);

			// Assert
			Assert.Fail();

		}

		[TestMethod, Ignore]
		public void DeleteFirstTwoColumnsFromRangeColumnsShouldBeDeleted()
		{
			// Arrange
			ExcelPackage pck = new ExcelPackage();
			using (
				 Stream file =
					  Assembly.GetExecutingAssembly()
							.GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
			{
				pck.Load(file);
			}
			var wsData = pck.Workbook.Worksheets[1];

			// Act
			wsData.DeleteColumn(1, 2);
			pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

			// Assert
			Assert.AreEqual("First Name", wsData.Cells["A1"].Text);
			Assert.AreEqual("Family Name", wsData.Cells["B1"].Text);

		}
		#endregion

		[TestMethod]
		public void RangeClearMethodShouldNotClearSurroundingCells()
		{
			var wks = workbook.Worksheets.Add("test");
			wks.Cells[2, 2].Value = "something";
			wks.Cells[2, 3].Value = "something";

			wks.Cells[2, 3].Clear();

			Assert.IsNotNull(wks.Cells[2, 2].Value);
			Assert.AreEqual("something", wks.Cells[2, 2].Value);
			Assert.IsNull(wks.Cells[2, 3].Value);
		}

		private static void CompareOrderOfWorksheetsAfterSaving(ExcelPackage editedPackage)
		{
			var packageStream = new MemoryStream();
			editedPackage.SaveAs(packageStream);

			var newPackage = new ExcelPackage(packageStream);
			var positionId = 1;
			foreach (var worksheet in editedPackage.Workbook.Worksheets)
			{
				Assert.AreEqual(worksheet.Name, newPackage.Workbook.Worksheets[positionId].Name, "Worksheets are not in the same order");
				positionId++;
			}
		}
	}
}
