using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	[TestClass]
	public class GetPivotDataTests
	{
		#region Constants
		private const double Delta = .0000001;
		#endregion

		#region TestMethods
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataValidField()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!C3)";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(17141.16, (double)sheet1.Cells[2, 10].Value, GetPivotDataTests.Delta);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataValidFieldWithPartialCollisionCellReference()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!A1:C3)";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(17141.16, (double)sheet1.Cells[2, 10].Value, GetPivotDataTests.Delta);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataValidFieldSalespersonFilter()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!C3, ""Salesperson"",""Gill"")";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(1749.87, (double)sheet1.Cells[2, 10].Value, GetPivotDataTests.Delta);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataMultipleFields()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!C3, ""Salesperson"",""Gill"", ""Item"",""Pencil"")";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(77.4, (double)sheet1.Cells[2, 10].Value, GetPivotDataTests.Delta);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataInvalidSalespersonFilter()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!C3, ""Salesperson"",""Shawnatron"")";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)sheet1.Cells[2, 10].Value).Type);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataIncompleteSalespersonFilter()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!C3, ""Salesperson"")";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)sheet1.Cells[2, 10].Value).Type);
			}
		}


		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataInvalidField()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Scales"", Sheet2!C3)"; // "Scales" is not a field on the table.
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)sheet1.Cells[2, 10].Value).Type);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataInvalidCellReference()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"", Sheet2!A1)";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)sheet1.Cells[2, 10].Value).Type);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataMissingPivotTableLocationErrors()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(""Sales"")";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)sheet1.Cells[2, 10].Value).Type);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\GetPivotDataTestData.xlsx")]
		public void GetPivotDataMissingFieldArgumentLocationErrors()
		{
			var file = new FileInfo(@"GetPivotDataTestData.xlsx");
			using (var package = new ExcelPackage(file))
			{
				var sheet1 = package.Workbook.Worksheets["Sheet1"];
				var sheet2 = package.Workbook.Worksheets["Sheet2"];
				var pivotTable = sheet2.PivotTables.First();
				sheet1.Cells[2, 10].Formula = @"GETPIVOTDATA(, Sheet2!C3)";
				sheet1.Cells[2, 10].Calculate();
				Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)sheet1.Cells[2, 10].Value).Type);
			}
		}
		#endregion
	}
}
