using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace EPPlusTest.FormulaParsing.IntegrationTests.ExcelDataProviderTests
{
	[TestClass]
	public class ExcelDataProviderIntegrationTests
	{
		#region ExcelCell tests
		private ExcelCell CreateItem(object val, int row)
		{
			return new ExcelCell(val, null, 0, row);
		}
		#endregion

		#region RangeInfo Tests
		#region AllValues Tests
		[TestMethod]
		public void AllValuesSingleAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("sheet1");
				var rangeInfo = new RangeInfo(worksheet, new ExcelAddress("B2"));
				Assert.AreEqual(1, rangeInfo.AllValues().Count());
				Assert.IsTrue(new List<object> { null }.SequenceEqual(rangeInfo.AllValues()));
			}
		}

		[TestMethod]
		public void AllValuesSingleAddressWithValue()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("sheet1");
				worksheet.Cells["B2"].Value = 4;
				var rangeInfo = new RangeInfo(worksheet, new ExcelAddress("B2"));
				Assert.AreEqual(1, rangeInfo.AllValues().Count());
				Assert.IsTrue(new List<object> { 4 }.SequenceEqual(rangeInfo.AllValues()));
			}
		}

		[TestMethod]
		public void AllValuesSingleRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("sheet1");
				worksheet.Cells["C2"].Value = "2";
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["C4"].Value = "hey";
				var rangeInfo = new RangeInfo(worksheet, new ExcelAddress("B2:C4"));
				Assert.AreEqual(6, rangeInfo.AllValues().Count());
				Assert.IsTrue(new List<object> { null, "2", 4, null, null, "hey" }.SequenceEqual(rangeInfo.AllValues()));
			}
		}

		[TestMethod]
		public void AllValuesUnionOfRanges()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("sheet1");
				worksheet.Cells["C2"].Value = "2";
				worksheet.Cells["C3"].Value = null;
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["C4"].Value = "hey";
				worksheet.Cells["A2"].Value = 7;
				var rangeInfo = new RangeInfo(worksheet, new ExcelAddress("B2:C4,A1:A2,Z9"));
				Assert.AreEqual(9, rangeInfo.AllValues().Count());
				Assert.IsTrue(new List<object> { null, "2", 4, null, null, "hey", null, 7, null }.SequenceEqual(rangeInfo.AllValues()));
			}
		}
		#endregion
		#endregion
	}
}
