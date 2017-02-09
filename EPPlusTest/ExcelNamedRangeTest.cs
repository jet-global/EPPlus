using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelNamedRangeTest
	{
		#region ActualSheetID and LocalSheetID Tests
		[TestMethod]
		public void SheetIDOnWorkbookScopedNamedRangeIsConstant()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var name = package.Workbook.Names.Add("MyNamedRange", new ExcelRangeBase(sheet, "$A$1"));
				Assert.AreEqual(-1, name.ActualSheetID);
				Assert.AreEqual(-1, name.LocalSheetID);
				Assert.IsNull(name.LocalSheet);
			}
		}

		[TestMethod]
		public void SheetIDOnWorksheetScopedNamedRange()
		{
			using (var package = new ExcelPackage())
			{
				for (int i = 0; i < 10; i++)
					package.Workbook.Worksheets.Add($"Sheet {i}");
				var sheet = package.Workbook.Worksheets.Add("Test Sheet");
				var name = sheet.Names.Add("MyNamedRange", new ExcelRangeBase(sheet, "$A$1"));
				Assert.AreEqual(sheet.SheetID, name.ActualSheetID);
				Assert.AreEqual(sheet.PositionID - 1, name.LocalSheetID);
				Assert.AreSame(sheet, name.LocalSheet);
			}
		}
		#endregion
	}
}