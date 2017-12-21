using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelNamedRangeTest
	{
		#region Address Tests
		[TestMethod]
		public void SettingAddressHandlesMultiAddresses()
		{

			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				var name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
				name.NameFormula = "Sheet1!C3";
				name.NameFormula = "Sheet1!D3";
				
				// TODO: Rewrite this test for new NamedRange changes.
				Assert.Fail("Test needs to be rewritten for new NamedRange changes.");
			
				//Assert.IsNull(name.Addresses);
				//name.Address = "C3:D3,E3:F3";
				//Assert.IsNotNull(name.Addresses);
				//name.Address = "Sheet1!C3";
				//Assert.IsNull(name.Addresses);
			}
		}
		#endregion

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

		#region GetRelativeAddress Tests
		[TestMethod]
		public virtual void TryGetNamedRangeAddress()
		{
			string sheetName1 = "Sheet1", sheetName2 = "Name with spaces";
			using (var package = new ExcelPackage())
			using (var workbook = package.Workbook)
			using (var worksheet1 = workbook.Worksheets.Add(sheetName1))
			using (var worksheet2 = workbook.Worksheets.Add(sheetName2))
			using (var cell1 = worksheet1.Cells["$D3"])
			using (var cell2 = worksheet1.Cells["G$5"])
			using (var cell3 = worksheet2.Cells["E5"])
			using (var cell4 = worksheet2.Cells["$C$8"])
			{
				var names = workbook.Names;
				names.Add("Name1", cell1);
				names.Add("Name2", cell2);
				names.Add("Name3", cell3);
				names.Add("Name4", cell4);

				// Relative references to a named range are relative to cell A1. Offsets that cause the 
				// relative address to exceed the maximum row or column will wrap around.
				// Examples:
				//	$B2 means on column B and down one row from the relative address.
				//	D$5 means on row 5 and right three columns from the relative address.
				//	C3 means right two and down three from the relative address.
				var expected = $"'{sheetName1}'!$D4";
				string actual = string.Join(string.Empty, package.Workbook.Names["Name1"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'{sheetName1}'!H$5";
				actual = string.Join(string.Empty, package.Workbook.Names["Name2"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'{sheetName2}'!F6";
				actual = string.Join(string.Empty, package.Workbook.Names["Name3"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'{sheetName2}'!$C$8";
				actual = string.Join(string.Empty, package.Workbook.Names["Name4"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
			}
		}
		#endregion
	}
}