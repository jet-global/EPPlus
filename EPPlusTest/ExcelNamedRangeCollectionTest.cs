using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelNamedRangeCollectionTest
	{
		#region Insert Tests
		[TestMethod]
		public void InsertRowsBeforeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(1, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C6", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsAfterNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				// No sheet name is added because the address was not modified in any way.
				Assert.AreEqual("C3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsInsideNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3:C5"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:C8", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsInsideMaxRowNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C:D"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C:D", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsHandlesNonContiguousNamedRangeAddresses()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3,D3:D5,E5"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!D3:D8,'Sheet'!E8", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsBeforeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(1, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C6", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsAfterNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("Sheet!C3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsInsideNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3:C5"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:C8", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsInsideMaxRowNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C:D"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C:D", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsHandlesNonContiguousNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3,Sheet!D3:D5,Sheet!E5"));
				namedRangeCollection.Insert(4, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!D3:D8,'Sheet'!E8", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsBeforeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(0, 1, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!F3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsAfterNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				// No sheet name is added because the address was not modified in any way.
				Assert.AreEqual("C3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3:E3"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:H3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideMaxColumnNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "2:3"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!2:3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsHandlesNonContiguousNamedRangeAddresses()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3,C4:E4,E5"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!C4:H4,'Sheet'!H5", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsBeforeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(0, 1, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!F3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsAfterNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("Sheet!C3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3:E3"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:H3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideMaxColumnNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!2:3"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!2:3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsHandlesNonContiguousNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3,Sheet!C4:E4,Sheet!E5"));
				namedRangeCollection.Insert(0, 4, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!C4:H4,'Sheet'!H5", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsWithInvalidRangeAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!#REF!" });
				namedRangeCollection.Insert(1, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!#REF!", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsWithInvalidRangeAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!#REF!" });
				namedRangeCollection.Insert(0, 1, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!#REF!", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsWithInvalidRangeAddressAndValidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "Sheet!#REF!" });
				namedRangeCollection.Insert(1, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("Sheet!#REF!", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsWithInvalidRangeAddressAndValidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "Sheet!#REF!" });
				namedRangeCollection.Insert(0, 1, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("Sheet!#REF!", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsWithValidRangeAddressAndInvalidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!C3" });
				namedRangeCollection.Insert(1, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'#REF'!C6", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsWithValidRangeAddressAndInvalidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!C3" });
				namedRangeCollection.Insert(0, 1, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'#REF'!F3", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertRowsWithInvalidMultiAddressList()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "Sheet!C3,#REF!C3,Sheet!#REF!" });
				namedRangeCollection.Insert(1, 0, 3, 0);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C6,'Sheet'!C6,'Sheet'!#REF!", namedRange.Address);
			}
		}

		[TestMethod]
		public void InsertColumnsWithInvalidMultiAddressList()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "Sheet!C3,#REF!C3,Sheet!#REF!" });
				namedRangeCollection.Insert(0, 1, 0, 3);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!F3,'Sheet'!F3,'Sheet'!#REF!", namedRange.Address);
			}
		}
		#endregion
	}
}
