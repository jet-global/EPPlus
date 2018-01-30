using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelNamedRangeCollectionTest
	{
		#region Add Tests
		[TestMethod]
		public void AddWorkbookScopedWithExcelRangeTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				Assert.IsTrue(namedRangeCollection.ContainsKey("namedrange"));
				var namedRange = namedRangeCollection["namedrange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void AddWorksheetScopedWithExcelRangeTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook, sheet);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				Assert.IsTrue(namedRangeCollection.ContainsKey("namedrange"));
				var namedRange = namedRangeCollection["namedrange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
				Assert.AreEqual(0, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void AddWorkbookScopedWithFormulaTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", "Sheet!D$5,offset(Sheet!C3,1,1)");
				Assert.IsTrue(namedRangeCollection.ContainsKey("namedrange"));
				var namedRange = namedRangeCollection["namedrange"];
				Assert.IsFalse(namedRange.IsNameHidden);
				Assert.AreEqual("Sheet!D$5,offset(Sheet!C3,1,1)", namedRange.NameFormula);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void AddWorkbookScopedHiddenRangeWithFormulaTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", "Sheet!D$5,offset(Sheet!C3,1,1)", isHidden: true);
				Assert.IsTrue(namedRangeCollection.ContainsKey("namedrange"));
				var namedRange = namedRangeCollection["namedrange"];
				Assert.AreEqual("Sheet!D$5,offset(Sheet!C3,1,1)", namedRange.NameFormula);
				Assert.IsTrue(namedRange.IsNameHidden);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void AddWorkbookScopedWithCommentsAndFormulaTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", "Sheet!D$5,offset(Sheet!C3,1,1)", comments: "Some comments.");
				Assert.IsTrue(namedRangeCollection.ContainsKey("namedrange"));
				var namedRange = namedRangeCollection["namedrange"];
				Assert.AreEqual("Sheet!D$5,offset(Sheet!C3,1,1)", namedRange.NameFormula);
				Assert.AreEqual("Some comments.", namedRange.NameComment);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void AddWorksheetScopedWithFormulaTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook, sheet);
				namedRangeCollection.Add("NamedRange", "Sheet!D$5,offset(Sheet!C3,1,1)");
				Assert.IsTrue(namedRangeCollection.ContainsKey("namedrange"));
				var namedRange = namedRangeCollection["namedrange"];
				Assert.AreEqual("Sheet!D$5,offset(Sheet!C3,1,1)", namedRange.NameFormula);
				Assert.AreEqual(0, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithNullNameRangeThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add(null, new ExcelRange(sheet, "Sheet!C3"));
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithEmptyNameRangeThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add(string.Empty, new ExcelRange(sheet, "Sheet!C3"));
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithEmptyRangeThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("namedRange", range: null);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithFormulaNullNameThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add(null, "2 + 2");
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithFormulaEmptyNameThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add(string.Empty, "2 + 2");
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithEmptyFormulaThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("namedRange", formula: string.Empty);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddWithNullFormulaThrowsExceptionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("namedRange", formula: null);
			}
		}
		#endregion

		#region Insert Tests
		[TestMethod]
		public void InsertRowsBeforeAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$6", namedRange.NameFormula);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void InsertRowsBeforeRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsAfterAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsAfterRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsInsideAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C$3:C$5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C$3:C$8", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsInsideRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3:C5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				// Relative named ranges are not expanded when rows are inserted inside them.
				Assert.AreEqual("'Sheet'!C3:C5", namedRange.NameFormula);
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
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C:D", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsInsideAbsolutesColumnNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C:$C"));
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C:$C", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsHandlesNonContiguousRelativeNamedRangeAddresses()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3,D3:D5,E5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!D3:D5,'Sheet'!E5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsHandlesNonContiguousAbsoluteNamedRangeAddresses()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C$3,D$3:D$5,E$5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C$3,'Sheet'!D$3:D$8,'Sheet'!E$8", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsBeforeAbsoluteNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C$3"));
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C$6", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsBeforeRelativeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsBeforeRelativeNamedRangeAbsoluteCrossSheetFormulaWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var sheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", "CONCATENATE(Sheet2!$B$2, Sheet2!C3, Sheet2!D$4, Sheet!$B$2, Sheet!C5)");
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("CONCATENATE(Sheet2!$B$2,Sheet2!C3,Sheet2!D$4,'Sheet'!$B$5,'Sheet'!C5)", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsAfterAbsoluteNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C$3"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsAfterRelativeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsInsideAbsoluteNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C$3:$C$5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3:$C$8", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsInsideRelativeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3:C5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:C5", namedRange.NameFormula);
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
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C:D", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsHandlesNonContiguousAbsoluteNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C$3,Sheet!$D$3:$D$5,Sheet!$E$5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3,'Sheet'!$D$3:$D$8,'Sheet'!$E$8", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsHandlesNonContiguousRelativeNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3,Sheet!D3:D5,Sheet!E5"));
				namedRangeCollection.Insert(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!D3:D5,'Sheet'!E5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsBeforeAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$F$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsInCompleteRowAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$3:$3"));
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$3:$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsInCompleteRowRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "3:3"));
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!3:3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsBeforeRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsAfterAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				var originalNamedRange = namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				Assert.AreEqual("'Sheet'!$C$3", originalNamedRange.NameFormula);
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsAfterRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				var originalNamedRange = namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				Assert.AreEqual("'Sheet'!C3", originalNamedRange.NameFormula);
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C3:$E3"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C3:$H3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3:E3"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:E3", namedRange.NameFormula);
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
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!2:3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsHandlesNonContiguousAbsoluteNamedRangeAddresses()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3,$C4:$E4,$E5"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3,'Sheet'!$C4:$H4,'Sheet'!$H5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsHandlesNonContiguousRelativeNamedRangeAddresses()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3,C4:E4,E5"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!C4:E4,'Sheet'!E5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsBeforeAbsoluteNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C$3"));
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$F$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsBeforeRelativeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3"));
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsAfterAbsoluteNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C3"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsAfterRelativeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C$3"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideAbsoluteNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C3:$E3"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C3:$H3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsInsideRelativeNamedRangeWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3:E3"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:E3", namedRange.NameFormula);
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
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!2:3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsHandlesNonContiguousAbsoluteNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C$3,Sheet!$C$4:$E$4,Sheet!$E$5"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3,'Sheet'!$C$4:$H$4,'Sheet'!$H$5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsHandlesNonContiguousRelativeNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3,Sheet!C4:E4,Sheet!E5"));
				namedRangeCollection.Insert(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!C4:E4,'Sheet'!E5", namedRange.NameFormula);
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
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!#REF!", namedRange.NameFormula);
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
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!#REF!", namedRange.NameFormula);
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
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("Sheet!#REF!", namedRange.NameFormula);
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
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("Sheet!#REF!", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsWithValidAbsoluteRangeAddressAndInvalidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!$C$3" });
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsWithValidRelativeRangeAddressAndInvalidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!C3" });
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsWithValidAbsoluteRangeAddressAndInvalidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!$C3" });
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!$C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsWithValidRelativeRangeAddressAndInvalidSheet()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "#REF!C3" });
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("#REF!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertRowsWithInvalidMultiAddressList()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "Sheet!$C$3,#REF!$C$3,Sheet!D6" });
				namedRangeCollection.Insert(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$6,#REF!$C$3,'Sheet'!D6", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void InsertColumnsWithInvalidMultiAddressList()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3") { Address = "Sheet!$C$3,#REF!$C$3,Sheet!D6" });
				namedRangeCollection.Insert(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$F$3,#REF!$C$3,'Sheet'!D6", namedRange.NameFormula);
			}
		}
		#endregion

		#region Delete Tests
		[TestMethod]
		public void DeleteRowsBeforeAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$5"));
				namedRangeCollection.Delete(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$2", namedRange.NameFormula);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}

		[TestMethod]
		public void DeleteRowsBeforeRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C6"));
				namedRangeCollection.Delete(1, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C6", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsAfterAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				namedRangeCollection.Delete(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsAfterRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Delete(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsInsideAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C$3:C$8"));
				namedRangeCollection.Delete(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C$3:C$5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsInsideBeyoundBoundaryAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C$3:C$7"));
				namedRangeCollection.Delete(4, 0, 5, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsInsideRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3:C5"));
				namedRangeCollection.Delete(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:C5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsHandlesNonContiguousAbsoluteNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!$C$3,Sheet!$D$3:$D$8,Sheet!$E$10"));
				namedRangeCollection.Delete(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3,'Sheet'!$D$3:$D$5,'Sheet'!$E$7", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteRowsHandlesNonContiguousRelativeNamedRangeAddressesWithSheetNames()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "Sheet!C3,Sheet!D3:D5,Sheet!E5"));
				namedRangeCollection.Delete(4, 0, 3, 0, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3,'Sheet'!D3:D5,'Sheet'!E5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsBeforeAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$F$3"));
				namedRangeCollection.Delete(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsBeforeRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				namedRangeCollection.Delete(0, 1, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsAfterAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				var originalNamedRange = namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				Assert.AreEqual("'Sheet'!$C$3", originalNamedRange.NameFormula);
				namedRangeCollection.Delete(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsAfterRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				var originalNamedRange = namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3"));
				Assert.AreEqual("'Sheet'!C3", originalNamedRange.NameFormula);
				namedRangeCollection.Delete(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsInsideAbsoluteNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C3:$G3"));
				namedRangeCollection.Delete(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!$C3:$D3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void DeleteColumnsInsideRelativeNamedRange()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "C3:E3"));
				namedRangeCollection.Delete(0, 4, 0, 3, sheet);
				var namedRange = namedRangeCollection["NamedRange"];
				Assert.AreEqual("'Sheet'!C3:E3", namedRange.NameFormula);
			}
		}
		#endregion

		#region Retrieval Tests
		[TestMethod]
		public void NamedRangeCollectionIndexerCaseInsensitiveTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				var namedRangeCollection = new ExcelNamedRangeCollection(excelPackage.Workbook);
				namedRangeCollection.Add("NamedRange", new ExcelRangeBase(sheet, "$C$3"));
				Assert.IsTrue(namedRangeCollection.ContainsKey("NAMEDrange"));
				var namedRange = namedRangeCollection["namedRANGE"];
				Assert.AreEqual("'Sheet'!$C$3", namedRange.NameFormula);
				Assert.AreEqual(-1, namedRange.LocalSheetID);
			}
		}
		#endregion

		#region Reference Resolution Tests
		[TestMethod]
		public void AbsoluteNamedRangeReferenceResolvesToAbsoluteLocation()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("AbsoluteNamedRange", new ExcelRangeBase(sheet, "$C$3"));
				sheet.Cells[3, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "AbsoluteNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithoutOffsetsResolvestoSameRowAsCellBeingEvaluated()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, "$C1"));
				sheet.Cells[1, 3].Value = "Wrong";
				sheet.Cells[4, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithPositiveRowOffsetResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, "$C6"));
				sheet.Cells[1, 3].Value = "Very Wrong";
				sheet.Cells[6, 3].Value = "Wrong";
				sheet.Cells[9, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithNegativeRowOffsetResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"$C{ExcelPackage.MaxRows}"));
				sheet.Cells[ExcelPackage.MaxRows, 3].Value = "Wrong";
				sheet.Cells[1, 3].Value = "Wrong";
				sheet.Cells[3, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeColumnNamedRangeWithoutOffsetsResolvestoSameColumnAsCellBeingEvaluated()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, "A$3"));
				sheet.Cells[3, 1].Value = "Wrong";
				sheet.Cells[3, 4].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithPositiveColumnOffsetResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, "B$3"));
				sheet.Cells[3, 1].Value = "Very Wrong";
				sheet.Cells[3, 2].Value = "Wrong";
				sheet.Cells[3, 5].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithNegativeColumnOffsetResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"{ExcelCellAddress.GetColumnLetter(ExcelPackage.MaxColumns)}$3"));
				sheet.Cells[3, ExcelPackage.MaxColumns].Value = "Wrong";
				sheet.Cells[3, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithPositiveRowAndColumnOffsetsResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, "B2"));
				sheet.Cells[2, 2].Value = "Wrong";
				sheet.Cells[5, 5].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithNegativeRowAndColumnOffsetsResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"{ExcelCellAddress.GetColumnLetter(ExcelPackage.MaxColumns)}{ExcelPackage.MaxRows}"));
				sheet.Cells[ExcelPackage.MaxRows, ExcelPackage.MaxColumns].Value = "Wrong";
				sheet.Cells[3, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithPositiveRowAndNegativeColumnOffsetResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"{ExcelCellAddress.GetColumnLetter(ExcelPackage.MaxColumns)}2"));
				sheet.Cells[2, ExcelPackage.MaxColumns].Value = "Wrong";
				sheet.Cells[5, 3].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeWithNegativeRowAndPositiveColumnOffsetResolvesCorrectly()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet");
				sheet.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"B{ExcelPackage.MaxRows}"));
				sheet.Cells[ExcelPackage.MaxRows, ExcelPackage.MaxColumns].Value = "Wrong";
				sheet.Cells[3, 5].Value = "Correct!";
				sheet.Cells[4, 4].Formula = "RelativeNamedRange";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[4, 4].Value);
			}
		}

		[TestMethod]
		public void RelativeNamedRangeResolvesDependencies()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				excelPackage.Workbook.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"$E1"));
				sheet.Cells[3, 4].Formula = "RelativeNamedRange";
				sheet.Cells[3, 5].Formula = "F3";
				sheet.Cells[3, 6].Value = "Correct!";
				sheet.Calculate();
				Assert.AreEqual("Correct!", sheet.Cells[3, 4].Value);
			}
		}

		[TestMethod]
		public void ReferencedRelativeNamedRangeResolvesDependencies()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var sheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				excelPackage.Workbook.Names.Add("RelativeNamedRange", new ExcelRangeBase(sheet, $"$E1"));
				sheet.Cells[3, 4].Formula = @"IF(RelativeNamedRange=""Correct!"", true, false)";
				sheet.Cells[3, 5].Formula = "F3";
				sheet.Cells[3, 6].Value = "Correct!";
				sheet.Calculate();
				Assert.AreEqual(true, sheet.Cells[3, 4].Value);
			}
		}
		#endregion
	}
}
