using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	/// <summary>
	/// Summary description for Address
	/// </summary>
	[TestClass]
	public class Address
	{
		private TestContext testContextInstance;

		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		public TestContext TestContext
		{
			get
			{
				return testContextInstance;
			}
			set
			{
				testContextInstance = value;
			}
		}

		#region Additional test attributes
		//
		// You can use the following additional attributes as you write your tests:
		//
		// Use ClassInitialize to run code before running the first test in the class
		// [ClassInitialize()]
		// public static void MyClassInitialize(TestContext testContext) { }
		//
		// Use ClassCleanup to run code after all tests in a class have run
		// [ClassCleanup()]
		// public static void MyClassCleanup() { }
		//
		// Use TestInitialize to run code before running each test 
		// [TestInitialize()]
		// public void MyTestInitialize() { }
		//
		// Use TestCleanup to run code after each test has run
		// [TestCleanup()]
		// public void MyTestCleanup() { }
		//
		#endregion

		[TestMethod]
		public void InsertDeleteTest()
		{
			var addr = new ExcelAddress("A1:B3");

			Assert.AreEqual(addr.AddRow(2, 4).Address, "A1:B7");
			Assert.AreEqual(addr.AddColumn(2, 4).Address, "A1:F3");
			Assert.AreEqual(addr.DeleteColumn(2, 1).Address, "A1:A3");
			Assert.AreEqual(addr.DeleteRow(2, 2).Address, "A1:B1");

			Assert.AreEqual(addr.DeleteRow(1, 3), null);
			Assert.AreEqual(addr.DeleteColumn(1, 2), null);
		}

		[TestMethod]
		public void Addresses()
		{
			var a1 = new ExcelAddress("SalesData!$K$445");
			var a2 = new ExcelAddress("SalesData!$K$445:$M$449,SalesData!$N$448:$Q$454,SalesData!$L$458:$O$464");
			var a3 = new ExcelAddress("SalesData!$K$445:$L$448");
			//var a4 = new ExcelAddressBase("'[1]Risk]TatTWRForm_TWRWEEKLY20130926090'!$N$527");
			var a5 = new ExcelAddress("Table1[[#All],[Title]]");
			var a6 = new ExcelAddress("Table1[#All]");
			var a7 = new ExcelAddress("Table1[[#Headers],[FirstName]:[LastName]]");
			var a8 = new ExcelAddress("Table1[#Headers]");
			var a9 = new ExcelAddress("Table2[[#All],[SubTotal]]");
			var a10 = new ExcelAddress("Table2[#All]");
			var a11 = new ExcelAddress("Table1[[#All],[Freight]]");
			var a12 = new ExcelAddress("[1]!Table1[[LastName]:[Name]]");
			var a13 = new ExcelAddress("Table1[[#All],[Freight]]");
			var a14 = new ExcelAddress("SalesData!$N$5+'test''1'!$J$33");
		}

		[TestMethod]
		public void IsValidCellAdress()
		{
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("A1"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("A1048576"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("XFD1"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("XFD1048576"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!A1"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!A1048576"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!XFD1"));
			Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!XFD1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("XFD"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("XFD"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A1:A1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A1:XFD1"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A1048576:XFD1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("XFD1:XFD1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!A1:A1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!A1:XFD1"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!A1048576:XFD1048576"));
			Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!XFD1:XFD1048576"));
		}

		[TestMethod]
		public void SplitAddressCorrectlySplitsWorksheetName()
		{
			string worksheetAddress = "'Sheet One'!$B$2:$B$3";
			string worksheetAddress2 = "Sheet1!$B$2:$B$3";
			string workbook, worksheet, address;
			ExcelAddress.SplitAddress(worksheetAddress, out workbook, out worksheet, out address);
			Assert.AreEqual("Sheet One", worksheet);
			Assert.AreEqual("$B$2:$B$3", address);
			ExcelAddress.SplitAddress(worksheetAddress2, out workbook, out worksheet, out address);
			Assert.AreEqual("Sheet1", worksheet);
			Assert.AreEqual("$B$2:$B$3", address);
		}

		[TestMethod]
		public void ShouldHandleWorksheetSpec()
		{
			var address = "Sheet1!A1:Sheet1!A2";
			var excelAddress = new ExcelAddress(address);
			Assert.AreEqual("Sheet1", excelAddress.WorkSheet);
			Assert.AreEqual(1, excelAddress._fromRow);
			Assert.AreEqual(2, excelAddress._toRow);
		}

		[TestMethod]
		public void NamedRangeMovesDownIfRowInsertedAbove()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var namedRange = sheet.Names.Add("NewNamedRange", new ExcelRangeBase(sheet, "$A$2:$C$3"));

				sheet.InsertRow(1, 1);

				Assert.AreEqual("'NEW'!$A$3:$C$4", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeDoesNotChangeIfRowInsertedBelow()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var range = sheet.Cells[2, 1, 3, 3];
				var namedRange = sheet.Names.Add("NewNamedRange", range);
				sheet.InsertRow(4, 1);
				Assert.AreEqual("'NEW'!A2:C3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeExpandsDownIfRowInsertedWithin()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var range = sheet.Cells[2, 1, 3, 3];
				var namedRange = sheet.Names.Add("NewNamedRange", new ExcelRangeBase(sheet, "$A$2:$C$3"));

				sheet.InsertRow(3, 1);

				Assert.AreEqual("'NEW'!$A$2:$C$4", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeMovesRightIfColInsertedBefore()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var namedRange = sheet.Names.Add("NewNamedRange", new ExcelRangeBase(sheet, "$B$2:$D$3"));

				sheet.InsertColumn(1, 1);

				Assert.AreEqual("'NEW'!$C$2:$E$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeUnchangedIfColInsertedAfter()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var range = sheet.Cells[2, 2, 3, 4];
				var namedRange = sheet.Names.Add("NewNamedRange", range);
				sheet.InsertColumn(5, 1);
				Assert.AreEqual("'NEW'!B2:D3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeExpandsToRightIfColInsertedWithin()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var namedRange = sheet.Names.Add("NewNamedRange", new ExcelRangeBase(sheet, "$B$2:$D$3"));

				sheet.InsertColumn(3, 1);

				Assert.AreEqual("'NEW'!$B$2:$E$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeWithWorkbookScopeIsMovedDownIfRowInsertedAbove()
		{
			using (var package = new ExcelPackage())
			{
				var workbook = package.Workbook;
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var namedRange = workbook.Names.Add("NewNamedRange", new ExcelRangeBase(sheet, "$A$2:$C$3"));

				sheet.InsertRow(1, 1);

				Assert.AreEqual("'NEW'!$A$3:$C$4", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeWithWorkbookScopeIsMovedRightIfColInsertedBefore()
		{
			using (var package = new ExcelPackage())
			{
				var workbook = package.Workbook;
				var sheet = package.Workbook.Worksheets.Add("NEW");
				var namedRange = workbook.Names.Add("NewNamedRange", new ExcelRangeBase(sheet, "$B$2:$C$3"));

				sheet.InsertColumn(1, 1);

				Assert.AreEqual("'NEW'!$C$2:$D$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void NamedRangeIsUnchangedForOutOfScopeSheet()
		{
			using (var package = new ExcelPackage())
			{
				var workbook = package.Workbook;
				var sheet1 = package.Workbook.Worksheets.Add("NEW");
				var sheet2 = package.Workbook.Worksheets.Add("NEW2");
				var namedRange = workbook.Names.Add("NewNamedRange", new ExcelRangeBase(sheet2, "$B$2:$C$3"));

				sheet1.InsertColumn(1, 1);

				Assert.AreEqual("'NEW2'!$B$2:$C$3", namedRange.NameFormula);
			}
		}
	}
}
