using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using static OfficeOpenXml.ExcelAddress;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelAddressTest
	{
		#region ExcelAddressBase Tests
		#region Address Tests
		[TestMethod]
		public void ExcelAddressBaseWithFullyQualifiedEndReferenceSheetNameInQuotes()
		{
			var address = "Sheet2!B2:'Sheet2'!C2";
			var result = new ExcelAddress(address);
			Assert.AreEqual(2, result._fromRow);
			Assert.AreEqual(2, result._fromCol);
			Assert.AreEqual(2, result._toRow);
			Assert.AreEqual(3, result._toCol);
			Assert.IsFalse(result._fromRowFixed);
			Assert.IsFalse(result._fromColFixed);
			Assert.IsFalse(result._toRowFixed);
			Assert.IsFalse(result._toColFixed);
			Assert.AreEqual("Sheet2", result.WorkSheet);
		}

		[TestMethod]
		public void ExcelAddressBaseWithFullyQualifiedEndReference()
		{
			var address = "'Sheet2'!B2:Sheet2!C2";
			var result = new ExcelAddress(address);
			Assert.AreEqual(2, result._fromRow);
			Assert.AreEqual(2, result._fromCol);
			Assert.AreEqual(2, result._toRow);
			Assert.AreEqual(3, result._toCol);
			Assert.IsFalse(result._fromRowFixed);
			Assert.IsFalse(result._fromColFixed);
			Assert.IsFalse(result._toRowFixed);
			Assert.IsFalse(result._toColFixed);
			Assert.AreEqual("Sheet2", result.WorkSheet);
		}

		[TestMethod]
		public void ExcelAddressBase_Address()
		{
			var excelAddress = new ExcelAddress("C3");
			Assert.AreEqual("C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddressBase_FullColumn()
		{
			var excelAddress = new ExcelAddress("C:C");
			Assert.AreEqual("C:C", excelAddress.Address);
			Assert.IsFalse(excelAddress._fromRowFixed);
			Assert.IsFalse(excelAddress._fromColFixed);
			Assert.IsFalse(excelAddress._toRowFixed);
			Assert.IsFalse(excelAddress._toColFixed);
		}

		[TestMethod]
		public void ExcelAddressBase_FullColumnAbsolute()
		{
			var excelAddress = new ExcelAddress("$C:$C");
			Assert.AreEqual("$C:$C", excelAddress.Address);
			Assert.IsFalse(excelAddress._fromRowFixed);
			Assert.IsTrue(excelAddress._fromColFixed);
			Assert.IsFalse(excelAddress._toRowFixed);
			Assert.IsTrue(excelAddress._toColFixed);
		}

		[TestMethod]
		public void ExcelAddressBase_FullRow()
		{
			var excelAddress = new ExcelAddress("5:5");
			Assert.AreEqual("5:5", excelAddress.Address);
			Assert.IsFalse(excelAddress._fromRowFixed);
			Assert.IsFalse(excelAddress._fromColFixed);
			Assert.IsFalse(excelAddress._toRowFixed);
			Assert.IsFalse(excelAddress._toColFixed);
		}

		[TestMethod]
		public void ExcelAddressBase_FullRowAbsolute()
		{
			var excelAddress = new ExcelAddress("$5:$5");
			Assert.AreEqual("$5:$5", excelAddress.Address);
			Assert.IsTrue(excelAddress._fromRowFixed);
			Assert.IsFalse(excelAddress._fromColFixed);
			Assert.IsTrue(excelAddress._toRowFixed);
			Assert.IsFalse(excelAddress._toColFixed);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressWithWorksheet()
		{
			var excelAddress = new ExcelAddress("worksheet!C3");
			Assert.AreEqual("worksheet!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressWithQuotedWorksheet()
		{
			var excelAddress = new ExcelAddress("'worksheet'!C3");
			Assert.AreEqual("'worksheet'!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressList()
		{
			var excelAddress = new ExcelAddress("C3,D4,E5");
			Assert.AreEqual("C3,D4,E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressListWithWorksheets()
		{
			var excelAddress = new ExcelAddress("worksheet!C3,worksheet!D4,worksheet!E5");
			Assert.AreEqual("worksheet!C3,worksheet!D4,worksheet!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressListWithQuotedWorksheets()
		{
			var excelAddress = new ExcelAddress("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5");
			Assert.AreEqual("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressWithWorksheetWithExclamationPointInName()
		{
			var excelAddress = new ExcelAddress("'work!sheet'!C3");
			Assert.AreEqual("work!sheet", excelAddress.WorkSheet);
		}
		#endregion

		#region FullAddress Tests
		[TestMethod]
		public void FullAddress()
		{
			var excelAddress = new ExcelAddress("[workbook]worksheet!C3");
			Assert.AreEqual("[workbook]worksheet!C3", excelAddress.FullAddress);
		}

		[TestMethod]
		public void FullAddressList()
		{
			var excelAddress = new ExcelAddress("C3,D4,E5");
			Assert.AreEqual("C3,D4,E5", excelAddress.FullAddress);
		}

		[TestMethod]
		public void FullAddressListWithSheetnames()
		{
			var excelAddress = new ExcelAddress("Sheet!C3,Sheet!D4,Sheet!E5");
			Assert.AreEqual("'Sheet'!C3,'Sheet'!D4,'Sheet'!E5", excelAddress.FullAddress);
		}

		[TestMethod]
		public void FullAddressListFollowingAddressesInheritFirstSheet()
		{
			var excelAddress = new ExcelAddress("Sheet!C3,Sheet2!D4,Sheet3!E5");
			Assert.AreEqual("'Sheet'!C3,'Sheet'!D4,'Sheet'!E5", excelAddress.FullAddress);
		}


		[TestMethod]
		public void ExcelAddressBaseWithDoubledSheetNameNoQuotes()
		{
			var address = new ExcelAddress("Sheet1!A1:Sheet1!A3");
			Assert.AreEqual("A1:A3", address.FirstAddress);
		}

		[TestMethod]
		public void ExcelAddressBaseWithDoubledSheetNameAllQuotes()
		{
			var address = new ExcelAddress("'Sheet 1'!A1:'Sheet 1'!A3");
			Assert.AreEqual("A1:A3", address.FirstAddress);
		}

		[TestMethod]
		public void ExcelAddressBaseWithDoubledSheetNameFirstQuotedOnly()
		{
			var address = new ExcelAddress("'Sheet1'!A1:Sheet1!A3");
			Assert.AreEqual("A1:A3", address.FirstAddress);
		}

		[TestMethod]
		public void ExcelAddressBaseWithDoubledSheetNameLastQuotedOnly()
		{
			var address = new ExcelAddress("Sheet1!A1:'Sheet1'!A3");
			Assert.AreEqual("A1:A3", address.FirstAddress);
		}
		#endregion

		#region Workbook Tests
		[TestMethod]
		public void Workbook()
		{
			var excelAddress = new ExcelAddress("[workbook]worksheet!C3");
			Assert.AreEqual("workbook", excelAddress.Workbook);
		}

		[TestMethod]
		public void WorkbookNotSet()
		{
			var excelAddress = new ExcelAddress("worksheet!C3");
			Assert.AreEqual(null, excelAddress.Workbook);
		}
		#endregion

		#region ChangeWorksheet Tests
		[TestMethod]
		public void ChangeWorksheet()
		{
			var excelAddress = new ExcelAddress("Sheet!C3");
			Assert.AreEqual("Sheet", excelAddress.WorkSheet);
			excelAddress.ChangeWorksheet("Sheet", "NewSheet");
			Assert.AreEqual("NewSheet", excelAddress.WorkSheet);
		}

		[TestMethod]
		public void ChangeWorksheetAppliesToNestedAddresses()
		{
			var excelAddress = new ExcelAddress("Sheet!C3,Sheet!D4,Sheet!E5");
			Assert.AreEqual("Sheet!C3,Sheet!D4,Sheet!E5", excelAddress.Address);
			excelAddress.ChangeWorksheet("Sheet", "NewSheet");
			Assert.AreEqual("'NewSheet'!C3,'NewSheet'!D4,'NewSheet'!E5", excelAddress.Address);
		}

		[TestMethod]
		public void ChangeWorksheetErrorsWorksheetReference()
		{
			var excelAddress = new ExcelAddress("Sheet!C3");
			Assert.AreEqual("Sheet!C3", excelAddress.Address);
			excelAddress.ChangeWorksheet("Sheet", null);
			Assert.AreEqual("#REF!C3", excelAddress.Address);
		}

		[TestMethod]
		public void ChangeWorksheetIgnoresSheetNameCase()
		{
			var excelAddress = new ExcelAddress("'SHEET1'!B2");
			Assert.AreEqual("'SHEET1'!B2", excelAddress.Address);
			excelAddress.ChangeWorksheet("Sheet1", "Sheet1 Copy");
			Assert.AreEqual("'Sheet1 Copy'!B2", excelAddress.Address);
		}

		[TestMethod]
		public void ChangeWorksheetAppliesToNestedAddressesMultiSheetList()
		{
			Assert.Fail("This test will fail. We suspect that the true error is that " +
					"a list of addresses that span multiple sheets is actually an invalid " +
					"state for ExcelAddressBase. Since we didn't write it, we don't know if " +
					"it's safe for us to assert that in the constructor." +
					"If some other bug comes through related to this we can revisit it.");
			var excelAddress = new ExcelAddress("Sheet!C3,Sheet!D4,OtherSheet!E5");
			Assert.AreEqual("Sheet!C3,Sheet!D4,OtherSheet!E5", excelAddress.Address);
			excelAddress.ChangeWorksheet("Sheet", "NewSheet");
			Assert.AreEqual("'NewSheet'!C3,'NewSheet'!D4,OtherSheet!E5", excelAddress.Address);
		}

		[TestMethod]
		public void ChangeWorksheetNotApplied()
		{
			var excelAddress = new ExcelAddress("Sheet!C3");
			Assert.AreEqual("Sheet", excelAddress.WorkSheet);
			excelAddress.ChangeWorksheet("OtherSheet", "NewSheet");
			Assert.AreEqual("Sheet", excelAddress.WorkSheet);
		}
		#endregion

		#region AddRow Tests
		[TestMethod]
		public void AddRowBeforeFromRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.AddRow(2, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(6, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeToRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.AddRow(4, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowAfterToRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.AddRow(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFromRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(2, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(6, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeToRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowAfterToRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(2, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeToRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowAfterToRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFixedFromRowFullColumnUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C$1:$E$1048576 => insert 3 rows at row 1 => $C$1:$E$1048576 (no change).
			var excelAddress = new ExcelAddress(1, 3, ExcelPackage.MaxRows, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(1, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(1, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFixedFromRowUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C$3:$E$5 => insert 3 rows at row 2 => $F$3:$H$8.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(2, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(6, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeRelativeFromRowUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C3:$E$5 => insert 3 rows at row 2 => $C3:$E$8.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, false, true, true, true);
			var newAddress = excelAddress.AddRow(2, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFixedToRowUpdateOnlyFixedAddresses()
		{
			// Update address as a reference in a named range formula.
			// $C$3:$E$5 => insert 3 rows at row 4 => $C$3:$E$8.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeRelativeToRowUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C$3:$E5 => insert 3 rows at row 4 => $C$3:$E5.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, false, true);
			var newAddress = excelAddress.AddRow(4, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingLessThanMinimumRowRelative()
		{
			var excelAddress = new ExcelAddress("A:A");
			excelAddress = excelAddress.AddRow(0, 1);
			Assert.AreEqual("A:A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingLessThanMinimumRowAbsolute()
		{
			var excelAddress = new ExcelAddress("$A:$A");
			excelAddress = excelAddress.AddRow(0, 1);
			Assert.AreEqual("$A:$A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingLessThanMinimumRowRelativeSetFixed()
		{
			var excelAddress = new ExcelAddress("A:A");
			excelAddress = excelAddress.AddRow(0, 1, true);
			Assert.AreEqual("A:A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingLessThanMinimumRowAbsoluteSetFixed()
		{
			var excelAddress = new ExcelAddress("$A:$A");
			excelAddress = excelAddress.AddRow(0, 1, true);
			Assert.AreEqual("$A:$A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingLessThanMinimumRowRelative()
		{
			var excelAddress = new ExcelAddress("1:1");
			excelAddress = excelAddress.AddColumn(0, 1);
			Assert.AreEqual("1:1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingLessThanMinimumRowAbsolute()
		{
			var excelAddress = new ExcelAddress("$1:$1");
			excelAddress = excelAddress.AddColumn(0, 1);
			Assert.AreEqual("$1:$1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingLessThanMinimumRowRelativeSetFixed()
		{
			var excelAddress = new ExcelAddress("1:1");
			excelAddress = excelAddress.AddColumn(0, 1, true);
			Assert.AreEqual("1:1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingLessThanMinimumRowAbsoluteSetFixed()
		{
			var excelAddress = new ExcelAddress("$1:$1");
			excelAddress = excelAddress.AddColumn(0, 1, true);
			Assert.AreEqual("$1:$1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingValidRowRelative()
		{
			var excelAddress = new ExcelAddress("A:A");
			excelAddress = excelAddress.AddRow(50, 1);
			Assert.AreEqual("A:A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingValidRowAbsolute()
		{
			var excelAddress = new ExcelAddress("$A:$A");
			excelAddress = excelAddress.AddRow(50, 1);
			Assert.AreEqual("$A:$A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingValidRowRelativeSetFixed()
		{
			var excelAddress = new ExcelAddress("A:A");
			excelAddress = excelAddress.AddRow(50, 1, true);
			Assert.AreEqual("A:A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingValidRowAbsoluteSetFixed()
		{
			var excelAddress = new ExcelAddress("$A:$A");
			excelAddress = excelAddress.AddRow(50, 1, true);
			Assert.AreEqual("$A:$A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingValidColumnRelative()
		{
			var excelAddress = new ExcelAddress("1:1");
			excelAddress = excelAddress.AddColumn(50, 1);
			Assert.AreEqual("1:1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingValidColumnAbsolute()
		{
			var excelAddress = new ExcelAddress("$1:$1");
			excelAddress = excelAddress.AddColumn(50, 1);
			Assert.AreEqual("$1:$1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingValidColumnRelativeSetFixed()
		{
			var excelAddress = new ExcelAddress("1:1");
			excelAddress = excelAddress.AddColumn(50, 1, true);
			Assert.AreEqual("1:1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingValidColumnAbsoluteSetFixed()
		{
			var excelAddress = new ExcelAddress("$1:$1");
			excelAddress = excelAddress.AddColumn(50, 1, true);
			Assert.AreEqual("$1:$1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingGreaterThanMaxRowRelative()
		{
			var excelAddress = new ExcelAddress("A:A");
			excelAddress = excelAddress.AddRow(ExcelPackage.MaxRows + 1, 1);
			Assert.AreEqual("A:A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingGreaterThanMaxRowAbsolute()
		{
			var excelAddress = new ExcelAddress("$A:$A");
			excelAddress = excelAddress.AddRow(ExcelPackage.MaxRows + 1, 1);
			Assert.AreEqual("$A:$A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingGreaterThanMaxRowRelativeSetFixed()
		{
			var excelAddress = new ExcelAddress("A:A");
			excelAddress = excelAddress.AddRow(ExcelPackage.MaxRows + 1, 1, true);
			Assert.AreEqual("A:A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddRowToFullColumnDoesNothingGreaterThanMaxRowAbsoluteSetFixed()
		{
			var excelAddress = new ExcelAddress("$A:$A");
			excelAddress = excelAddress.AddRow(ExcelPackage.MaxRows + 1, 1, true);
			Assert.AreEqual("$A:$A", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(ExcelPackage.MaxRows, excelAddress.End.Row);
			Assert.AreEqual(1, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingGreaterThanMaxColumnRelative()
		{
			var excelAddress = new ExcelAddress("1:1");
			excelAddress = excelAddress.AddColumn(ExcelPackage.MaxColumns + 1, 1);
			Assert.AreEqual("1:1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingGreaterThanMaxColumnAbsolute()
		{
			var excelAddress = new ExcelAddress("$1:$1");
			excelAddress = excelAddress.AddColumn(ExcelPackage.MaxColumns + 1, 1);
			Assert.AreEqual("$1:$1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingGreaterThanMaxColumnRelativeSetFixed()
		{
			var excelAddress = new ExcelAddress("1:1");
			excelAddress = excelAddress.AddColumn(ExcelPackage.MaxColumns + 1, 1, true);
			Assert.AreEqual("1:1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnToFullRowDoesNothingGreaterThanMaxColumnAbsoluteSetFixed()
		{
			var excelAddress = new ExcelAddress("$1:$1");
			excelAddress = excelAddress.AddColumn(ExcelPackage.MaxColumns + 1, 1, true);
			Assert.AreEqual("$1:$1", excelAddress.Address);
			Assert.AreEqual(1, excelAddress.Start.Row);
			Assert.AreEqual(1, excelAddress.Start.Column);
			Assert.AreEqual(1, excelAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, excelAddress.End.Column);
		}
		#endregion

		#region DeleteRow Tests
		[TestMethod]
		public void DeleteRowAfterToRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFromRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(1, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(1, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInside()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteRowPartialBeforeFromRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(2, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRow()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(4, 1);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowOverDelete()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(4, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowAfterToRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFromRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(1, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(1, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInsideFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteRowPartialBeforeFromRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 1);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowOverDeleteFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowAfterToRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(6, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(1, 2, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInsideFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 4, true);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteRowPartialBeforeFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 2, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowRangeExtendsBeyondFixedRange()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(5, 10);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 1, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowOverDeleteFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 2, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFixedUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeRelativeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, false, true, false, true);
			var newAddress = excelAddress.DeleteRow(2, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeRelativeFromColumnUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, false, true);
			var newAddress = excelAddress.DeleteRow(2, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeRelativeToColumnUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, false, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 1, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInFixedRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 1, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInRelativeRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, false, true, false, true);
			var newAddress = excelAddress.DeleteRow(4, 1, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowRangeExtendsBeyondFixedRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(5, 10, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowRangeExtendsBeyondRelativeRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, false, true, false, true);
			var newAddress = excelAddress.DeleteRow(5, 10, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowAfterFixedUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(6, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowAfterRelativeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, false, true, false, true);
			var newAddress = excelAddress.DeleteRow(6, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowEntireRelativeRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, false, true, false, true);
			var newAddress = excelAddress.DeleteRow(2, 10, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowEntireFixedFromRowRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, false, true);
			Assert.IsNull(excelAddress.DeleteRow(2, 10, false, true));
		}

		[TestMethod]
		public void DeleteRowEntireFixedToRowRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, false, true, true, true);
			Assert.IsNull(excelAddress.DeleteRow(2, 10, false, true));
		}

		[TestMethod]
		public void DeleteRowEntireFixedRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			Assert.IsNull(excelAddress.DeleteRow(2, 10, false, true));
		}
		#endregion

		#region AddColumn Tests
		[TestMethod]
		public void AddColumnAfterToColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.AddColumn(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFromColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.AddColumn(2, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(6, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeToColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.AddColumn(4, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnAfterToColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFromColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(2, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(6, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeToColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(4, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnAfterToColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(6, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(2, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeToColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(4, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFixedFromColumnFullRowUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $A$3:$XFD$5 => insert 3 columns at column 1 => $A$3:$XFD$5 (no change).
			var excelAddress = new ExcelAddress(3, 1, 5, ExcelPackage.MaxColumns, true, true, true, true);
			var newAddress = excelAddress.AddColumn(1, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(1, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(ExcelPackage.MaxColumns, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFixedFromColumnUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C$3:$E$5 => insert 3 columns at column 2 => $F$3:$H$5.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(2, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(6, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeRelativeFromColumnUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C3:$E$5 => insert 3 columns at column 2 => $C3:$H$5.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, false, true, true);
			var newAddress = excelAddress.AddColumn(2, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFixedToColumnUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C$3:$E$5 => insert 3 columns at column 4 => $C$3:$H$5.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(4, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeRelativeToColumnUpdateOnlyFixed()
		{
			// Update address as a reference in a named range formula.
			// $C$3:$E5 => insert 3 columns at column 4 => $C$3:$E5.
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, false);
			var newAddress = excelAddress.AddColumn(4, 3, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}
		#endregion

		#region DeleteColumn Tests
		[TestMethod]
		public void DeleteColumnAfterToColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFromColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(1, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(1, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInside()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteColumnPartialBeforeFromColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(2, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumn()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(4, 1);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnOverDelete()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(4, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnAfterToColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(6, 3);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFromColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(1, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(1, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInsideFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteColumnPartialBeforeFromColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 1);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnOverDeleteFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 2);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnRangeExtendsBeyondFixedRange()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(5, 10);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnAfterToColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(6, 3, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(1, 2, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInsideFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 4, true);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteColumnPartialBeforeFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 2, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 1, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnOverDeleteFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 2, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFixedUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeRelativeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, false, true, false);
			var newAddress = excelAddress.DeleteColumn(2, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeRelativeFromColumnUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, false);
			var newAddress = excelAddress.DeleteColumn(2, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeRelativeToColumnUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, false, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 1, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInFixedRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 1, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInRelativeRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, false, true, false);
			var newAddress = excelAddress.DeleteColumn(4, 1, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnRangeExtendsBeyondFixedRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(5, 10, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnRangeExtendsBeyondRelativeRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 6, 6, true, false, true, false);
			var newAddress = excelAddress.DeleteColumn(5, 10, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(6, newAddress.End.Row);
			Assert.AreEqual(6, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnAfterFixedUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(6, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnAfterRelativeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, false, true, false);
			var newAddress = excelAddress.DeleteColumn(6, 2, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnEntireRelativeRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, false, true, false);
			var newAddress = excelAddress.DeleteColumn(2, 10, false, true);
			Assert.AreEqual(excelAddress.WorkSheet, newAddress.WorkSheet);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnEntireFixedFromRowRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, false);
			Assert.IsNull(excelAddress.DeleteColumn(2, 10, false, true));
		}

		[TestMethod]
		public void DeleteColumnEntireFixedToRowRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, false, true, true);
			Assert.IsNull(excelAddress.DeleteColumn(2, 10, false, true));
		}

		[TestMethod]
		public void DeleteColumnEntireFixedRangeUpdateOnlyFixed()
		{
			var excelAddress = new ExcelAddress(3, 3, 5, 5, true, true, true, true);
			Assert.IsNull(excelAddress.DeleteColumn(2, 10, false, true));
		}
		#endregion

		#region IsValidRowCol Tests
		[TestMethod]
		public void IsValidRowCol()
		{
			var excelAddress = new ExcelAddress(1, 2, 3, 4);
			Assert.IsTrue(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColFromRowAfterToRow()
		{
			var excelAddress = new ExcelAddress
			{
				_fromRow = 4,
				_fromCol = 2,
				_toRow = 3,
				_toCol = 4
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColFromColAfterToCol()
		{
			var excelAddress = new ExcelAddress
			{
				_fromRow = 1,
				_fromCol = 5,
				_toRow = 3,
				_toCol = 4
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColFromRowTooLow()
		{
			var excelAddress = new ExcelAddress
			{
				_fromRow = 0,
				_fromCol = 2,
				_toRow = 3,
				_toCol = 4
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColFromColTooLow()
		{
			var excelAddress = new ExcelAddress
			{
				_fromRow = 1,
				_fromCol = 0,
				_toRow = 3,
				_toCol = 4
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColToRowTooHigh()
		{
			var excelAddress = new ExcelAddress
			{
				_fromRow = 1,
				_fromCol = 2,
				_toRow = ExcelPackage.MaxRows + 1,
				_toCol = 4
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColToColTooHigh()
		{
			var excelAddress = new ExcelAddress
			{
				_fromRow = 1,
				_fromCol = 2,
				_toRow = 3,
				_toCol = ExcelPackage.MaxColumns + 1
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}
		#endregion

		#region IsValid Tests
		[TestMethod]
		public void IsValidInternalAddressTest()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("C5"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("C5:D6"));
		}

		[TestMethod]
		public void IsValidInternalAddressWithSheetTest()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5:D6"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5:Sheet1!D6"));
		}

		[TestMethod]
		public void IsValidInternalAddressWithQuotedSheetTest()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5:D6"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5:Sheet1!D6"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5:'Sheet1'!D6"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5:'Sheet1'!D6"));
		}

		[TestMethod]
		public void IsValidInternalAddressMaxColumns()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet!1:1"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet!1:3"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet'!1:1"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet'!1:3"));
		}

		[TestMethod]
		public void IsValidInternalAddressMaxRows()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet!A:A"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet!A:C"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet'!A:A"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet'!A:C"));
		}

		[TestMethod]
		public void IsValidInternalAddressCommaSeparated()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("C5, C6,G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5, C6,G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5, Sheet1!C6,G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5, Sheet1!C6,Sheet1!G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5, Sheet2!C6,Sheet3!G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5, C6,G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5, 'Sheet1'!C6,G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5, 'Sheet1'!C6,'Sheet1'!G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5, 'Sheet2'!C6,'Sheet3'!G7"));
		}

		[TestMethod]
		public void IsValidInternalAddressCommaSeparatedRanges()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("C5:D7, C6:F8,G7:P8"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5:Sheet1!D7, C6:F8,G7:P8"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5:D7, Sheet1!C6:F8,G7:P11"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("Sheet1!C5, Sheet1!C6:F8,Sheet1!G7:Sheet1!M12"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5:D7, C6,G7:P8"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5, 'Sheet1'!C6:'Sheet1'!F8,G7"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'Sheet1'!C5:Sheet1!D7, 'Sheet2'!C6:F9"));
		}

		[TestMethod]
		public void IsValidInternalAddressWithEdgeCaseSheetNameTest()
		{
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'#REF!'!C5"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'#REF!'!C5:D6"));
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'#REF'!C5:D6"));
			// Sheet names can contain single quotes if they are in between other characters and escaped in references. 
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'a''b'!C5"));
			// Sheet names can contain double quotes.
			Assert.AreEqual(AddressType.InternalAddress, ExcelAddress.IsValid("'\"'!C5"));
		}

		[TestMethod]
		public void IsValidErrorInternalAddressTest()
		{
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!C5"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("Sheet1!#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("'Sheet1'!#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!:#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!:C5"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("C5:#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("C5:D7,#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!#REF!:#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!#REF!:#REF!#REF!"));

			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!, #REF!,#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!#REF!, #REF!,#REF!"));
			Assert.AreEqual(AddressType.Invalid, ExcelAddress.IsValid("#REF!#REF!, #REF!:C5,#REF!C6:Sheet1!#REF!"));
		}

		[TestMethod]
		public void IsValidExternalAddresTest()
		{
			Assert.AreEqual(AddressType.ExternalAddress, ExcelAddress.IsValid("[1]Sheet1!C5"));
			Assert.AreEqual(AddressType.ExternalAddress, ExcelAddress.IsValid("'[1]Sheet1'!C5"));
			Assert.AreEqual(AddressType.ExternalAddress, ExcelAddress.IsValid("[1]Sheet1!C5:D6"));
			Assert.AreEqual(AddressType.ExternalAddress, ExcelAddress.IsValid("'[1]Sheet1'!C5:'[1]Sheet1'D6"));
		}
		#endregion

		#region ContainsCoordinate Tests
		[TestMethod]
		public void ContainsCoordinate()
		{
			var excelAddressBase = new ExcelAddress(3, 3, 5, 5);
			Assert.IsTrue(excelAddressBase.ContainsCoordinate(4, 4));      // Inside
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(2, 4));  // Above
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(6, 4));  // Below
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(4, 2));  // Left
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(4, 6));  // Right
		}
		#endregion
		#endregion

		#region ExcelAddress Tests
		#region Address Tests
		[TestMethod]
		public void ExcelAddress_Address()
		{
			var excelAddress = new ExcelAddress("C3");
			Assert.AreEqual("C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
			excelAddress.Address = "C3";
			Assert.AreEqual("C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddress_AddressWithWorksheet()
		{
			var excelAddress = new ExcelAddress("worksheet!C3");
			Assert.AreEqual("worksheet!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
			excelAddress.Address = "worksheet!C3";
			Assert.AreEqual("worksheet!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddress_AddressWithQuotedWorksheet()
		{
			var excelAddress = new ExcelAddress("'worksheet'!C3");
			Assert.AreEqual("'worksheet'!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
			excelAddress.Address = "'worksheet'!C3";
			Assert.AreEqual("'worksheet'!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddress_AddressList()
		{
			var excelAddress = new ExcelAddress("C3,D4,E5");
			Assert.AreEqual("C3,D4,E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
			excelAddress.Address = "C3,D4,E5";
			Assert.AreEqual("C3,D4,E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddress_AddressListWithWorksheets()
		{
			var excelAddress = new ExcelAddress("worksheet!C3,worksheet!D4,worksheet!E5");
			Assert.AreEqual("worksheet!C3,worksheet!D4,worksheet!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
			excelAddress.Address = "worksheet!C3,worksheet!D4,worksheet!E5";
			Assert.AreEqual("worksheet!C3,worksheet!D4,worksheet!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddress_AddressListWithQuotedWorksheets()
		{
			var excelAddress = new ExcelAddress("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5");
			Assert.AreEqual("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
			excelAddress.Address = "'worksheet'!C3,'worksheet'!D4,'worksheet'!E5";
			Assert.AreEqual("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}
		#endregion

		#region Full Row/Column Tests
		[TestMethod]
		public void IsFullColumn()
		{
			var address = new ExcelAddress("C:C");
			Assert.IsTrue(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("$C:$C");
			Assert.IsTrue(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("C:$C");
			Assert.IsTrue(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("$C:C");
			Assert.IsTrue(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("'Sheet1'!$C:$C");
			Assert.IsTrue(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("'Sheet1'!$C:'Sheet1'!$C");
			Assert.IsTrue(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
		}

		[TestMethod]
		public void IsFullRowTest()
		{
			var address = new ExcelAddress("4:4");
			Assert.IsTrue(address._isFullRow);
			Assert.IsFalse(address._isFullColumn);
			address = new ExcelAddress("$4:$4");
			Assert.IsTrue(address._isFullRow);
			Assert.IsFalse(address._isFullColumn);
			address = new ExcelAddress("4:$4");
			Assert.IsTrue(address._isFullRow);
			Assert.IsFalse(address._isFullColumn);
			address = new ExcelAddress("$4:4");
			Assert.IsTrue(address._isFullRow);
			Assert.IsFalse(address._isFullColumn);
			address = new ExcelAddress("'Sheet1'!$4:$4");
			Assert.IsTrue(address._isFullRow);
			Assert.IsFalse(address._isFullColumn);
			address = new ExcelAddress("'Sheet1'!$4:'Sheet1'!$4");
			Assert.IsTrue(address._isFullRow);
			Assert.IsFalse(address._isFullColumn);
		}

		[TestMethod]
		public void IsFullRowsAndColumnsMultiAddressesAreFalse()
		{
			var address = new ExcelAddress("B5");
			Assert.IsFalse(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("B5:C6");
			Assert.IsFalse(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("B5,C:C");
			Assert.IsFalse(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("C:C,B5");
			Assert.IsFalse(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("C:C,B5:D9");
			Assert.IsFalse(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
			address = new ExcelAddress("B5:D9,C:C");
			Assert.IsFalse(address._isFullColumn);
			Assert.IsFalse(address._isFullRow);
		}
		#endregion
		#endregion

		#region ExcelFormulaAddress Tests
		#region Address Tests
		[TestMethod]
		public void ExcelFormulaAddress_Address()
		{
			var excelAddress = new ExcelFormulaAddress("C3");
			Assert.AreEqual("C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
			excelAddress.Address = "C3";
			Assert.AreEqual("C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelFormulaAddress_AddressWithWorksheet()
		{
			var excelAddress = new ExcelFormulaAddress("worksheet!C3");
			Assert.AreEqual("worksheet!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
			excelAddress.Address = "worksheet!C3";
			Assert.AreEqual("worksheet!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelFormulaAddress_AddressWithQuotedWorksheet()
		{
			var excelAddress = new ExcelFormulaAddress("'worksheet'!C3");
			Assert.AreEqual("'worksheet'!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
			excelAddress.Address = "'worksheet'!C3";
			Assert.AreEqual("'worksheet'!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelFormulaAddress_AddressList()
		{
			var excelAddress = new ExcelFormulaAddress("C3,D4,E5");
			Assert.AreEqual("C3,D4,E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
			excelAddress.Address = "C3,D4,E5";
			Assert.AreEqual("C3,D4,E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelFormulaAddress_AddressListWithWorksheets()
		{
			var excelAddress = new ExcelFormulaAddress("worksheet!C3,worksheet!D4,worksheet!E5");
			Assert.AreEqual("worksheet!C3,worksheet!D4,worksheet!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
			excelAddress.Address = "worksheet!C3,worksheet!D4,worksheet!E5";
			Assert.AreEqual("worksheet!C3,worksheet!D4,worksheet!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelFormulaAddress_AddressListWithQuotedWorksheets()
		{
			var excelAddress = new ExcelFormulaAddress("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5");
			Assert.AreEqual("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
			excelAddress.Address = "'worksheet'!C3,'worksheet'!D4,'worksheet'!E5";
			Assert.AreEqual("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}
		#endregion
		#endregion
	}
}