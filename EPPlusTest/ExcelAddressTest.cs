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
	public class ExcelAddressTest
	{
		#region ExcelAddressBase Tests
		#region Address Tests
		[TestMethod]
		public void ExcelAddressBase_Address()
		{
			var excelAddress = new ExcelAddressBase("C3");
			Assert.AreEqual("C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

        [TestMethod]
        public void ExcelAddressBase_FullColumn()
        {
            var excelAddress = new ExcelAddressBase("C:C");
            Assert.AreEqual("C:C", excelAddress.Address);
            Assert.IsFalse(excelAddress._fromRowFixed);
            Assert.IsFalse(excelAddress._fromColFixed);
            Assert.IsFalse(excelAddress._toRowFixed);
            Assert.IsFalse(excelAddress._toColFixed);
        }

        [TestMethod]
        public void ExcelAddressBase_FullColumnAbsolute()
        {
            var excelAddress = new ExcelAddressBase("$C:$C");
            Assert.AreEqual("$C:$C", excelAddress.Address);
            Assert.IsFalse(excelAddress._fromRowFixed);
            Assert.IsTrue(excelAddress._fromColFixed);
            Assert.IsFalse(excelAddress._toRowFixed);
            Assert.IsTrue(excelAddress._toColFixed);
        }

        [TestMethod]
        public void ExcelAddressBase_FullRow()
        {
            var excelAddress = new ExcelAddressBase("5:5");
            Assert.AreEqual("5:5", excelAddress.Address);
            Assert.IsFalse(excelAddress._fromRowFixed);
            Assert.IsFalse(excelAddress._fromColFixed);
            Assert.IsFalse(excelAddress._toRowFixed);
            Assert.IsFalse(excelAddress._toColFixed);
        }

        [TestMethod]
        public void ExcelAddressBase_FullRowAbsolute()
        {
            var excelAddress = new ExcelAddressBase("$5:$5");
            Assert.AreEqual("$5:$5", excelAddress.Address);
            Assert.IsTrue(excelAddress._fromRowFixed);
            Assert.IsFalse(excelAddress._fromColFixed);
            Assert.IsTrue(excelAddress._toRowFixed);
            Assert.IsFalse(excelAddress._toColFixed);
        }

        [TestMethod]
		public void ExcelAddressBase_AddressWithWorksheet()
		{
			var excelAddress = new ExcelAddressBase("worksheet!C3");
			Assert.AreEqual("worksheet!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressWithQuotedWorksheet()
		{
			var excelAddress = new ExcelAddressBase("'worksheet'!C3");
			Assert.AreEqual("'worksheet'!C3", excelAddress.Address);
			Assert.IsNull(excelAddress.Addresses);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressList()
		{
			var excelAddress = new ExcelAddressBase("C3,D4,E5");
			Assert.AreEqual("C3,D4,E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressListWithWorksheets()
		{
			var excelAddress = new ExcelAddressBase("worksheet!C3,worksheet!D4,worksheet!E5");
			Assert.AreEqual("worksheet!C3,worksheet!D4,worksheet!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}

		[TestMethod]
		public void ExcelAddressBase_AddressListWithQuotedWorksheets()
		{
			var excelAddress = new ExcelAddressBase("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5");
			Assert.AreEqual("'worksheet'!C3,'worksheet'!D4,'worksheet'!E5", excelAddress.Address);
			Assert.AreEqual(3, excelAddress.Addresses.Count);
			Assert.AreEqual("C3", excelAddress.Addresses[0].Address);
			Assert.AreEqual("D4", excelAddress.Addresses[1].Address);
			Assert.AreEqual("E5", excelAddress.Addresses[2].Address);
		}
		#endregion

		#region FullAddress Tests
		[TestMethod]
		public void FullAddress()
		{
			var excelAddress = new ExcelAddressBase("[workbook]worksheet!C3");
			Assert.AreEqual("[workbook]worksheet!C3", excelAddress.FullAddress);
		}

		[TestMethod]
		public void FullAddressList()
		{
			var excelAddress = new ExcelAddressBase("C3,D4,E5");
			Assert.AreEqual("C3,D4,E5", excelAddress.FullAddress);
		}

		[TestMethod]
		public void FullAddressListWithSheetnames()
		{
			var excelAddress = new ExcelAddressBase("Sheet!C3,Sheet!D4,Sheet!E5");
			Assert.AreEqual("'Sheet'!C3,'Sheet'!D4,'Sheet'!E5", excelAddress.FullAddress);
		}

		[TestMethod]
		public void FullAddressListFollowingAddressesInheritFirstSheet()
		{
			var excelAddress = new ExcelAddressBase("Sheet!C3,Sheet2!D4,Sheet3!E5");
			Assert.AreEqual("'Sheet'!C3,'Sheet'!D4,'Sheet'!E5", excelAddress.FullAddress);
		}
		#endregion

		#region Workbook Tests
		[TestMethod]
		public void Workbook()
		{
			var excelAddress = new ExcelAddressBase("[workbook]worksheet!C3");
			Assert.AreEqual("workbook", excelAddress.Workbook);
		}

		[TestMethod]
		public void WorkbookNotSet()
		{
			var excelAddress = new ExcelAddressBase("worksheet!C3");
			Assert.AreEqual(null, excelAddress.Workbook);
		}
		#endregion

		#region ChangeWorksheet Tests
		[TestMethod]
		public void ChangeWorksheet()
		{
			var excelAddress = new ExcelAddressBase("Sheet!C3");
			Assert.AreEqual("Sheet", excelAddress.WorkSheet);
			excelAddress.ChangeWorksheet("Sheet", "NewSheet");
			Assert.AreEqual("NewSheet", excelAddress.WorkSheet);
		}

		[TestMethod]
		public void ChangeWorksheetAppliesToNestedAddresses()
		{
			var excelAddress = new ExcelAddressBase("Sheet!C3,Sheet!D4,Sheet!E5");
			Assert.AreEqual("Sheet!C3,Sheet!D4,Sheet!E5", excelAddress.Address);
			excelAddress.ChangeWorksheet("Sheet", "NewSheet");
			Assert.AreEqual("'NewSheet'!C3,'NewSheet'!D4,'NewSheet'!E5", excelAddress.Address);
		}

		[TestMethod]
		public void ChangeWorksheetAppliesToNestedAddressesMultiSheetList()
		{
			Assert.Fail("This test will fail. We suspect that the true error is that " +
					"a list of addresses that span multiple sheets is actually an invalid " +
					"state for ExcelAddressBase. Since we didn't write it, we don't know if " +
					"it's safe for us to assert that in the constructor." +
					"If some other bug comes through related to this we can revisit it.");
			var excelAddress = new ExcelAddressBase("Sheet!C3,Sheet!D4,OtherSheet!E5");
			Assert.AreEqual("Sheet!C3,Sheet!D4,OtherSheet!E5", excelAddress.Address);
			excelAddress.ChangeWorksheet("Sheet", "NewSheet");
			Assert.AreEqual("'NewSheet'!C3,'NewSheet'!D4,OtherSheet!E5", excelAddress.Address);
		}

		[TestMethod]
		public void ChangeWorksheetNotApplied()
		{
			var excelAddress = new ExcelAddressBase("Sheet!C3");
			Assert.AreEqual("Sheet", excelAddress.WorkSheet);
			excelAddress.ChangeWorksheet("OtherSheet", "NewSheet");
			Assert.AreEqual("Sheet", excelAddress.WorkSheet);
		}
		#endregion

		#region AddRow Tests some other thing
		[TestMethod]
		public void AddRowBeforeFromRow()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.AddRow(2, 3);
			Assert.AreEqual(6, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeToRow()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.AddRow(4, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowAfterToRow()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.AddRow(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFromRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(2, 3);
			Assert.AreEqual(6, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeToRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(8, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowAfterToRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(2, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowBeforeToRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddRowAfterToRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddRow(4, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

        [TestMethod]
        public void AddRowToFullColumnDoesNothingLessThanMinimumRowRelative()
        {
            var excelAddress = new ExcelAddressBase("A:A");
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
            var excelAddress = new ExcelAddressBase("$A:$A");
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
            var excelAddress = new ExcelAddressBase("A:A");
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
            var excelAddress = new ExcelAddressBase("$A:$A");
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
            var excelAddress = new ExcelAddressBase("1:1");
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
            var excelAddress = new ExcelAddressBase("$1:$1");
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
            var excelAddress = new ExcelAddressBase("1:1");
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
            var excelAddress = new ExcelAddressBase("$1:$1");
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
            var excelAddress = new ExcelAddressBase("A:A");
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
            var excelAddress = new ExcelAddressBase("$A:$A");
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
            var excelAddress = new ExcelAddressBase("A:A");
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
            var excelAddress = new ExcelAddressBase("$A:$A");
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
            var excelAddress = new ExcelAddressBase("1:1");
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
            var excelAddress = new ExcelAddressBase("$1:$1");
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
            var excelAddress = new ExcelAddressBase("1:1");
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
            var excelAddress = new ExcelAddressBase("$1:$1");
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
            var excelAddress = new ExcelAddressBase("A:A");
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
            var excelAddress = new ExcelAddressBase("$A:$A");
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
            var excelAddress = new ExcelAddressBase("A:A");
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
            var excelAddress = new ExcelAddressBase("$A:$A");
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
            var excelAddress = new ExcelAddressBase("1:1");
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
            var excelAddress = new ExcelAddressBase("$1:$1");
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
            var excelAddress = new ExcelAddressBase("1:1");
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
            var excelAddress = new ExcelAddressBase("$1:$1");
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
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFromRow()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(1, 2);
			Assert.AreEqual(1, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInside()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteRowPartialBeforeFromRow()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(2, 2);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRow()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(4, 1);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowOverDelete()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteRow(4, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowAfterToRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFromRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(1, 2);
			Assert.AreEqual(1, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInsideFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteRowPartialBeforeFromRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 2);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 1);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(4, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowOverDeleteFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(3, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowAfterToRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(6, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowBeforeFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(1, 2, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowInsideFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 4, true);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteRowPartialBeforeFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(2, 2, true);
			Assert.AreEqual(2, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 1, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteRowPartialAfterFromRowOverDeleteFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteRow(4, 2, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}
		#endregion

		#region AddColumn Tests
		[TestMethod]
		public void AddColumnAfterToColumn()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.AddColumn(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFromColumn()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.AddColumn(2, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(6, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeToColumn()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.AddColumn(4, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnAfterToColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFromColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(2, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(6, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeToColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(4, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(8, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnAfterToColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(6, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(2, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void AddColumnBeforeToColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.AddColumn(4, 3, true);
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
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFromColumn()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(1, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(1, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInside()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteColumnPartialBeforeFromColumn()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(2, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumn()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(4, 1);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnOverDelete()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5);
			var newAddress = excelAddress.DeleteColumn(4, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnAfterToColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(6, 3);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFromColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(1, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(1, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInsideFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 4);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteColumnPartialBeforeFromColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 1);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(4, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnOverDeleteFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 2);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(3, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnAfterToColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(6, 3, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnBeforeFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(1, 2, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnInsideFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 4, true);
			Assert.IsNull(newAddress);
		}

		[TestMethod]
		public void DeleteColumnPartialBeforeFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(2, 2, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(2, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 1, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}

		[TestMethod]
		public void DeleteColumnPartialAfterFromColumnOverDeleteFixedAndSetFixed()
		{
			var excelAddress = new ExcelAddressBase(3, 3, 5, 5, true, true, true, true);
			var newAddress = excelAddress.DeleteColumn(4, 2, true);
			Assert.AreEqual(3, newAddress.Start.Row);
			Assert.AreEqual(3, newAddress.Start.Column);
			Assert.AreEqual(5, newAddress.End.Row);
			Assert.AreEqual(5, newAddress.End.Column);
		}
		#endregion

		#region IsValidRowCol Tests
		[TestMethod]
		public void IsValidRowCol()
		{
			var excelAddress = new ExcelAddressBase(1, 2, 3, 4);
			Assert.IsTrue(excelAddress.IsValidRowCol());
		}

		[TestMethod]
		public void IsValidRowColFromRowAfterToRow()
		{
			var excelAddress = new ExcelAddressBase
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
			var excelAddress = new ExcelAddressBase
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
			var excelAddress = new ExcelAddressBase
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
			var excelAddress = new ExcelAddressBase
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
			var excelAddress = new ExcelAddressBase
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
			var excelAddress = new ExcelAddressBase
			{
				_fromRow = 1,
				_fromCol = 2,
				_toRow = 3,
				_toCol = ExcelPackage.MaxColumns + 1
			};
			Assert.IsFalse(excelAddress.IsValidRowCol());
		}
		#endregion

		#region ContainsCoordinate Tests
		[TestMethod]
		public void ContainsCoordinate()
		{
			var excelAddressBase = new ExcelAddressBase(3, 3, 5, 5);
			Assert.IsTrue(excelAddressBase.ContainsCoordinate(4, 4));		// Inside
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(2, 4));	// Above
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(6, 4));	// Below
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(4, 2));	// Left
			Assert.IsFalse(excelAddressBase.ContainsCoordinate(4, 6));	// Right
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