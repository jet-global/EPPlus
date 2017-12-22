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
	public class ZCellStoreTest
	{
		#region GetRowCoordinates Tests
		[TestMethod]
		public void GetRowCoordinates()
		{
			// First cell
			Assert.IsTrue(ZCellStore<object>.GetRowCoordinates(1, out int rowPage, out int pageIndex));
			Assert.AreEqual(0, rowPage);
			Assert.AreEqual(0, pageIndex);
			// Last Cell
			Assert.IsTrue(ZCellStore<object>.GetRowCoordinates(ExcelPackage.MaxRows, out rowPage, out pageIndex));
			Assert.AreEqual(1023, rowPage);
			Assert.AreEqual(1023, pageIndex);
			// 2nd Page First Cell
			Assert.IsTrue(ZCellStore<object>.GetRowCoordinates(1024 + 1, out rowPage, out pageIndex));
			Assert.AreEqual(1, rowPage);
			Assert.AreEqual(0, pageIndex);
			// 4th Page 245 Cell
			Assert.IsTrue(ZCellStore<object>.GetRowCoordinates(1024 + 1024 + 1024 + 245, out rowPage, out pageIndex));
			Assert.AreEqual(3, rowPage);
			Assert.AreEqual(244, pageIndex);
			// Row too small
			Assert.IsFalse(ZCellStore<object>.GetRowCoordinates(0, out rowPage, out pageIndex));
			// Row too large
			Assert.IsFalse(ZCellStore<object>.GetRowCoordinates(ExcelPackage.MaxRows + 1, out rowPage, out pageIndex));
		}
		#endregion

		#region GetColumnCoordinates Tests
		[TestMethod]
		public void GetColumnCoordinates()
		{
			// First column
			Assert.IsTrue(ZCellStore<object>.GetColumnCoordinates(1, out int columnPage, out int pageIndex));
			Assert.AreEqual(0, columnPage);
			Assert.AreEqual(0, pageIndex);
			// Last column
			Assert.IsTrue(ZCellStore<object>.GetColumnCoordinates(ExcelPackage.MaxColumns, out columnPage, out pageIndex));
			Assert.AreEqual(127, columnPage);
			Assert.AreEqual(127, pageIndex);
			// 2nd Page First column
			Assert.IsTrue(ZCellStore<object>.GetColumnCoordinates(128 + 1, out columnPage, out pageIndex));
			Assert.AreEqual(1, columnPage);
			Assert.AreEqual(0, pageIndex);
			// 4th Page 100 column
			Assert.IsTrue(ZCellStore<object>.GetColumnCoordinates(128 + 128 + 128 + 100, out columnPage, out pageIndex));
			Assert.AreEqual(3, columnPage);
			Assert.AreEqual(99, pageIndex);
			// Column too small
			Assert.IsFalse(ZCellStore<object>.GetColumnCoordinates(0, out columnPage, out pageIndex));
			// Column too large
			Assert.IsFalse(ZCellStore<object>.GetColumnCoordinates(ExcelPackage.MaxColumns + 1, out columnPage, out pageIndex));
		}
		#endregion

		#region GetValue Tests
		[TestMethod]
		public void GetValue()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.Initialize(new[]
			{
				new Tuple<int, int, int>(1, 1, 1),
				new Tuple<int, int, int>(2, 2, 2),
				new Tuple<int, int, int>(1024, 10, 3),
				new Tuple<int, int, int>(1024 + 1024 + 3, 128 + 4, 4),
				new Tuple<int, int, int>(1024 + 1024 + 1024 + 500, 128 + 128 + 100, 5),
			});
			var value = cellStore.GetValue(1, 1);
			Assert.AreEqual(1, value);
			value = cellStore.GetValue(2, 2);
			Assert.AreEqual(2, value);
			value = cellStore.GetValue(1024, 10);
			Assert.AreEqual(3, value);
			value = cellStore.GetValue(1024 + 1024 + 3, 128 + 4);
			Assert.AreEqual(4, value);
			value = cellStore.GetValue(1024 + 1024 + 1024 + 500, 128 + 128 + 100);
			Assert.AreEqual(5, value);
			// Non-existent value returns default(T)
			value = cellStore.GetValue(12345, 12345);
			Assert.AreEqual(0, value);
		}

		[TestMethod]
		public void GetValueReturnsDefaultForInvalidCoordinates()
		{
			var cellStore = new ZCellStore<int>();
			// Invalid row too small returns default(T)
			var value = cellStore.GetValue(0, 10);
			Assert.AreEqual(0, value);
			// Invalid row too large returns default(T)
			value = cellStore.GetValue(ExcelPackage.MaxRows + 1, 10);
			Assert.AreEqual(0, value);
			// Invalid column too small returns default(T)
			value = cellStore.GetValue(10, 0);
			Assert.AreEqual(0, value);
			// Invalid column too large returns default(T)
			value = cellStore.GetValue(10, ExcelPackage.MaxColumns + 1);
			Assert.AreEqual(0, value);
		}
		#endregion

		#region SetValue Tests
		[TestMethod]
		public void SetValue()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.Initialize(new[]
			{
				new Tuple<int, int, int>(1024 + 1024 + 3, 128 + 4, 4),
			});
			var value = cellStore.GetValue(1024 + 1024 + 3, 128 + 4);
			Assert.AreEqual(4, value);
			cellStore.SetValue(1024 + 1024 + 3, 128 + 4, 9);
			value = cellStore.GetValue(1024 + 1024 + 3, 128 + 4);
			Assert.AreEqual(9, value);
			// Non-existent value returns default(T)
			value = cellStore.GetValue(12345, 12345);
			Assert.AreEqual(0, value);
			cellStore.SetValue(12345, 12345, 123);
			value = cellStore.GetValue(12345, 12345);
			Assert.AreEqual(123, value);
		}

		[TestMethod]
		public void SetValueIgnoresInvalidCoordinates()
		{
			var cellStore = new ZCellStore<int>();
			// row too small is ignored
			cellStore.SetValue(0, 1, 13);
			// row too large is ignored
			cellStore.SetValue(ExcelPackage.MaxRows + 1, 1, 13);
			// column too small is ignored
			cellStore.SetValue(1, 0, 13);
			// column too large is ignored
			cellStore.SetValue(1, ExcelPackage.MaxColumns + 1, 13);
		}
		#endregion

		#region Exists Tests
		[TestMethod]
		public void Exists()
		{
			var cellStore = new ZCellStore<int>();
			Assert.IsFalse(cellStore.Exists(3, 3));
			cellStore.SetValue(3, 3, 13);
			Assert.IsTrue(cellStore.Exists(3, 3));
			Assert.IsFalse(cellStore.Exists(123984, 12345));
			cellStore.SetValue(123984, 12345, 16);
			Assert.IsTrue(cellStore.Exists(123984, 12345));
		}

		[TestMethod]
		public void ExistsInvalidCoordinatesReturnFalse()
		{
			var cellStore = new ZCellStore<int>();
			// row too small is false
			Assert.IsFalse(cellStore.Exists(0, 1));
			// row too large is false
			Assert.IsFalse(cellStore.Exists(ExcelPackage.MaxRows + 1, 1));
			// column too small is false
			Assert.IsFalse(cellStore.Exists(1, 0));
			// column too large is false
			Assert.IsFalse(cellStore.Exists(1, ExcelPackage.MaxColumns + 1));
		}

		[TestMethod]
		public void ExistsWithOutValue()
		{
			var cellStore = new ZCellStore<int>();
			Assert.IsFalse(cellStore.Exists(3, 3, out int value));
			cellStore.SetValue(3, 3, 13);
			Assert.IsTrue(cellStore.Exists(3, 3, out value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(123984, 12345, out value));
			cellStore.SetValue(123984, 12345, 16);
			Assert.IsTrue(cellStore.Exists(123984, 12345, out value));
			Assert.AreEqual(16, value);
		}

		[TestMethod]
		public void ExistsWithOutValueInvalidCoordinatesReturnFalse()
		{
			var cellStore = new ZCellStore<int>();
			// row too small is false
			Assert.IsFalse(cellStore.Exists(0, 1, out int value));
			// row too large is false
			Assert.IsFalse(cellStore.Exists(ExcelPackage.MaxRows + 1, 1, out value));
			// column too small is false
			Assert.IsFalse(cellStore.Exists(1, 0, out value));
			// column too large is false
			Assert.IsFalse(cellStore.Exists(1, ExcelPackage.MaxColumns + 1, out value));
		}
		#endregion

		#region NextCell Tests
		[TestMethod]
		public void NextCellEmptyStore()
		{
			var cellStore = new ZCellStore<int>();
			int row = 0, column = 0;
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}

		[TestMethod]
		public void NextCellStartsWithZero()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
		}

		[TestMethod]
		public void NextCellFindsNextInRowAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(1, 2, 1); // Next on same page
			cellStore.SetValue(1, 128 + 1, 1); // Next on next page
			cellStore.SetValue(1, 128 + 128 + 128 + 128 + 50, 1); // Next several pages later
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(2, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(129, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(562, column);
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}

		[TestMethod]
		public void NextCellFindsNextInColumnAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(2, 1, 1); // Next on same page
			cellStore.SetValue(1024 + 1, 1, 1); // Next on next page
			cellStore.SetValue(1024 + 1024 + 1024 + 238, 1, 1); // Next several pages later
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(2, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1025, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(3310, row);
			Assert.AreEqual(1, column);
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}

		[TestMethod]
		public void NextCellFindsNextDiagonallyAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(2, 2, 1); // Next on same page
			cellStore.SetValue(1024 + 1, 128 + 1, 1); // Next on next page
			cellStore.SetValue(1024 + 1024 + 1024 + 238, 128 + 128 + 128 + 1, 1); // Next several pages later
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(2, row);
			Assert.AreEqual(2, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1025, row);
			Assert.AreEqual(129, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(3310, row);
			Assert.AreEqual(385, column);
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}
		#endregion


		[TestMethod]
		public void DELME()
		{
			var cellstore = new CellStore<int>();
			cellstore.SetValue(1, 1, 1);
			cellstore.SetValue(1, 2, 2);
			cellstore.SetValue(1, 3, 3);
			cellstore.SetValue(2, 2, 4);
			cellstore.SetValue(3, 3, 5);
			int row = 0, column = 0;
			var found = cellstore.NextCell(ref row, ref column);
			found = cellstore.NextCell(ref row, ref column);
			found = cellstore.NextCell(ref row, ref column);
			found = cellstore.NextCell(ref row, ref column);
			found = cellstore.NextCell(ref row, ref column);
			found = cellstore.NextCell(ref row, ref column);
		}
	}
}
