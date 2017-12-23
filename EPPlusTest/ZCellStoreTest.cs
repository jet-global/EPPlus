using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using static OfficeOpenXml.ZCellStore<int>;

namespace EPPlusTest
{
	[TestClass]
	public class ZCellStoreTest
	{
		/*
		 * How to read these tests:
		 * 
		 * There are two helper methods called BuildRow and BuildColumn.
		 * Since we know the static internal structure of this cell store 
		 * we can use these to explicitly target an internal page and an index on that page.
		 * 
		 * This is helpful because the nature of this data structure makes it hard to picture
		 * where your data is going. With this technique you actually get coordinates into the 
		 * associated page structures.
		 * 
		 * 
		 */

		
		#region GetValue Tests
		[TestMethod]
		public void GetValue()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(this.BuildRow(0, 1), this.BuildColumn(0, 1), 1);
			cellStore.SetValue(this.BuildRow(0, 2), this.BuildColumn(0, 2), 2);
			cellStore.SetValue(this.BuildRow(0, 1024), this.BuildColumn(0, 10), 3);
			cellStore.SetValue(this.BuildRow(2, 3), this.BuildColumn(1, 4), 4);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(2, 100), 5);
			var value = cellStore.GetValue(this.BuildRow(0, 1), this.BuildColumn(0, 1));
			Assert.AreEqual(1, value);
			value = cellStore.GetValue(this.BuildRow(0, 2), this.BuildColumn(0, 2));
			Assert.AreEqual(2, value);
			value = cellStore.GetValue(this.BuildRow(0, 1024), this.BuildColumn(0, 10));
			Assert.AreEqual(3, value);
			value = cellStore.GetValue(this.BuildRow(2, 3), this.BuildColumn(1, 4));
			Assert.AreEqual(4, value);
			value = cellStore.GetValue(this.BuildRow(3, 500), this.BuildColumn(2, 100));
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
			cellStore.SetValue(this.BuildRow(2, 3), this.BuildColumn(1, 4), 4);
			var value = cellStore.GetValue(this.BuildRow(2, 3), this.BuildColumn(1, 4));
			Assert.AreEqual(4, value);
			cellStore.SetValue(this.BuildRow(2, 3), this.BuildColumn(1, 4), 9);
			value = cellStore.GetValue(this.BuildRow(2, 3), this.BuildColumn(1, 4));
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
			cellStore.SetValue(1, this.BuildColumn(1, 1), 1); // Next on next page
			cellStore.SetValue(1, this.BuildColumn(4, 50), 1); // Next several pages later
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(2, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(this.BuildColumn(1, 1), column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(this.BuildColumn(4, 50), column);
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}

		[TestMethod]
		public void NextCellFindsNextInColumnAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(2, 1, 1); // Next on same page
			cellStore.SetValue(this.BuildRow(1, 1), 1, 1); // Next on next page
			cellStore.SetValue(this.BuildRow(3, 238), 1, 1); // Next several pages later
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(2, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(1, 1), row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(3, 238), row);
			Assert.AreEqual(1, column);
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}

		[TestMethod]
		public void NextCellFindsNextDiagonallyAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(2, 2, 1); // Next on same page
			cellStore.SetValue(this.BuildRow(1, 1), this.BuildColumn(1, 1), 1); // Next on next page
			cellStore.SetValue(this.BuildRow(3, 238), this.BuildColumn(3, 1), 1); // Next several pages later
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(2, row);
			Assert.AreEqual(2, column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(1, 1), row);
			Assert.AreEqual(this.BuildColumn(1, 1), column);
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(3, 238), row);
			Assert.AreEqual(this.BuildColumn(3, 1), column);
			Assert.IsFalse(cellStore.NextCell(ref row, ref column));
		}
		#endregion

		#region PrevCell Tests
		[TestMethod]
		public void PrevCellEmptyStore()
		{
			var cellStore = new ZCellStore<int>();
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsFalse(cellStore.PrevCell(ref row, ref column));
		}

		[TestMethod]
		public void PrevCellStartsWithMaxRowAndColumn()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
		}

		[TestMethod]
		public void PrevCellFindsNextInRowAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(1, 2, 1); // Next on same page
			cellStore.SetValue(1, this.BuildColumn(1, 1), 1); // Next on next page
			cellStore.SetValue(1, this.BuildColumn(4, 50), 1); // Next several pages later
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(this.BuildColumn(4, 50), column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(this.BuildColumn(1, 1), column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(2, column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsFalse(cellStore.PrevCell(ref row, ref column));
		}

		[TestMethod]
		public void PrevCellFindsNextInColumnAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(2, 1, 1); // Next on same page
			cellStore.SetValue(this.BuildRow(1, 1), 1, 1); // Next on next page
			cellStore.SetValue(this.BuildRow(3, 238), 1, 1); // Next several pages later
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(3, 238), row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(1, 1), row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(2, row);
			Assert.AreEqual(1, column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsFalse(cellStore.PrevCell(ref row, ref column));
		}

		[TestMethod]
		public void PrevCellFindsNextDiagonallyAcrossPages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			cellStore.SetValue(2, 2, 1); // Next on same page
			cellStore.SetValue(this.BuildRow(1, 1), this.BuildColumn(1, 1), 1); // Next on next page
			cellStore.SetValue(this.BuildRow(3, 238), this.BuildColumn(3, 1), 1); // Next several pages later
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(3, 238), row);
			Assert.AreEqual(this.BuildColumn(3, 1), column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(this.BuildRow(1, 1), row);
			Assert.AreEqual(this.BuildColumn(1, 1), column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(2, row);
			Assert.AreEqual(2, column);
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
			Assert.IsFalse(cellStore.PrevCell(ref row, ref column));
		}
		#endregion

		#region Delete Tests
		[TestMethod]
		public void DeleteRowsNoPageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(2, 0, 3, 0);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(7, 1, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(10, 1, out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(3, 497), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(5, 297), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
		}

		[TestMethod]
		public void DeleteRowsWithSinglePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(1000, 0, 750, 0);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(10, 1, out value));
			Assert.AreEqual(14, value);
			Assert.IsTrue(cellStore.Exists(this.BuildRow(2, 774), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(4, 574), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
		}

		[TestMethod]
		public void DeleteRowsWithMultiplePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(1000, 0, this.BuildRow(2, 250), 0);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(10, 1, out value));
			Assert.AreEqual(14, value);
			Assert.IsTrue(cellStore.Exists(this.BuildRow(1, 250), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(3, 50), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
		}

		[TestMethod]
		public void DeleteColumnsNoPageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.Delete(0, 2, 0, 2);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(1, 8, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(1, 10, out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 1), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(2, 4), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
		}

		[TestMethod]
		public void DeleteColumnsWithSinglePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.Delete(0, this.BuildColumn(1, 1), 0, 75);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(1, 10, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 59), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
		}

		[TestMethod]
		public void DeleteColumnsWithMultiplePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.SetValue(1, this.BuildColumn(5, 37), 25);
			cellStore.Delete(0, 100, 0, this.BuildColumn(3, 1));
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(1, 10, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(2, 36), out value));
			Assert.AreEqual(25, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(5, 37), out value));
		}
		#endregion

		#region Delete (with Shift flag) Tests
		[TestMethod]
		public void DeleteRowsWithShiftNoPageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(2, 0, 3, 0, true);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(7, 1, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(10, 1, out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(3, 497), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(5, 297), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
		}

		[TestMethod]
		public void DeleteRowsWithShiftWithSinglePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(1000, 0, 750, 0, true);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(10, 1, out value));
			Assert.AreEqual(14, value);
			Assert.IsTrue(cellStore.Exists(this.BuildRow(2, 774), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(4, 574), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
		}

		[TestMethod]
		public void DeleteRowsWithShiftWithMultiplePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(1000, 0, this.BuildRow(2, 250), 0, true); 
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(10, 1, out value));
			Assert.AreEqual(14, value);
			Assert.IsTrue(cellStore.Exists(this.BuildRow(1, 250), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(3, 50), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
		}

		[TestMethod]
		public void DeleteRowsWithoutShiftSinglePage()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(2, 0, 20, 0, false);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(10, 1, out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(3,500), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsTrue(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
		}

		[TestMethod]
		public void DeleteRowsWithoutShiftMultiplePages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Delete(2, 0, this.BuildRow(4, 500), 0, false); 
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(10, 1, out value));
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
		}

		[TestMethod]
		public void DeleteColumnsWithShiftNoPageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.Delete(0, 2, 0, 2, true);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(1, 8, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(1, 10, out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 1), out value));
			Assert.AreEqual(15, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(2, 4), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
		}

		[TestMethod]
		public void DeleteColumnsWithShiftWithSinglePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.Delete(0, this.BuildColumn(1, 1), 0, 75, true); 
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(1, 10, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 59), out value));
			Assert.AreEqual(20, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
		}

		[TestMethod]
		public void DeleteColumnsWithShiftWithMultiplePageShifting()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.SetValue(1, this.BuildColumn(5, 37), 25);
			cellStore.Delete(0, 100, 0, this.BuildColumn(3, 1), true); 
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsTrue(cellStore.Exists(1, 10, out value));
			Assert.AreEqual(14, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(2, 36), out value));
			Assert.AreEqual(25, value);
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(5, 37), out value));
		}

		[TestMethod]
		public void DeleteColumnsWithoutShiftSinglePage()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(1, 6), 20);
			cellStore.Delete(0, 2, 0, 20, false);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(1, 10, out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
		}

		[TestMethod]
		public void DeleteColumnsWithoutShiftMultiplePages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.Delete(0, 2, 0, this.BuildColumn(1, 30), false); 
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(1, 10, out value));
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
			Assert.AreEqual(20, value);
		}
		#endregion

		#region Clear Tests
		[TestMethod]
		public void ClearRowsSinglePage()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Clear(2, 0, 20, 0);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(10, 1, out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsTrue(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
		}

		[TestMethod]
		public void ClearRowsMultiplePages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(10, 1, 14);
			cellStore.SetValue(this.BuildRow(3, 500), this.BuildColumn(1, 3), 15);
			cellStore.SetValue(this.BuildRow(5, 300), this.BuildColumn(1, 6), 20);
			cellStore.Clear(2, 0, this.BuildRow(4, 500), 0);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(10, 1, out value));
			Assert.IsFalse(cellStore.Exists(this.BuildRow(3, 500), this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(this.BuildRow(5, 300), this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
		}
		
		[TestMethod]
		public void ClearColumnsSinglePage()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(1, 6), 20);
			cellStore.Clear(0, 2, 0, 20);
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(1, 10, out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.AreEqual(15, value);
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(1, 6), out value));
			Assert.AreEqual(20, value);
		}

		[TestMethod]
		public void ClearColumnsMultiplePages()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 13);
			cellStore.SetValue(1, 10, 14);
			cellStore.SetValue(1, this.BuildColumn(1, 3), 15);
			cellStore.SetValue(1, this.BuildColumn(2, 6), 20);
			cellStore.Clear(0, 2, 0, this.BuildColumn(1, 30));
			Assert.IsTrue(cellStore.Exists(1, 1, out int value));
			Assert.AreEqual(13, value);
			Assert.IsFalse(cellStore.Exists(1, 10, out value));
			Assert.IsFalse(cellStore.Exists(1, this.BuildColumn(1, 3), out value));
			Assert.IsTrue(cellStore.Exists(1, this.BuildColumn(2, 6), out value));
			Assert.AreEqual(20, value);
		}
		#endregion

		#region Nested Class Tests

		#region PagedStructure Tests
		#region Constructor Tests
		[TestMethod]
		public void ConstructorSetsPageDimensions()
		{
			var pagedStructure = new PagedStructure<int>(10);
			Assert.AreEqual(10, pagedStructure.PageBits);
			Assert.AreEqual(1024, pagedStructure.PageSize);
			Assert.AreEqual(1023, pagedStructure.PageMask);
			Assert.AreEqual(ExcelPackage.MaxRows - 1, pagedStructure.MaximumIndex);
			pagedStructure = new PagedStructure<int>(4);
			Assert.AreEqual(4, pagedStructure.PageBits);
			Assert.AreEqual(16, pagedStructure.PageSize);
			Assert.AreEqual(15, pagedStructure.PageMask);
			Assert.AreEqual(255, pagedStructure.MaximumIndex);
		}
		#endregion

		#region GetItem Tests
		[TestMethod]
		public void GetItemFirstItem()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 13, null, null, null },
				{ 3, null, null, null },
				{ null, null, 7, null },
				{ null, null, null, 15 }
			};
			pagedStructure.LoadPages(items);
			Assert.AreEqual(13, pagedStructure.GetItem(0));
		}

		[TestMethod]
		public void GetItemLastItem()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 13, null, null, null },
				{ 3, null, null, null },
				{ null, null, 7, null },
				{ null, null, null, 15 }
			};
			pagedStructure.LoadPages(items);
			Assert.AreEqual(15, pagedStructure.GetItem(15));
		}

		[TestMethod]
		public void GetItemInnerPageInnerIndex()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 13, null, null, null },
				{ 3, null, null, null },
				{ null, null, 7, null },
				{ null, null, null, 15 }
			};
			pagedStructure.LoadPages(items);
			Assert.AreEqual(7, pagedStructure.GetItem(10));
		}

		[TestMethod]
		public void GetItemInnerPageFirstIndex()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 13, null, null, null },
				{ 3, null, null, null },
				{ null, null, 7, null },
				{ null, null, null, 15 }
			};
			pagedStructure.LoadPages(items);
			Assert.AreEqual(3, pagedStructure.GetItem(4));
		}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void GetItemIndexOutOfBoundsTooLow()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.GetItem(-1);
		//}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void GetItemIndexOutOfBoundsTooHigh()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.GetItem(16);
		//}
		#endregion

		#region SetItem Tests
		[TestMethod]
		public void SetItemFirstItem()
		{
			var pagedStructure = new PagedStructure<int>(2);
			pagedStructure.SetItem(0, 13);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(0, pagedStructure.MaximumUsedIndex);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 13, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void SetItemLastItem()
		{
			var pagedStructure = new PagedStructure<int>(2);
			pagedStructure.SetItem(15, 13);
			Assert.AreEqual(15, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, 13 }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void SetItemInnerPageInnerIndex()
		{
			var pagedStructure = new PagedStructure<int>(2);
			pagedStructure.SetItem(10, 13);
			Assert.AreEqual(10, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(10, pagedStructure.MaximumUsedIndex);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, 13, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void SetItemInnerPageFirstIndex()
		{
			var pagedStructure = new PagedStructure<int>(2);
			pagedStructure.SetItem(4, 13);
			Assert.AreEqual(4, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(4, pagedStructure.MaximumUsedIndex);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, null, null, null },
				{ 13, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void SetItemIndexOutOfBoundsTooLow()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.SetItem(-1, 13);
		//}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void SetItemIndexOutOfBoundsTooHigh()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.SetItem(16, 13);
		//}
		#endregion

		#region ShiftItems Tests
		[TestMethod]
		public void ShiftItemsZeroDoesNothing()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(0, 0);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
			pagedStructure.ShiftItems(7, 0);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsFromStartPositiveShiftGoesForward()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(0, 1);
			Assert.AreEqual(1, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, 1, 2, 3 },
				{ 4, 5, 6, 7 },
				{ 8, 9, 10, 11 },
				{ 12, 13, 14, 15 }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsFromMiddlePositiveShiftGoesForward()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(5, 5);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, null, null, null },
				{ null, null, 6, 7 },
				{ 8, 9, 10, 11 }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsFromMiddleMultiplePagesPositiveShiftGoesForward()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(5, 10);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, null, null, null },
				{ null, null, null, null },
				{ null, null, null, 6 },
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsForwardAmountTooHighIsResolvedSilently()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(0, 100);
			Assert.AreEqual(-1, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(-1, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsFromStartNegativeShiftGoesBackwards()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(0, -1);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(14, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 2, 3, 4, 5 },
				{ 6, 7, 8, 9 },
				{ 10, 11, 12, 13 },
				{ 14, 15, 16, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsFromMiddleNegativeShiftGoesBackwards()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(5, -3);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(12, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 9, 10, 11 },
				{ 12, 13, 14, 15 },
				{ 16, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsFromMiddleMultiplePagesNegativeShiftGoesBackwards()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(5, -10);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(5, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 16, null, null },
				{ null, null, null, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ShiftItemsBackwardAmountTooHighIsResolvedSilently()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(0, -100);
			Assert.AreEqual(-1, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(-1, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void ShiftItemsIndexOutOfBoundsTooLow()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.ShiftItems(-1, 13);
		//}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void ShiftItemsIndexOutOfBoundsTooHigh()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.ShiftItems(16, 13);
		//}
		#endregion

		#region ClearItems Tests
		[TestMethod]
		public void ClearItemsFromStart()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ClearItems(0, 2);
			Assert.AreEqual(2, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ null, null, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ClearItemsFromMiddle()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ClearItems(7, 5);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, null },
				{ null, null, null, null },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ClearItemsCountTooHighIsResolvedSilently()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ClearItems(7, 100);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(6, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, null },
				{ null, null, null, null },
				{ null, null, null, null }
			};
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ClearItemsZeroDoesNothing()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ClearItems(0, 0);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		[TestMethod]
		public void ClearItemsNegativeAmountDoesNothing()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ClearItems(0, -10);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			pagedStructure.ValidatePages(items, (row, column, expected) => Assert.Fail($"Row {row} :: Column {column} :: Did not match. {expected}"));
		}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void ClearItemsIndexOutOfBoundsTooLow()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.ClearItems(-1, 2);
		//}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentOutOfRangeException))]
		//public void ClearItemsIndexOutOfBoundsTooHigh()
		//{
		//	var pagedStructure = new PagedStructure<int>(2);
		//	pagedStructure.ClearItems(16, 2);
		//}
		#endregion

		#region NextItem Tests
		[TestMethod]
		public void NextItemFromStart()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 0;
			Assert.IsTrue(pagedStructure.NextItem(ref index));
			Assert.AreEqual(1, index);
		}

		[TestMethod]
		public void NextItemFromMiddle()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 7;
			Assert.IsTrue(pagedStructure.NextItem(ref index));
			Assert.AreEqual(8, index);
		}

		[TestMethod]
		public void NextItemFromEnd()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 15;
			Assert.IsFalse(pagedStructure.NextItem(ref index));
			Assert.AreEqual(16, index);
		}

		[TestMethod]
		public void NextItemSkipsMissingItems()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, null, null },
				{ null, null, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 5;
			Assert.IsTrue(pagedStructure.NextItem(ref index));
			Assert.AreEqual(10, index);
		}

		[TestMethod]
		public void NextItemIndexTooLowFindsFirstItem()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, null, null },
				{ null, null, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = -3;
			Assert.IsTrue(pagedStructure.NextItem(ref index));
			Assert.AreEqual(0, index);
		}

		[TestMethod]
		public void NextItemIndexTooHighFindsNothing()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, null, null },
				{ null, null, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 23;
			Assert.IsFalse(pagedStructure.NextItem(ref index));
			Assert.AreEqual(24, index);
		}
		#endregion

		#region PreviousItem Tests
		[TestMethod]
		public void PreviousItemFromStart()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 0;
			Assert.IsFalse(pagedStructure.PreviousItem(ref index));
			Assert.AreEqual(-1, index);
		}

		[TestMethod]
		public void PreviousItemFromMiddle()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 7;
			Assert.IsTrue(pagedStructure.PreviousItem(ref index));
			Assert.AreEqual(6, index);
		}

		[TestMethod]
		public void PreviousItemFromEnd()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, 7, 8 },
				{ 9, 10, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 15;
			Assert.IsTrue(pagedStructure.PreviousItem(ref index));
			Assert.AreEqual(14, index);
		}

		[TestMethod]
		public void PreviousItemSkipsMissingItems()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, null, null },
				{ null, null, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 10;
			Assert.IsTrue(pagedStructure.PreviousItem(ref index));
			Assert.AreEqual(5, index);
		}

		[TestMethod]
		public void PreviousItemIndexTooLowFindsNothing()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, null, null },
				{ null, null, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = -3;
			Assert.IsFalse(pagedStructure.PreviousItem(ref index));
			Assert.AreEqual(-4, index);
		}

		[TestMethod]
		public void PreviousItemIndexTooHighFindsLastItem()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ 5, 6, null, null },
				{ null, null, 11, 12 },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			int index = 23;
			Assert.IsTrue(pagedStructure.PreviousItem(ref index));
			Assert.AreEqual(15, index);
		}
		#endregion
		#endregion

		#endregion

		#region Helper Methods
		private int BuildRow(int page, int indexOnPage)
		{
			if (indexOnPage < 1 || indexOnPage > 1024)
				throw new InvalidOperationException("Row pages take indices between 1 and 1024.");
			if (page < 0 || page > 1023)
				throw new InvalidOperationException("Pages are 0-indexed and can be between 0 and 1023.");
			return page * 1024 + indexOnPage;
		}

		private int BuildColumn(int page, int indexOnPage)
		{
			if (indexOnPage < 1 || indexOnPage > 128)
				throw new InvalidOperationException("Column pages take indices between 1 and 128.");
			if (page < 0 || page > 127)
				throw new InvalidOperationException("Pages are 0-indexed and can be between 0 and 127.");
			return page * 128 + indexOnPage;
		}
		#endregion
	}
}
