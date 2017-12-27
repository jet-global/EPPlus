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
			var cellStore = this.GetCellStore();
			Assert.AreEqual(1, cellStore.GetValue(1, 1));
			Assert.AreEqual(103, cellStore.GetValue(7, 7));
			Assert.AreEqual(122, cellStore.GetValue(8, 10));
			Assert.AreEqual(219, cellStore.GetValue(14, 11));
			Assert.AreEqual(256, cellStore.GetValue(16, 16));
			// Non-existent value returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(25, 25));
		}

		[TestMethod]
		public void GetValueReturnsDefaultForInvalidCoordinates()
		{
			var cellStore = this.GetCellStore();
			// Invalid row too small returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(0, 10));
			// Invalid row too large returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(20, 10));
			// Invalid column too small returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(10, 0));
			// Invalid column too large returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(10, 20));
		}
		#endregion

		#region SetValue Tests
		[TestMethod]
		public void SetValue()
		{
			var cellStore = this.GetCellStore(false);
			cellStore.SetValue(1, 1, 1);
			Assert.AreEqual(1, cellStore.GetValue(1, 1));
			cellStore.SetValue(7, 7, 103);
			Assert.AreEqual(103, cellStore.GetValue(7, 7));
			cellStore.SetValue(8, 10, 122);
			Assert.AreEqual(122, cellStore.GetValue(8, 10));
			cellStore.SetValue(14, 11, 219);
			Assert.AreEqual(219, cellStore.GetValue(14, 11));
			cellStore.SetValue(16, 16, 256);
			Assert.AreEqual(256, cellStore.GetValue(16, 16));
		}

		[TestMethod]
		public void SetValueIgnoresInvalidCoordinates()
		{
			var cellStore = this.GetCellStore(false);
			// row too small is ignored
			cellStore.SetValue(0, 1, 13);
			// row too large is ignored
			cellStore.SetValue(20, 1, 13);
			// column too small is ignored
			cellStore.SetValue(1, 0, 13);
			// column too large is ignored
			cellStore.SetValue(1, 20, 13);
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

		#region GetDimension Tests
		[TestMethod]
		public void GetDimensionEmptyCellStore()
		{
			var cellStore = new ZCellStore<int>();
			Assert.IsFalse(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
		}

		[TestMethod]
		public void GetDimensionSingleItem()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(400, 6000, 5);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(400, fromRow);
			Assert.AreEqual(6000, fromCol);
			Assert.AreEqual(400, toRow);
			Assert.AreEqual(6000, toCol);
		}

		[TestMethod]
		public void GetDimensionTwoItems()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(20, 300, 5);
			cellStore.SetValue(400, 6000, 5);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(20, fromRow);
			Assert.AreEqual(300, fromCol);
			Assert.AreEqual(400, toRow);
			Assert.AreEqual(6000, toCol);
		}

		[TestMethod]
		public void GetDimensionFourItems()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(20, 100, 5);
			cellStore.SetValue(15, 300, 5);
			cellStore.SetValue(400, 6000, 5);
			cellStore.SetValue(500, 5000, 5);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(15, fromRow);
			Assert.AreEqual(100, fromCol);
			Assert.AreEqual(500, toRow);
			Assert.AreEqual(6000, toCol);
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
			Assert.AreEqual(ExcelPackage.MaxRows, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(-1, pagedStructure.MaximumUsedIndex);
			Assert.IsTrue(pagedStructure.IsEmpty);
			Assert.AreEqual(-1, pagedStructure.MaximumUsedIndex);
			pagedStructure = new PagedStructure<int>(4);
			Assert.AreEqual(4, pagedStructure.PageBits);
			Assert.AreEqual(16, pagedStructure.PageSize);
			Assert.AreEqual(15, pagedStructure.PageMask);
			Assert.AreEqual(255, pagedStructure.MaximumIndex);
			Assert.AreEqual(256, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(-1, pagedStructure.MaximumUsedIndex);
			Assert.IsTrue(pagedStructure.IsEmpty);
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
			Assert.AreEqual(16, pagedStructure.MinimumUsedIndex);
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
			Assert.AreEqual(16, pagedStructure.MinimumUsedIndex);
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

		#region Nested Class Tests
		#region Page Tests
		#region Indexer Tests
		[TestMethod]
		public void PageIndexerReturnsNullForNonExistentItems()
		{
			var page = new ZCellStore<int>.PagedStructure<int>.Page(10);
			Assert.AreEqual(10, page.MinimumUsedIndex);
			Assert.AreEqual(-1, page.MaximumUsedIndex);
			Assert.IsNull(page[0]);
			Assert.IsNull(page[4]);
			Assert.IsNull(page[9]);
		}

		[TestMethod]
		public void PageIndexerSetsUsedIndices()
		{
			var page = new ZCellStore<int>.PagedStructure<int>.Page(10);
			Assert.AreEqual(10, page.MinimumUsedIndex);
			Assert.AreEqual(-1, page.MaximumUsedIndex);
			Assert.IsTrue(page.IsEmpty);
			Assert.IsNull(page[0]);
			page[0] = 5;
			Assert.AreEqual(0, page.MinimumUsedIndex);
			Assert.AreEqual(0, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.AreEqual(5, page[0]);
			page[7] = 13;
			Assert.AreEqual(0, page.MinimumUsedIndex);
			Assert.AreEqual(7, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.AreEqual(13, page[7]);
			page[0] = null;
			Assert.AreEqual(7, page.MinimumUsedIndex);
			Assert.AreEqual(7, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.IsNull(page[0]);
			page[3] = 13;
			Assert.AreEqual(3, page.MinimumUsedIndex);
			Assert.AreEqual(7, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.AreEqual(13, page[3]);
			page[9] = 20;
			Assert.AreEqual(3, page.MinimumUsedIndex);
			Assert.AreEqual(9, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.AreEqual(20, page[9]);
			page[9] = null;
			Assert.AreEqual(3, page.MinimumUsedIndex);
			Assert.AreEqual(7, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.IsNull(page[9]);
			page[7] = null;
			Assert.AreEqual(3, page.MinimumUsedIndex);
			Assert.AreEqual(3, page.MaximumUsedIndex);
			Assert.IsFalse(page.IsEmpty);
			Assert.IsNull(page[7]);
			page[3] = null;
			Assert.AreEqual(10, page.MinimumUsedIndex);
			Assert.AreEqual(-1, page.MaximumUsedIndex);
			Assert.IsTrue(page.IsEmpty);
			Assert.IsNull(page[3]);
		}
		#endregion

		#region TryGetNextIndex Tests
		[TestMethod]
		public void TryGetNextIndex()
		{
			var page = new ZCellStore<int>.PagedStructure<int>.Page(10);
			Assert.IsFalse(page.TryGetNextIndex(3, out int nextIndex));
			page[4] = 45;
			Assert.IsTrue(page.TryGetNextIndex(3, out nextIndex));
			Assert.AreEqual(4, nextIndex);
			page[9] = 13;
			Assert.IsTrue(page.TryGetNextIndex(4, out nextIndex));
			Assert.AreEqual(9, nextIndex);
			Assert.IsFalse(page.TryGetNextIndex(9, out nextIndex));
		}
		#endregion

		#region TryGetPreviousIndex Tests
		[TestMethod]
		public void TryGetPreviousIndex()
		{
			var page = new ZCellStore<int>.PagedStructure<int>.Page(10);
			Assert.IsFalse(page.TryGetPreviousIndex(7, out int previousIndex));
			page[4] = 45;
			Assert.IsTrue(page.TryGetPreviousIndex(7, out previousIndex));
			Assert.AreEqual(4, previousIndex);
			page[2] = 13;
			Assert.IsTrue(page.TryGetPreviousIndex(4, out previousIndex));
			Assert.AreEqual(2, previousIndex);
			Assert.IsFalse(page.TryGetPreviousIndex(0, out previousIndex));
		}
		#endregion
		#endregion
		#endregion

		#endregion

		#region ZCellStoreEnumerator Tests
		[TestMethod]
		public void ZCellStoreEnumeratorEnumerateEmptySet()
		{
			var cellStore = new ZCellStore<int>();
			var enumerator = cellStore.GetEnumerator();
			Assert.AreEqual(1, enumerator.Row);
			Assert.AreEqual(0, enumerator.Column);
			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesAllValues()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator();
			for (value = 1; value <= 100; value++)
			{
				Assert.IsTrue(enumerator.MoveNext());
				Assert.AreEqual(value, enumerator.Value);
			}
			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockTopRows()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(1, 1, 2, 10);
			for (value = 1; value <= 20; value++)
			{
				Assert.IsTrue(enumerator.MoveNext());
				Assert.AreEqual(value, enumerator.Value);
			}
			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockBottomRows()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(7, 1, 10, 10);
			for (value = 61; value <= 100; value++)
			{
				Assert.IsTrue(enumerator.MoveNext());
				Assert.AreEqual(value, enumerator.Value);
			}
			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockBottomRowsHandlesLargeBlock()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(7, 1, 100, 100);
			for (value = 61; value <= 100; value++)
			{
				Assert.IsTrue(enumerator.MoveNext());
				Assert.AreEqual(value, enumerator.Value);
			}
			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockLeftColumns()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(1, 1, 10, 2);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(1, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(2, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(11, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(12, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(21, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(22, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(31, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(32, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(41, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(42, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(51, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(52, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(61, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(62, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(71, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(72, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(81, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(82, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(91, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(92, enumerator.Value);

			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockRightColumns()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(1, 7, 10, 10);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(7, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(8, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(9, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(10, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(17, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(18, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(19, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(20, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(27, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(28, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(29, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(30, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(37, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(38, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(39, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(40, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(47, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(48, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(49, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(50, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(57, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(58, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(59, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(60, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(67, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(68, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(69, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(70, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(77, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(78, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(79, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(80, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(87, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(88, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(89, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(90, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(97, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(98, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(99, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(100, enumerator.Value);

			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockRightColumnsHandlesLargeBlock()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(1, 7, 100, 100);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(7, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(8, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(9, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(10, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(17, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(18, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(19, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(20, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(27, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(28, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(29, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(30, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(37, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(38, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(39, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(40, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(47, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(48, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(49, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(50, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(57, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(58, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(59, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(60, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(67, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(68, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(69, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(70, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(77, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(78, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(79, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(80, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(87, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(88, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(89, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(90, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(97, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(98, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(99, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(100, enumerator.Value);

			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesInternalBlock()
		{
			var cellStore = new ZCellStore<int>();
			int value = 1;
			for (int row = 1; row <= 10; row++)
			{
				for (int column = 1; column <= 10; column++)
				{
					cellStore.SetValue(row, column, value++);
				}
			}
			var enumerator = cellStore.GetEnumerator(4, 4, 7, 7);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(34, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(35, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(36, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(37, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(44, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(45, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(46, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(47, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(54, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(55, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(56, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(57, enumerator.Value);

			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(64, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(65, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(66, enumerator.Value);
			Assert.IsTrue(enumerator.MoveNext());
			Assert.AreEqual(67, enumerator.Value);

			Assert.IsFalse(enumerator.MoveNext());
		}
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

		private ZCellStore<int> GetCellStore(bool fill = true)
		{
			var cellStore = new ZCellStore<int>(4, 4);
			if (fill)
			{
				var currentStore = new int[,]
				{
/*1*/			{  1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16},
/*2*/			{ 17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32},
/*3*/			{ 33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48},
/*4*/			{ 49,  50,  51,  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64},
/*5*/			{ 65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80},
/*6*/			{ 81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96},
/*7*/			{ 97,  98,  99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112},
/*8*/			{113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128},
/*9*/			{129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144},
/*10*/		{145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160},
/*11*/		{161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176},
/*12*/		{177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192},
/*13*/		{193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208},
/*14*/		{209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224},
/*15*/		{225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240},
/*16*/		{241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256}
				};
				this.LoadCellStore(currentStore, cellStore);
			}
			return cellStore;
		}

		private void LoadCellStore(int[,] sheet, ZCellStore<int> cellStore)
		{
			for (int row = 0; row <= sheet.GetUpperBound(0); ++row)
			{
				for (int column = 0; column <= sheet.GetUpperBound(1); ++column)
				{
					cellStore.SetValue(row + 1, column + 1, sheet[row, column]);
				}
			}
		}

		private void ValidateCellStore(int[,] sheet, ZCellStore<int> cellStore)
		{
			for (int row = 0; row <= sheet.GetUpperBound(0); ++row)
			{
				for (int column = 0; column <= sheet.GetUpperBound(1); ++column)
				{
					var data = sheet[row, column];
					var item = cellStore.GetValue(row + 1, column + 1);
					Assert.AreEqual(data, item);
				}
			}
		}
		#endregion

		[TestMethod]
		public void DELME()
		{
			var ogCellStore = new CellStore<int>();
			ogCellStore.SetValue(0, 1, 3);

		}
	}
}
