using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using static OfficeOpenXml.ZCellStore<int>;

namespace EPPlusTest
{
	[TestClass]
	public class ZCellStoreTest
	{
		/*
		 * it is valid to use 0 for column and row coordinates.
		 * These values are used for column and row metadata.
		 * 
		 * [0,0] is NOT considered valid as far as I can tell.
		 * 
		 * 
		 * Assume the current invariant:
		 * X : Invalid coordinate.
		 * C : Column metadata.
		 * R : Row metadata.
		 * * : Cell content.
		 * 
		 *          XCCCCCCCC
		 *          R********
		 *          R********
		 *          R********
		 *          R********
		 *          R********
		 *          R********
		 *          R********
		 *          R********
		 * 
		 * */

		#region GetValue Tests
		[TestMethod]
		public void GetValue()
		{
			var cellStore = this.GetCellStore();
			Assert.AreEqual(1001, cellStore.GetValue(0, 1));
			Assert.AreEqual(2001, cellStore.GetValue(1, 0));
			Assert.AreEqual(1, cellStore.GetValue(1, 1));
			Assert.AreEqual(103, cellStore.GetValue(7, 7));
			Assert.AreEqual(122, cellStore.GetValue(8, 10));
			Assert.AreEqual(219, cellStore.GetValue(14, 11));
			Assert.AreEqual(256, cellStore.GetValue(16, 16));
		}

		[TestMethod]
		public void GetValueReturnsDefaultForInvalidCoordinates()
		{
			var cellStore = this.GetCellStore();
			// [0,0] is not a real address
			Assert.AreEqual(0, cellStore.GetValue(0, 0));
			// Invalid row too small returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(-1, 10));
			// Invalid row too large returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(20, 10));
			// Invalid column too small returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(10, -1));
			// Invalid column too large returns default(T)
			Assert.AreEqual(0, cellStore.GetValue(10, 20));
		}
		#endregion

		#region SetValue Tests
		[TestMethod]
		public void SetValue()
		{
			var cellStore = this.GetCellStore(false);
			cellStore.SetValue(0, 1, 13);
			Assert.AreEqual(13, cellStore.GetValue(0, 1));
			cellStore.SetValue(1, 0, 14);
			Assert.AreEqual(14, cellStore.GetValue(1, 0));
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
			// [0,0] is not a real address
			cellStore.SetValue(0, 0, 13);
			// row too small is ignored
			cellStore.SetValue(-1, 1, 13);
			// row too large is ignored
			cellStore.SetValue(20, 1, 13);
			// column too small is ignored
			cellStore.SetValue(1, -1, 13);
			// column too large is ignored
			cellStore.SetValue(1, 20, 13);
		}
		#endregion

		#region Exists Tests
		[TestMethod]
		public void Exists()
		{
			var cellStore = new ZCellStore<int>();
			Assert.IsFalse(cellStore.Exists(0, 1));
			cellStore.SetValue(0, 1, 6);
			Assert.IsTrue(cellStore.Exists(0, 1));
			Assert.IsFalse(cellStore.Exists(1, 0));
			cellStore.SetValue(1, 0, 7);
			Assert.IsTrue(cellStore.Exists(1, 0));
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
			// [0,0] is not a real address
			Assert.IsFalse(cellStore.Exists(0, 0));
			// row too small is false
			Assert.IsFalse(cellStore.Exists(-1, 1));
			// row too large is false
			Assert.IsFalse(cellStore.Exists(ExcelPackage.MaxRows + 1, 1));
			// column too small is false
			Assert.IsFalse(cellStore.Exists(1, -1));
			// column too large is false
			Assert.IsFalse(cellStore.Exists(1, ExcelPackage.MaxColumns + 1));
		}

		[TestMethod]
		public void ExistsWithOutValue()
		{
			var cellStore = new ZCellStore<int>();
			Assert.IsFalse(cellStore.Exists(0, 1, out int value));
			cellStore.SetValue(0, 1, 6);
			Assert.IsTrue(cellStore.Exists(0, 1, out value));
			Assert.AreEqual(6, value);
			Assert.IsFalse(cellStore.Exists(1, 0, out value));
			cellStore.SetValue(1, 0, 7);
			Assert.IsTrue(cellStore.Exists(1, 0, out value));
			Assert.AreEqual(7, value);
			Assert.IsFalse(cellStore.Exists(3, 3, out value));
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
			// [0,0] is not a real address
			Assert.IsFalse(cellStore.Exists(0, 0, out int value));
			// row too small is false
			Assert.IsFalse(cellStore.Exists(-1, 1, out value));
			// row too large is false
			Assert.IsFalse(cellStore.Exists(ExcelPackage.MaxRows + 1, 1, out value));
			// column too small is false
			Assert.IsFalse(cellStore.Exists(1, -1, out value));
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
		public void NextCellStartsWithZeroHitsColumnMetaData()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(0, 1, 1);
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(0, row);
			Assert.AreEqual(1, column);
		}

		[TestMethod]
		public void NextCellStartsWithZeroHitsRowMetaData()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 0, 1);
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(0, column);
		}

		[TestMethod]
		public void NextCellStartsWithZeroHitsCellData()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			int row = 0, column = 0;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
		}

		[TestMethod]
		public void NextCellNormalizesStartSearchIndices()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 1, 1);
			int row = -111, column = -111;
			Assert.IsTrue(cellStore.NextCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(1, column);
		}

		[TestMethod]
		public void NextCellEnumeratesFullSheet()
		{
			var cellStore = this.GetCellStore();
			int row = 0, column = 0, value = 0;
			int rowMetadata = 2001;
			int columnMetadata = 1001;
			while (cellStore.NextCell(ref row, ref column))
			{
				if (row == 0)
				{
					Assert.AreEqual(columnMetadata % 100, column);
					Assert.AreEqual(columnMetadata++, cellStore.GetValue(row, column));
				}
				else if (column == 0)
				{
					Assert.AreEqual(rowMetadata % 100, row);
					Assert.AreEqual(rowMetadata++, cellStore.GetValue(row, column));
				}
				else
				{
					value++;
					Assert.AreEqual(value, cellStore.GetValue(row, column));
					Assert.AreEqual(((value - 1) / 16) + 1, row);
					Assert.AreEqual(((value - 1) % 16) + 1, column);
				}
			}
			Assert.AreEqual(256, value);
			Assert.AreEqual(2017, rowMetadata);
			Assert.AreEqual(1017, columnMetadata);
		}

		[TestMethod]
		public void NextCellEnumeratesDiagonal()
		{
			var cellStore = this.GetCellStore(false);
			var currentStore = new int?[,]
			{
/*1*/		{    1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null,    2, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null,    3, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null,    4, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null,    5, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null,    6, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null,    7, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null,    8, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null,    9, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null,   10, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null,   11, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null,   12, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null,   13, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null,   14, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null,   15, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null,   16 }
			};
			this.LoadCellStore(null, null, currentStore, cellStore);
			int row = 0, column = 0, value = 0;
			while (cellStore.NextCell(ref row, ref column))
			{
				value++;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(value, row);
				Assert.AreEqual(value, column);
			}
			Assert.AreEqual(16, value);
		}

		[TestMethod]
		public void NextCellEnumeratesColumn()
		{
			var cellStore = this.GetCellStore(false);
			var currentStore = new int?[,]
			{
/*1*/		{ null, null, null, null,  1, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null,  2, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null,  3, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null,  4, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null,  5, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null,  6, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null,  7, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null,  8, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null,  9, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, 10, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, 11, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, 12, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, 13, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, 14, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, 15, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, 16, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, currentStore, cellStore);
			int row = 0, column = 0, value = 0;
			while (cellStore.NextCell(ref row, ref column))
			{
				value++;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(value, row);
				Assert.AreEqual(5, column);
			}
			Assert.AreEqual(16, value);
		}

		[TestMethod]
		public void NextCellEnumeratesColumnWithMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var currentSheetData = new int?[,]
			{
/*1*/		{ null, null, null, null,  1, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null,  2, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null,  3, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null,  4, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null,  5, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null,  6, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null,  7, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null,  8, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null,  9, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, 10, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, 11, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, 12, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, 13, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, 14, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, 15, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, 16, null, null, null, null, null, null, null, null, null, null, null }
			};
			var currentColumnData = new int?[] { null, null, null, null, 1005, null, null, null, null, null, null, null, null, null, null, null };
			this.LoadCellStore(currentColumnData, null, currentSheetData, cellStore);
			int row = 0, column = 0, value = 0;
			bool hitMetadata = false;
			while (cellStore.NextCell(ref row, ref column))
			{
				if (row == 0)
				{
					Assert.AreEqual(1005, cellStore.GetValue(row, column));
					Assert.AreEqual(5, column);
					hitMetadata = true;
				}
				else
				{
					value++;
					Assert.AreEqual(value, cellStore.GetValue(row, column));
					Assert.AreEqual(value, row);
					Assert.AreEqual(5, column);
				}
			}
			Assert.AreEqual(16, value);
			Assert.IsTrue(hitMetadata);
		}

		[TestMethod]
		public void NextCellEnumeratesRow()
		{
			var cellStore = this.GetCellStore(false);
			var currentStore = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, currentStore, cellStore);
			int row = 0, column = 0, value = 0;
			while (cellStore.NextCell(ref row, ref column))
			{
				value++;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(9, row);
				Assert.AreEqual(value, column);
			}
			Assert.AreEqual(16, value);
		}

		[TestMethod]
		public void NextCellEnumeratesRowWithMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var currentSheetData = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			var currentRowData = new int?[] { null, null, null, null, null, null, null, null, 2009, null, null, null, null, null, null, null };
			this.LoadCellStore(null, currentRowData, currentSheetData, cellStore);
			int row = 0, column = 0, value = 0;
			bool hitMetadata = false;
			while (cellStore.NextCell(ref row, ref column))
			{
				if (column == 0)
				{
					Assert.AreEqual(2009, cellStore.GetValue(row, column));
					Assert.AreEqual(9, row);
					hitMetadata = true;
				}
				else
				{
					value++;
					Assert.AreEqual(value, cellStore.GetValue(row, column));
					Assert.AreEqual(9, row);
					Assert.AreEqual(value, column);
				}
			}
			Assert.AreEqual(16, value);
			Assert.IsTrue(hitMetadata);
		}

		[TestMethod]
		public void NextCellEnumeratesColumnMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var columnMetadata = new int?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };
			this.LoadCellStore(columnMetadata, null, null, cellStore);
			int row = 0, column = 0, value = 0;
			while (cellStore.NextCell(ref row, ref column))
			{
				value++;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(0, row);
				Assert.AreEqual(value, column);
			}
			Assert.AreEqual(16, value);
		}

		[TestMethod]
		public void NextCellEnumeratesRowMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var rowMetadata = new int?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };
			this.LoadCellStore(null, rowMetadata, null, cellStore);
			int row = 0, column = 0, value = 0;
			while (cellStore.NextCell(ref row, ref column))
			{
				value++;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(value, row);
				Assert.AreEqual(0, column);
			}
			Assert.AreEqual(16, value);
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
		public void PrevCellStartsWithZeroHitsColumnMetaData()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(0, 1, 1);
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(0, row);
			Assert.AreEqual(1, column);
		}

		[TestMethod]
		public void PrevCellStartsWithZeroHitsRowMetaData()
		{
			var cellStore = new ZCellStore<int>();
			cellStore.SetValue(1, 0, 1);
			int row = ExcelPackage.MaxRows + 1, column = ExcelPackage.MaxColumns + 1;
			Assert.IsTrue(cellStore.PrevCell(ref row, ref column));
			Assert.AreEqual(1, row);
			Assert.AreEqual(0, column);
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
		public void PrevCellEnumeratesFullSheet()
		{
			var cellStore = this.GetCellStore();
			int row = 17, column = 17, value = 257;
			int rowMetadata = 2016;
			int columnMetadata = 1016;
			while (cellStore.PrevCell(ref row, ref column))
			{
				if (row == 0)
				{
					Assert.AreEqual(columnMetadata % 100, column);
					Assert.AreEqual(columnMetadata--, cellStore.GetValue(row, column));
				}
				else if (column == 0)
				{
					Assert.AreEqual(rowMetadata % 100, row);
					Assert.AreEqual(rowMetadata--, cellStore.GetValue(row, column));
				}
				else
				{
					value--;
					Assert.AreEqual(value, cellStore.GetValue(row, column));
					Assert.AreEqual(((value - 1) / 16) + 1, row);
					Assert.AreEqual(((value - 1) % 16) + 1, column);
				}
			}
			Assert.AreEqual(1, value);
			Assert.AreEqual(2000, rowMetadata);
			Assert.AreEqual(1000, columnMetadata);
		}

		[TestMethod]
		public void PrevCellEnumeratesDiagonal()
		{
			var cellStore = this.GetCellStore(false);
			var currentStore = new int?[,]
			{
/*1*/		{    1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null,    2, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null,    3, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null,    4, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null,    5, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null,    6, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null,    7, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null,    8, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null,    9, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null,   10, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null,   11, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null,   12, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null,   13, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null,   14, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null,   15, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null,   16 }
			};
			this.LoadCellStore(null, null, currentStore, cellStore);
			int row = 17, column = 17, value = 17;
			while (cellStore.PrevCell(ref row, ref column))
			{
				value--;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(value, row);
				Assert.AreEqual(value, column);
			}
			Assert.AreEqual(1, value);
		}

		[TestMethod]
		public void PrevCellEnumeratesColumn()
		{
			var cellStore = this.GetCellStore(false);
			var currentStore = new int?[,]
			{
/*1*/		{ null, null, null, null,  1, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null,  2, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null,  3, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null,  4, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null,  5, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null,  6, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null,  7, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null,  8, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null,  9, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, 10, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, 11, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, 12, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, 13, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, 14, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, 15, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, 16, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, currentStore, cellStore);
			int row = 17, column = 17, value = 17;
			while (cellStore.PrevCell(ref row, ref column))
			{
				value--;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(value, row);
				Assert.AreEqual(5, column);
			}
			Assert.AreEqual(1, value);
		}

		[TestMethod]
		public void PrevCellEnumeratesColumnWithMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var currentSheetData = new int?[,]
			{
/*1*/		{ null, null, null, null,  1, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null,  2, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null,  3, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null,  4, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null,  5, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null,  6, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null,  7, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null,  8, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null,  9, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, 10, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, 11, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, 12, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, 13, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, 14, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, 15, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, 16, null, null, null, null, null, null, null, null, null, null, null }
			};
			var currentColumnData = new int?[] { null, null, null, null, 1005, null, null, null, null, null, null, null, null, null, null, null };
			this.LoadCellStore(currentColumnData, null, currentSheetData, cellStore);
			int row = 17, column = 17, value = 17;
			bool hitMetadata = false;
			while (cellStore.PrevCell(ref row, ref column))
			{
				if (row == 0)
				{
					Assert.AreEqual(1005, cellStore.GetValue(row, column));
					Assert.AreEqual(5, column);
					hitMetadata = true;
				}
				else
				{

					value--;
					Assert.AreEqual(value, cellStore.GetValue(row, column));
					Assert.AreEqual(value, row);
					Assert.AreEqual(5, column);
				}
			}
			Assert.AreEqual(1, value);
			Assert.IsTrue(hitMetadata);
		}

		[TestMethod]
		public void PrevCellEnumeratesRow()
		{
			var cellStore = this.GetCellStore(false);
			var currentStore = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, currentStore, cellStore);
			int row = 17, column = 17, value = 17;
			while (cellStore.PrevCell(ref row, ref column))
			{
				value--;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(9, row);
				Assert.AreEqual(value, column);
			}
			Assert.AreEqual(1, value);
		}

		[TestMethod]
		public void PrevCellEnumeratesRowWithMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var currentSheetData = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			var currentRowData = new int?[] { null, null, null, null, null, null, null, null, 2009, null, null, null, null, null, null, null };
			this.LoadCellStore(null, currentRowData, currentSheetData, cellStore);
			int row = 17, column = 17, value = 17;
			bool hitMetadata = false;
			while (cellStore.PrevCell(ref row, ref column))
			{
				if (column == 0)
				{
					Assert.AreEqual(2009, cellStore.GetValue(row, column));
					Assert.AreEqual(9, row);
					hitMetadata = true;
				}
				else
				{
					value--;
					Assert.AreEqual(value, cellStore.GetValue(row, column));
					Assert.AreEqual(9, row);
					Assert.AreEqual(value, column);
				}
			}
			Assert.AreEqual(1, value);
			Assert.IsTrue(hitMetadata);
		}

		[TestMethod]
		public void PrevCellEnumeratesColumnMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var columnMetadata = new int?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };
			this.LoadCellStore(columnMetadata, null, null, cellStore);
			int row = 17, column = 17, value = 17;
			while (cellStore.PrevCell(ref row, ref column))
			{
				value--;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(0, row);
				Assert.AreEqual(value, column);
			}
			Assert.AreEqual(1, value);
		}

		[TestMethod]
		public void PrevCellEnumeratesRowMetadata()
		{
			var cellStore = this.GetCellStore(false);
			var rowMetadata = new int?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };
			this.LoadCellStore(null, rowMetadata, null, cellStore);
			int row = 17, column = 17, value = 17;
			while (cellStore.PrevCell(ref row, ref column))
			{
				value--;
				Assert.AreEqual(value, cellStore.GetValue(row, column));
				Assert.AreEqual(value, row);
				Assert.AreEqual(0, column);
			}
			Assert.AreEqual(1, value);
		}
		#endregion

		#region Delete Tests
		[TestMethod]
		public void DeleteRowsAcrossAllColumns()
		{
			var cellStore = this.GetCellStore();
			cellStore.Delete(2, 0, 5, 0);
			var expectedSheetData = new int?[,]
			{
/*1*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/		{   97,   98,   99,  100,  101,  102,  103,  104,  105,  106,  107,  108,  109,  110,  111,  112 },
/*3*/		{  113,  114,  115,  116,  117,  118,  119,  120,  121,  122,  123,  124,  125,  126,  127,  128 },
/*4*/		{  129,  130,  131,  132,  133,  134,  135,  136,  137,  138,  139,  140,  141,  142,  143,  144 },
/*5*/		{  145,  146,  147,  148,  149,  150,  151,  152,  153,  154,  155,  156,  157,  158,  159,  160 },
/*6*/		{  161,  162,  163,  164,  165,  166,  167,  168,  169,  170,  171,  172,  173,  174,  175,  176 },
/*7*/		{  177,  178,  179,  180,  181,  182,  183,  184,  185,  186,  187,  188,  189,  190,  191,  192 },
/*8*/		{  193,  194,  195,  196,  197,  198,  199,  200,  201,  202,  203,  204,  205,  206,  207,  208 },
/*9*/		{  209,  210,  211,  212,  213,  214,  215,  216,  217,  218,  219,  220,  221,  222,  223,  224 },
/*10*/	{  225,  226,  227,  228,  229,  230,  231,  232,  233,  234,  235,  236,  237,  238,  239,  240 },
/*11*/	{  241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  256 },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, null, null, null, null, null };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedSheetData, cellStore);
		}
		
		[TestMethod]
		public void DeleteColumnsAcrossAllRows()
		{
			var cellStore = this.GetCellStore();
			cellStore.Delete(0, 2, 0, 5);
			var expectedSheetData = new int?[,]
			{
/*1*/		{   1,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16, null, null, null, null, null },
/*2*/		{  17,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32, null, null, null, null, null },
/*3*/		{  33,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48, null, null, null, null, null },
/*4*/		{  49,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64, null, null, null, null, null },
/*5*/		{  65,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80, null, null, null, null, null },
/*6*/		{  81,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96, null, null, null, null, null },
/*7*/		{  97, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, null, null, null, null, null },
/*8*/		{ 113, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, null, null, null, null, null },
/*9*/		{ 129, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, null, null, null, null, null },
/*10*/	{ 145, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, null, null, null, null, null },
/*11*/	{ 161, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, null, null, null, null, null },
/*12*/	{ 177, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, null, null, null, null, null },
/*13*/	{ 193, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, null, null, null, null, null },
/*14*/	{ 209, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, null, null, null, null, null },
/*15*/	{ 225, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, null, null, null, null, null },
/*16*/	{ 241, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, null, null, null, null, null }
			};
			var expectedColumnData = new int?[] { 1001, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016, null, null, null, null, null };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedSheetData, cellStore);
		}
		
		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void DeleteAllButFirstAndLastRowsAndColumns()
		{
			var cellStore = this.GetCellStore();
			cellStore.Delete(2, 2, 14, 14);
			#region Expected Result for if we decide to implement this.
//			var expectedStore = new int?[,]
//			{
///*1*/		{   1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,  16 },
///*2*/		{  17,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  32 },
///*3*/		{  33, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  48 },
///*4*/		{  49, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  64 },
///*5*/		{  65, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  80 },
///*6*/		{  81, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  96 },
///*7*/		{  97, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 112 },
///*8*/		{ 113, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 128 },
///*9*/		{ 129, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 144 },
///*10*/	{ 145, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 160 },
///*11*/	{ 161, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 176 },
///*12*/	{ 177, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 192 },
///*13*/	{ 193, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 208 },
///*14*/	{ 209, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 224 },
///*15*/	{ 225, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 240 },
///*16*/	{ 241, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 256 }
//			};
//			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
//			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
//			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
			#endregion
		}
		#endregion

		#region Clear Tests
		[TestMethod]
		public void ClearRows()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(2, 0, 5, ExcelPackage.MaxColumns);
			var expectedStore = new int?[,]
			{
/*1*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{   97,   98,   99,  100,  101,  102,  103,  104,  105,  106,  107,  108,  109,  110,  111,  112 },
/*8*/		{  113,  114,  115,  116,  117,  118,  119,  120,  121,  122,  123,  124,  125,  126,  127,  128 },
/*9*/		{  129,  130,  131,  132,  133,  134,  135,  136,  137,  138,  139,  140,  141,  142,  143,  144 },
/*10*/	{  145,  146,  147,  148,  149,  150,  151,  152,  153,  154,  155,  156,  157,  158,  159,  160 },
/*11*/	{  161,  162,  163,  164,  165,  166,  167,  168,  169,  170,  171,  172,  173,  174,  175,  176 },
/*12*/	{  177,  178,  179,  180,  181,  182,  183,  184,  185,  186,  187,  188,  189,  190,  191,  192 },
/*13*/	{  193,  194,  195,  196,  197,  198,  199,  200,  201,  202,  203,  204,  205,  206,  207,  208 },
/*14*/	{  209,  210,  211,  212,  213,  214,  215,  216,  217,  218,  219,  220,  221,  222,  223,  224 },
/*15*/	{  225,  226,  227,  228,  229,  230,  231,  232,  233,  234,  235,  236,  237,  238,  239,  240 },
/*16*/	{  241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  256 }
			};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, null, null, null, null, null, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearColumns()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(0, 2, ExcelPackage.MaxRows, 5);
			var expectedStore = new int?[,]
			{
/*1*/		{   1, null, null, null, null, null,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/		{  17, null, null, null, null, null,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/		{  33, null, null, null, null, null,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/		{  49, null, null, null, null, null,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/		{  65, null, null, null, null, null,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/		{  81, null, null, null, null, null,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/		{  97, null, null, null, null, null, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/		{ 113, null, null, null, null, null, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/		{ 129, null, null, null, null, null, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/	{ 145, null, null, null, null, null, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/	{ 161, null, null, null, null, null, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/	{ 177, null, null, null, null, null, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/	{ 193, null, null, null, null, null, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/	{ 209, null, null, null, null, null, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/	{ 225, null, null, null, null, null, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/	{ 241, null, null, null, null, null, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
			};
			var expectedColumnData = new int?[] { 1001, null, null, null, null, null, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearRowsAndColumns()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(2, 2, 5, 5);
			var expectedStore = new int?[,]
			{
/*1*/		{   1,    2,    3,    4,    5,    6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/		{  17, null, null, null, null, null,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/		{  33, null, null, null, null, null,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/		{  49, null, null, null, null, null,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/		{  65, null, null, null, null, null,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/		{  81, null, null, null, null, null,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/		{  97,   98,   99,  100,  101,  102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/		{ 113,  114,  115,  116,  117,  118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/		{ 129,  130,  131,  132,  133,  134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/	{ 145,  146,  147,  148,  149,  150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/	{ 161,  162,  163,  164,  165,  166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/	{ 177,  178,  179,  180,  181,  182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/	{ 193,  194,  195,  196,  197,  198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/	{ 209,  210,  211,  212,  213,  214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/	{ 225,  226,  227,  228,  229,  230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/	{ 241,  242,  243,  244,  245,  246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
			};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearAllButFirstAndLastRowsAndColumns()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(2, 2, 14, 14);
			var expectedStore = new int?[,]
			{
/*1*/		{   1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,  16 },
/*2*/		{  17, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  32 },
/*3*/		{  33, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  48 },
/*4*/		{  49, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  64 },
/*5*/		{  65, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  80 },
/*6*/		{  81, null, null, null, null, null, null, null, null, null, null, null, null, null, null,  96 },
/*7*/		{  97, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 112 },
/*8*/		{ 113, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 128 },
/*9*/		{ 129, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 144 },
/*10*/	{ 145, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 160 },
/*11*/	{ 161, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 176 },
/*12*/	{ 177, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 192 },
/*13*/	{ 193, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 208 },
/*14*/	{ 209, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 224 },
/*15*/	{ 225, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 240 },
/*16*/	{ 241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255, 256 }
			};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearPartialColumnMetadata()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(0, 4, 0, 5);
			var expectedStore = new int?[,]
				{
/*1*/			{   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/			{  17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/			{  33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/			{  49,  50,  51,  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/			{  65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/			{  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/			{  97,  98,  99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/			{ 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/			{ 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/		{ 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/		{ 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/		{ 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/		{ 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/		{ 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/		{ 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/		{ 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
				};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, null, null, null, null, null, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearAllColumnMetadata()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(0, 1, 0, ExcelPackage.MaxColumns);
			var expectedStore = new int?[,]
				{
/*1*/			{   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/			{  17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/			{  33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/			{  49,  50,  51,  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/			{  65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/			{  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/			{  97,  98,  99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/			{ 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/			{ 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/		{ 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/		{ 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/		{ 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/		{ 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/		{ 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/		{ 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/		{ 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
				};
			var expectedColumnData = new int?[] { null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearPartialRowMetadata()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(4, 0, 5, 0);
			var expectedStore = new int?[,]
				{
/*1*/			{   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/			{  17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/			{  33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/			{  49,  50,  51,  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/			{  65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/			{  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/			{  97,  98,  99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/			{ 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/			{ 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/		{ 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/		{ 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/		{ 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/		{ 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/		{ 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/		{ 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/		{ 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
				};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, 2002, 2003, null, null, null, null, null, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
		}

		[TestMethod]
		public void ClearAllRowMetadata()
		{
			var cellStore = this.GetCellStore();
			cellStore.Clear(1, 0, ExcelPackage.MaxRows, 0);
			var expectedStore = new int?[,]
				{
/*1*/			{   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/			{  17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/			{  33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/			{  49,  50,  51,  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/			{  65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/			{  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/			{  97,  98,  99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/			{ 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/			{ 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/		{ 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/		{ 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/		{ 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/		{ 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/		{ 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/		{ 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/		{ 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
				};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedStore, cellStore);
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
		public void GetDimensionFullCellStoreIgnoresRowAndColumnMetadata()
		{
			var cellStore = this.GetCellStore();
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(1, fromRow);
			Assert.AreEqual(1, fromCol);
			Assert.AreEqual(16, toRow);
			Assert.AreEqual(16, toCol);
		}

		[TestMethod]
		public void GetDimensionSingleItem()
		{
			var cellStore = this.GetCellStore(false);
			var contents = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null,   13, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, contents, cellStore);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(6, fromRow);
			Assert.AreEqual(6, fromCol);
			Assert.AreEqual(6, toRow);
			Assert.AreEqual(6, toCol);
		}

		[TestMethod]
		public void GetDimensionFirstColumn()
		{
			var cellStore = this.GetCellStore(false);
			var contents = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{    1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{    2, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{    3, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, contents, cellStore);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(2, fromRow);
			Assert.AreEqual(1, fromCol);
			Assert.AreEqual(9, toRow);
			Assert.AreEqual(1, toCol);
		}

		[TestMethod]
		public void GetDimensionFirstRow()
		{
			var cellStore = this.GetCellStore(false);
			var contents = new int?[,]
			{
/*1*/		{    1, null,    2, null, null, null, null,    3, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, contents, cellStore);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(1, fromRow);
			Assert.AreEqual(1, fromCol);
			Assert.AreEqual(1, toRow);
			Assert.AreEqual(8, toCol);
		}

		[TestMethod]
		public void GetDimensionTwoItems()
		{
			var cellStore = this.GetCellStore(false);
			var contents = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null,   13, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null,   14, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, contents, cellStore);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(6, fromRow);
			Assert.AreEqual(6, fromCol);
			Assert.AreEqual(12, toRow);
			Assert.AreEqual(10, toCol);
		}

		[TestMethod]
		public void GetDimensionThreeItems()
		{
			var cellStore = this.GetCellStore(false);
			var contents = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null,   15, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null,   13, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null,   14, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, contents, cellStore);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(2, fromRow);
			Assert.AreEqual(6, fromCol);
			Assert.AreEqual(12, toRow);
			Assert.AreEqual(12, toCol);
		}

		[TestMethod]
		public void GetDimensionFourItems()
		{
			var cellStore = this.GetCellStore(false);
			var contents = new int?[,]
			{
/*1*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null,   15, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null,   13, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null,   16, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null,   14, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			this.LoadCellStore(null, null, contents, cellStore);
			Assert.IsTrue(cellStore.GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol));
			Assert.AreEqual(2, fromRow);
			Assert.AreEqual(6, fromCol);
			Assert.AreEqual(12, toRow);
			Assert.AreEqual(15, toCol);
		}
		#endregion

		#region Insert Tests
		[TestMethod]
		public void InsertRowsAcrossAllColumns()
		{
			var cellStore = this.GetCellStore(false);
			var originalSheetData = new int?[,]
			{
/*1*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/		{   17,   18,   19,   20,   21,   22,   23,   24,   25,   26,   27,   28,   29,   30,   31,   32 },
/*3*/		{   33,   34,   35,   36,   37,   38,   39,   40,   41,   42,   43,   44,   45,   46,   47,   48 },
/*4*/		{   49,   50,   51,   52,   53,   54,   55,   56,   57,   58,   59,   60,   61,   62,   63,   64 },
/*5*/		{   65,   66,   67,   68,   69,   70,   71,   72,   73,   74,   75,   76,   77,   78,   79,   80 },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			var originalColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var originalRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, null, null, null, null, null, null, null, null, null, null, null };
			this.LoadCellStore(originalColumnData, originalRowData, originalSheetData, cellStore);
			cellStore.Insert(2, 0, 5, cellStore.MaximumColumn);
			var expectedSheetData = new int?[,]
			{
/*1*/		{    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*3*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*4*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*5*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*6*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*7*/		{   17,   18,   19,   20,   21,   22,   23,   24,   25,   26,   27,   28,   29,   30,   31,   32 },
/*8*/		{   33,   34,   35,   36,   37,   38,   39,   40,   41,   42,   43,   44,   45,   46,   47,   48 },
/*9*/		{   49,   50,   51,   52,   53,   54,   55,   56,   57,   58,   59,   60,   61,   62,   63,   64 },
/*10*/	{   65,   66,   67,   68,   69,   70,   71,   72,   73,   74,   75,   76,   77,   78,   79,   80 },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			var expectedColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
			var expectedRowData = new int?[] { 2001, null, null, null, null, null, 2002, 2003, 2004, 2005, null, null, null, null, null, null };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedSheetData, cellStore);
		}
		
		[TestMethod]
		public void InsertColumnsAcrossAllRows()
		{
			var cellStore = this.GetCellStore(false);
			var originalSheetData = new int?[,]
			{
/*1*/		{   1,    2,    3,    4,    5,    6, null, null, null, null, null, null, null, null, null, null },
/*2*/		{  17,   18,   19,   20,   21,   22, null, null, null, null, null, null, null, null, null, null },
/*3*/		{  33,   34,   35,   36,   37,   38, null, null, null, null, null, null, null, null, null, null },
/*4*/		{  49,   50,   51,   52,   53,   54, null, null, null, null, null, null, null, null, null, null },
/*5*/		{  65,   66,   67,   68,   69,   70, null, null, null, null, null, null, null, null, null, null },
/*6*/		{  81,   82,   83,   84,   85,   86, null, null, null, null, null, null, null, null, null, null },
/*7*/		{  97,   98,   99,  100,  101,  102, null, null, null, null, null, null, null, null, null, null },
/*8*/		{ 113,  114,  115,  116,  117,  118, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ 129,  130,  131,  132,  133,  134, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ 145,  146,  147,  148,  149,  150, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ 161,  162,  163,  164,  165,  166, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ 177,  178,  179,  180,  181,  182, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ 193,  194,  195,  196,  197,  198, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ 209,  210,  211,  212,  213,  214, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ 225,  226,  227,  228,  229,  230, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ 241,  242,  243,  244,  245,  246, null, null, null, null, null, null, null, null, null, null }
			};
			var originalColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, null, null, null, null, null, null, null, null, null, null };
			var originalRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.LoadCellStore(originalColumnData, originalRowData, originalSheetData, cellStore);
			cellStore.Insert(0, 2, cellStore.MaximumRow, 5);
			var expectedSheetData = new int?[,]
			{
/*1*/		{   1, null, null, null, null, null,    2,    3,    4,    5,    6, null, null, null, null, null },
/*2*/		{  17, null, null, null, null, null,   18,   19,   20,   21,   22, null, null, null, null, null },
/*3*/		{  33, null, null, null, null, null,   34,   35,   36,   37,   38, null, null, null, null, null },
/*4*/		{  49, null, null, null, null, null,   50,   51,   52,   53,   54, null, null, null, null, null },
/*5*/		{  65, null, null, null, null, null,   66,   67,   68,   69,   70, null, null, null, null, null },
/*6*/		{  81, null, null, null, null, null,   82,   83,   84,   85,   86, null, null, null, null, null },
/*7*/		{  97, null, null, null, null, null,   98,   99,  100,  101,  102, null, null, null, null, null },
/*8*/		{ 113, null, null, null, null, null,  114,  115,  116,  117,  118, null, null, null, null, null },
/*9*/		{ 129, null, null, null, null, null,  130,  131,  132,  133,  134, null, null, null, null, null },
/*10*/	{ 145, null, null, null, null, null,  146,  147,  148,  149,  150, null, null, null, null, null },
/*11*/	{ 161, null, null, null, null, null,  162,  163,  164,  165,  166, null, null, null, null, null },
/*12*/	{ 177, null, null, null, null, null,  178,  179,  180,  181,  182, null, null, null, null, null },
/*13*/	{ 193, null, null, null, null, null,  194,  195,  196,  197,  198, null, null, null, null, null },
/*14*/	{ 209, null, null, null, null, null,  210,  211,  212,  213,  214, null, null, null, null, null },
/*15*/	{ 225, null, null, null, null, null,  226,  227,  228,  229,  230, null, null, null, null, null },
/*16*/	{ 241, null, null, null, null, null,  242,  243,  244,  245,  246, null, null, null, null, null }
			};
			var expectedColumnData = new int?[] { 1001, null, null, null, null, null, 1002, 1003, 1004, 1005, 1006, null, null, null, null, null };
			var expectedRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
			this.ValidateCellStore(expectedColumnData, expectedRowData, expectedSheetData, cellStore);
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void InsertInsideShiftsBlock()
		{
			var cellStore = this.GetCellStore(false);
			var originalSheetData = new int?[,]
			{
/*1*/		{    1,    2,    3,    4,    5,    6, null, null, null, null, null, null, null, null, null, null },
/*2*/		{   17,   18,   19,   20,   21,   22, null, null, null, null, null, null, null, null, null, null },
/*3*/		{   33,   34,   35,   36,   37,   38, null, null, null, null, null, null, null, null, null, null },
/*4*/		{   49,   50,   51,   52,   53,   54, null, null, null, null, null, null, null, null, null, null },
/*5*/		{   65,   66,   67,   68,   69,   70, null, null, null, null, null, null, null, null, null, null },
/*6*/		{   81,   82,   83,   84,   85,   86, null, null, null, null, null, null, null, null, null, null },
/*7*/		{   97,   98,   99,  100,  101,  102, null, null, null, null, null, null, null, null, null, null },
/*8*/		{  113,  114,  113,  114,  115,  116, null, null, null, null, null, null, null, null, null, null },
/*9*/		{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*10*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*11*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*12*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*13*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
/*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
			};
			var originalColumnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, null, null, null, null, null, null, null, null, null, null };
			var originalRowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, null, null, null, null, null, null, null, null };
			this.LoadCellStore(originalColumnData, originalRowData, originalSheetData, cellStore);
			cellStore.Insert(2, 2, 5, 6);
			#region Expected Result for if we decide to implement this.
//			var expectedStore = new int?[,]
//			{
///*1*/		{    1,    2,    3,    4,    5,    6, null, null, null, null, null, null, null, null, null, null },
///*2*/		{   17, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*3*/		{   33, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*4*/		{   49, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*5*/		{   65, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*6*/		{   81, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*7*/		{   97, null, null, null, null, null, null,   18,   19,   20,   21,   22, null, null, null, null },
///*8*/		{  113, null, null, null, null, null, null,   34,   35,   36,   37,   38, null, null, null, null },
///*9*/		{ null, null, null, null, null, null, null,   50,   51,   52,   53,   54, null, null, null, null },
///*10*/	{ null, null, null, null, null, null, null,   66,   67,   68,   69,   70, null, null, null, null },
///*11*/	{ null, null, null, null, null, null, null,   82,   83,   84,   85,   86, null, null, null, null },
///*12*/	{ null, null, null, null, null, null, null,   98,   99,  100,  101,  102, null, null, null, null },
///*13*/	{ null, null, null, null, null, null, null,  114,  113,  114,  115,  116, null, null, null, null },
///*14*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*15*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
///*16*/	{ null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }
//			};
//			this.ValidateCellStore(originalColumnData, originalRowData, expectedStore, cellStore);
			#endregion
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
				{ 13, 14, 15, null }
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
				{ 9, 10, 11, null },
				{ null, null, null, null }
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
		public void ShiftItemsFromStartPositiveShiftGoesForwardHandlesEmptyPages()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, 4 },
				{ null, null, null, null },
				{ 9, 10, 11, 12 },
				{ null, null, null, null }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(2, 4);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(15, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, null, null },
				{ null, null, 3, 4 },
				{ null, null, null, null },
				{ 9, 10, 11, 12 }
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
				{ 5, 6, null, null },
				{ null, null, null, null },
				{ null, null, null, null }
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
		public void ShiftItemsFromMiddleMultiplePagesNegativeShiftHandlesEmptyPages()
		{
			var pagedStructure = new PagedStructure<int>(2);
			var items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, 3, null },
				{ null, null, null, null },
				{ null, null, null, null },
				{ 13, 14, 15, 16 }
			};
			pagedStructure.LoadPages(items);
			pagedStructure.ShiftItems(2, -5);
			Assert.AreEqual(0, pagedStructure.MinimumUsedIndex);
			Assert.AreEqual(10, pagedStructure.MaximumUsedIndex);
			items = new ZCellStore<int>.PagedStructure<int>.ValueHolder?[,]
			{
				{ 1, 2, null, null },
				{ null, null, null, 13 },
				{ 14, 15, 16, null },
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

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void ShiftItemsForwardAmountTooHighThrowsException()
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
		}
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
			Assert.AreEqual(0, enumerator.Row);
			Assert.AreEqual(-1, enumerator.Column);
			Assert.IsFalse(enumerator.MoveNext());
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesAllValues()
		{
			var expectedData = new int?[,]
			{
					{ null, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 },
/*1*/			{ 2001,    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/			{ 2002,   17,   18,   19,   20,   21,   22,   23,   24,   25,   26,   27,   28,   29,   30,   31,   32 },
/*3*/			{ 2003,   33,   34,   35,   36,   37,   38,   39,   40,   41,   42,   43,   44,   45,   46,   47,   48 },
/*4*/			{ 2004,   49,   50,   51,   52,   53,   54,   55,   56,   57,   58,   59,   60,   61,   62,   63,   64 },
/*5*/			{ 2005,   65,   66,   67,   68,   69,   70,   71,   72,   73,   74,   75,   76,   77,   78,   79,   80 },
/*6*/			{ 2006,   81,   82,   83,   84,   85,   86,   87,   88,   89,   90,   91,   92,   93,   94,   95,   96 },
/*7*/			{ 2007,   97,   98,   99,  100,  101,  102,  103,  104,  105,  106,  107,  108,  109,  110,  111,  112 },
/*8*/			{ 2008,  113,  114,  115,  116,  117,  118,  119,  120,  121,  122,  123,  124,  125,  126,  127,  128 },
/*9*/			{ 2009,  129,  130,  131,  132,  133,  134,  135,  136,  137,  138,  139,  140,  141,  142,  143,  144 },
/*10*/		{ 2010,  145,  146,  147,  148,  149,  150,  151,  152,  153,  154,  155,  156,  157,  158,  159,  160 },
/*11*/		{ 2011,  161,  162,  163,  164,  165,  166,  167,  168,  169,  170,  171,  172,  173,  174,  175,  176 },
/*12*/		{ 2012,  177,  178,  179,  180,  181,  182,  183,  184,  185,  186,  187,  188,  189,  190,  191,  192 },
/*13*/		{ 2013,  193,  194,  195,  196,  197,  198,  199,  200,  201,  202,  203,  204,  205,  206,  207,  208 },
/*14*/		{ 2014,  209,  210,  211,  212,  213,  214,  215,  216,  217,  218,  219,  220,  221,  222,  223,  224 },
/*15*/		{ 2015,  225,  226,  227,  228,  229,  230,  231,  232,  233,  234,  235,  236,  237,  238,  239,  240 },
/*16*/		{ 2016,  241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  256 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator();
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(288, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesAllValuesSkipsEmptyCells()
		{
			var nullCells = new Tuple<int, int>[]
			{
				new Tuple<int, int>(2, 8),
				new Tuple<int, int>(3, 3),
				new Tuple<int, int>(3, 8),
				new Tuple<int, int>(4, 8),
				new Tuple<int, int>(5, 8),
				new Tuple<int, int>(6, 8),
				new Tuple<int, int>(7, 0),
				new Tuple<int, int>(7, 8),
				new Tuple<int, int>(8, 8),
				new Tuple<int, int>(9, 8),
				new Tuple<int, int>(10, 8),
				new Tuple<int, int>(13, 1),
				new Tuple<int, int>(13, 4),
				new Tuple<int, int>(13, 13)
			};
			var expectedData = new int?[,]
			{
					{ null, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 },
/*1*/			{ 2001,    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/			{ 2002,   17,   18,   19,   20,   21,   22,   23, null,   25,   26,   27,   28,   29,   30,   31,   32 },
/*3*/			{ 2003,   33,   34, null,   36,   37,   38,   39, null,   41,   42,   43,   44,   45,   46,   47,   48 },
/*4*/			{ 2004,   49,   50,   51,   52,   53,   54,   55, null,   57,   58,   59,   60,   61,   62,   63,   64 },
/*5*/			{ 2005,   65,   66,   67,   68,   69,   70,   71, null,   73,   74,   75,   76,   77,   78,   79,   80 },
/*6*/			{ 2006,   81,   82,   83,   84,   85,   86,   87, null,   89,   90,   91,   92,   93,   94,   95,   96 },
/*7*/			{ null,   97,   98,   99,  100,  101,  102,  103, null,  105,  106,  107,  108,  109,  110,  111,  112 },
/*8*/			{ 2008,  113,  114,  115,  116,  117,  118,  119, null,  121,  122,  123,  124,  125,  126,  127,  128 },
/*9*/			{ 2009,  129,  130,  131,  132,  133,  134,  135, null,  137,  138,  139,  140,  141,  142,  143,  144 },
/*10*/		{ 2010,  145,  146,  147,  148,  149,  150,  151, null,  153,  154,  155,  156,  157,  158,  159,  160 },
/*11*/		{ 2011,  161,  162,  163,  164,  165,  166,  167,  168,  169,  170,  171,  172,  173,  174,  175,  176 },
/*12*/		{ 2012,  177,  178,  179,  180,  181,  182,  183,  184,  185,  186,  187,  188,  189,  190,  191,  192 },
/*13*/		{ 2013, null,  194,  195, null,  197,  198,  199,  200,  201,  202,  203,  204, null,  206,  207,  208 },
/*14*/		{ 2014,  209,  210,  211,  212,  213,  214,  215,  216,  217,  218,  219,  220,  221,  222,  223,  224 },
/*15*/		{ 2015,  225,  226,  227,  228,  229,  230,  231,  232,  233,  234,  235,  236,  237,  238,  239,  240 },
/*16*/		{ 2016,  241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  256 }
			};
			var cellStore = this.GetCellStore(nullCoordinates: nullCells);
			var enumerator = cellStore.GetEnumerator();
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column];
				Assert.IsTrue(item.HasValue, $"Erroneously hit: [{enumerator.Row}:{enumerator.Column}]");
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(274, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockTopRowsIncludingMetadata()
		{
			var expectedData = new int?[,]
			{
					{ null, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 },
/*1*/			{ 2001,    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/			{ 2002,   17,   18,   19,   20,   21,   22,   23,   24,   25,   26,   27,   28,   29,   30,   31,   32 },
/*3*/			{ 2003,   33,   34,   35,   36,   37,   38,   39,   40,   41,   42,   43,   44,   45,   46,   47,   48 },
/*4*/			{ 2004,   49,   50,   51,   52,   53,   54,   55,   56,   57,   58,   59,   60,   61,   62,   63,   64 },
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(0, 0, 4, 16);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(84, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockTopRowsExcludingMetadata()
		{
			var expectedData = new int?[,]
			{
/*1*/			{ 2001,    1,    2,    3,    4,    5,    6,    7,    8,    9,   10,   11,   12,   13,   14,   15,   16 },
/*2*/			{ 2002,   17,   18,   19,   20,   21,   22,   23,   24,   25,   26,   27,   28,   29,   30,   31,   32 },
/*3*/			{ 2003,   33,   34,   35,   36,   37,   38,   39,   40,   41,   42,   43,   44,   45,   46,   47,   48 },
/*4*/			{ 2004,   49,   50,   51,   52,   53,   54,   55,   56,   57,   58,   59,   60,   61,   62,   63,   64 },
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(1, 0, 4, 16);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row - 1, enumerator.Column];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(68, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockBottomRows()
		{
			var expectedData = new int?[,]
			{
/*12*/		{ 2012,  177,  178,  179,  180,  181,  182,  183,  184,  185,  186,  187,  188,  189,  190,  191,  192 },
/*13*/		{ 2013,  193,  194,  195,  196,  197,  198,  199,  200,  201,  202,  203,  204,  205,  206,  207,  208 },
/*14*/		{ 2014,  209,  210,  211,  212,  213,  214,  215,  216,  217,  218,  219,  220,  221,  222,  223,  224 },
/*15*/		{ 2015,  225,  226,  227,  228,  229,  230,  231,  232,  233,  234,  235,  236,  237,  238,  239,  240 },
/*16*/		{ 2016,  241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  256 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(12, 0, 16, 16);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row - 12, enumerator.Column];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(85, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockBottomRowsIgnoresLudicrousDimensions()
		{
			var expectedData = new int?[,]
			{
/*12*/		{ 2012,  177,  178,  179,  180,  181,  182,  183,  184,  185,  186,  187,  188,  189,  190,  191,  192 },
/*13*/		{ 2013,  193,  194,  195,  196,  197,  198,  199,  200,  201,  202,  203,  204,  205,  206,  207,  208 },
/*14*/		{ 2014,  209,  210,  211,  212,  213,  214,  215,  216,  217,  218,  219,  220,  221,  222,  223,  224 },
/*15*/		{ 2015,  225,  226,  227,  228,  229,  230,  231,  232,  233,  234,  235,  236,  237,  238,  239,  240 },
/*16*/		{ 2016,  241,  242,  243,  244,  245,  246,  247,  248,  249,  250,  251,  252,  253,  254,  255,  256 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(12, 0, 1000, 1000);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row - 12, enumerator.Column];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(85, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockLeftColumnsIncludingMetadata()
		{
			var expectedData = new int?[,]
			{
					{ null, 1001, 1002, 1003, 1004 },
/*1*/			{ 2001,    1,    2,    3,    4 },
/*2*/			{ 2002,   17,   18,   19,   20 },
/*3*/			{ 2003,   33,   34,   35,   36 },
/*4*/			{ 2004,   49,   50,   51,   52 },
/*5*/			{ 2005,   65,   66,   67,   68 },
/*6*/			{ 2006,   81,   82,   83,   84 },
/*7*/			{ 2007,   97,   98,   99,  100 },
/*8*/			{ 2008,  113,  114,  115,  116 },
/*9*/			{ 2009,  129,  130,  131,  132 },
/*10*/		{ 2010,  145,  146,  147,  148 },
/*11*/		{ 2011,  161,  162,  163,  164 },
/*12*/		{ 2012,  177,  178,  179,  180 },
/*13*/		{ 2013,  193,  194,  195,  196 },
/*14*/		{ 2014,  209,  210,  211,  212 },
/*15*/		{ 2015,  225,  226,  227,  228 },
/*16*/		{ 2016,  241,  242,  243,  244 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(0, 0, 16, 4);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(84, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockLeftColumnsExcludingMetadata()
		{
			var expectedData = new int?[,]
			{
					/*	1			2			3			4	 */
					{ 1001, 1002, 1003, 1004 },
/*1*/			{    1,    2,    3,    4 },
/*2*/			{   17,   18,   19,   20 },
/*3*/			{   33,   34,   35,   36 },
/*4*/			{   49,   50,   51,   52 },
/*5*/			{   65,   66,   67,   68 },
/*6*/			{   81,   82,   83,   84 },
/*7*/			{   97,   98,   99,  100 },
/*8*/			{  113,  114,  115,  116 },
/*9*/			{  129,  130,  131,  132 },
/*10*/		{  145,  146,  147,  148 },
/*11*/		{  161,  162,  163,  164 },
/*12*/		{  177,  178,  179,  180 },
/*13*/		{  193,  194,  195,  196 },
/*14*/		{  209,  210,  211,  212 },
/*15*/		{  225,  226,  227,  228 },
/*16*/		{  241,  242,  243,  244 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(0, 1, 16, 4);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column - 1];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(68, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockRightColumns()
		{
			var expectedData = new int?[,]
			{
					/*	11		12		13		14		15		16 */
					{ 1011, 1012, 1013, 1014, 1015, 1016 },
/*1*/			{   11,   12,   13,   14,   15,   16 },
/*2*/			{   27,   28,   29,   30,   31,   32 },
/*3*/			{   43,   44,   45,   46,   47,   48 },
/*4*/			{   59,   60,   61,   62,   63,   64 },
/*5*/			{   75,   76,   77,   78,   79,   80 },
/*6*/			{   91,   92,   93,   94,   95,   96 },
/*7*/			{  107,  108,  109,  110,  111,  112 },
/*8*/			{  123,  124,  125,  126,  127,  128 },
/*9*/			{  139,  140,  141,  142,  143,  144 },
/*10*/		{  155,  156,  157,  158,  159,  160 },
/*11*/		{  171,  172,  173,  174,  175,  176 },
/*12*/		{  187,  188,  189,  190,  191,  192 },
/*13*/		{  203,  204,  205,  206,  207,  208 },
/*14*/		{  219,  220,  221,  222,  223,  224 },
/*15*/		{  235,  236,  237,  238,  239,  240 },
/*16*/		{  251,  252,  253,  254,  255,  256 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(0, 11, 16, 16);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column - 11];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(102, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesBlockRightColumnsIgnoresLudicrousDimensions()
		{
			var expectedData = new int?[,]
			{
					/*	11		12		13		14		15		16 */
					{ 1011, 1012, 1013, 1014, 1015, 1016 },
/*1*/			{   11,   12,   13,   14,   15,   16 },
/*2*/			{   27,   28,   29,   30,   31,   32 },
/*3*/			{   43,   44,   45,   46,   47,   48 },
/*4*/			{   59,   60,   61,   62,   63,   64 },
/*5*/			{   75,   76,   77,   78,   79,   80 },
/*6*/			{   91,   92,   93,   94,   95,   96 },
/*7*/			{  107,  108,  109,  110,  111,  112 },
/*8*/			{  123,  124,  125,  126,  127,  128 },
/*9*/			{  139,  140,  141,  142,  143,  144 },
/*10*/		{  155,  156,  157,  158,  159,  160 },
/*11*/		{  171,  172,  173,  174,  175,  176 },
/*12*/		{  187,  188,  189,  190,  191,  192 },
/*13*/		{  203,  204,  205,  206,  207,  208 },
/*14*/		{  219,  220,  221,  222,  223,  224 },
/*15*/		{  235,  236,  237,  238,  239,  240 },
/*16*/		{  251,  252,  253,  254,  255,  256 }
			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(0, 11, 1000, 1000);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row, enumerator.Column - 11];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(102, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesInternalBlock()
		{
			var expectedData = new int?[,]
			{
					/*  5		  6			7			8			9    10  */
/*6*/			{   85,   86,   87,   88,   89,   90 },
/*7*/			{  101,  102,  103,  104,  105,  106 },
/*8*/			{  117,  118,  119,  120,  121,  122 },
/*9*/			{  133,  134,  135,  136,  137,  138 },
/*10*/		{  149,  150,  151,  152,  153,  154 },
/*11*/		{  165,  166,  167,  168,  169,  170 },
/*12*/		{  181,  182,  183,  184,  185,  186 },

			};
			var cellStore = this.GetCellStore();
			var enumerator = cellStore.GetEnumerator(6, 5, 12, 10);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row - 6, enumerator.Column - 5];
				Assert.IsTrue(item.HasValue);
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(42, resultCount);
		}

		[TestMethod]
		public void ZCellStoreEnumeratorEnumeratesInternalBlockSkipsEmptyCells()
		{
			var nullCells = new Tuple<int, int>[]
			{
				new Tuple<int, int>(7, 6),
				new Tuple<int, int>(7, 7),
				new Tuple<int, int>(7, 10),
				new Tuple<int, int>(8, 10),
				new Tuple<int, int>(9, 10),
				new Tuple<int, int>(10, 10),
				new Tuple<int, int>(11, 10)
			};
			var expectedData = new int?[,]
			{
					/*  5		  6			7			8			9    10  */
/*6*/			{   85,   86,   87,   88,   89,   90 },
/*7*/			{  101, null, null,  104,  105, null },
/*8*/			{  117,  118,  119,  120,  121, null },
/*9*/			{  133,  134,  135,  136,  137, null },
/*10*/		{  149,  150,  151,  152,  153, null },
/*11*/		{  165,  166,  167,  168,  169, null },
/*12*/		{  181,  182,  183,  184,  185,  186 },

			};
			var cellStore = this.GetCellStore(nullCoordinates: nullCells);
			var enumerator = cellStore.GetEnumerator(6, 5, 12, 10);
			int resultCount = 0;
			while (enumerator.MoveNext())
			{
				resultCount++;
				var item = expectedData[enumerator.Row - 6, enumerator.Column - 5];
				Assert.AreEqual(item.Value, enumerator.Value);
			}
			Assert.AreEqual(35, resultCount);
		}
		#endregion
		#endregion

		#region Helper Methods
		private ICellStore<int> GetCellStore(bool fill = true, Tuple<int, int>[] nullCoordinates = null)
		{
			var cellStore = new ZCellStore<int>(2, 2);
			if (fill)
			{
				var columnData = new int?[] { 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016 };
				var rowData = new int?[] { 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016 };
				var currentStore = new int?[,]
				{
/*1*/			{   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15,  16 },
/*2*/			{  17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27,  28,  29,  30,  31,  32 },
/*3*/			{  33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,  46,  47,  48 },
/*4*/			{  49,  50,  51,  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64 },
/*5*/			{  65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80 },
/*6*/			{  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95,  96 },
/*7*/			{  97,  98,  99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112 },
/*8*/			{ 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128 },
/*9*/			{ 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144 },
/*10*/		{ 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160 },
/*11*/		{ 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176 },
/*12*/		{ 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192 },
/*13*/		{ 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208 },
/*14*/		{ 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224 },
/*15*/		{ 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240 },
/*16*/		{ 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256 }
				};
				this.LoadCellStore(columnData, rowData, currentStore, cellStore);
				if (nullCoordinates != null)
				{
					foreach (var coordinate in nullCoordinates)
					{
						cellStore.Clear(coordinate.Item1, coordinate.Item2, 1, 1);
					}
				}
			}
			return cellStore;
		}

		private void LoadCellStore(int?[] columnData, int?[] rowData, int?[,] sheetData, ICellStore<int> cellStore)
		{
			if (columnData != null)
			{
				for (int column = 0; column < columnData.Length; column++)
				{
					if (columnData[column].HasValue)
						cellStore.SetValue(0, column + 1, columnData[column].Value);
				}
			}
			if (rowData != null)
			{
				for (int row = 0; row < rowData.Length; row++)
				{
					if (rowData[row].HasValue)
						cellStore.SetValue(row + 1, 0, rowData[row].Value);
				}
			}
			if (sheetData != null)
			{
				for (int row = 0; row <= sheetData.GetUpperBound(0); ++row)
				{
					for (int column = 0; column <= sheetData.GetUpperBound(1); ++column)
					{
						var data = sheetData[row, column];
						if (data.HasValue)
							cellStore.SetValue(row + 1, column + 1, data.Value);
					}
				}
			}
		}

		private void ValidateCellStore(int?[] columnData, int?[] rowData, int?[,] sheetData, ICellStore<int> cellStore)
		{
			if (columnData != null)
			{
				for (int column = 0; column < columnData.Length; column++)
				{
					var data = columnData[column];
					var item = cellStore.GetValue(0, column + 1);
					if (data.HasValue)
						Assert.AreEqual(data.Value, item, $"Data mismatch row: 0 column: {column}");
					else
						Assert.AreEqual(default(int), item, $"Data mismatch row: 0 column: {column}");
				}
			}
			if (rowData != null)
			{
				for (int row = 0; row < rowData.Length; row++)
				{
					var data = rowData[row];
					var item = cellStore.GetValue(row + 1, 0);
					if (data.HasValue)
						Assert.AreEqual(data.Value, item, $"Data mismatch row: {row} column: 0");
					else
						Assert.AreEqual(default(int), item, $"Data mismatch row: {row} column: 0");
				}
			}
			if (sheetData != null)
			{
				for (int row = 0; row <= sheetData.GetUpperBound(0); ++row)
				{
					for (int column = 0; column <= sheetData.GetUpperBound(1); ++column)
					{
						var data = sheetData[row, column];
						var item = cellStore.GetValue(row + 1, column + 1);
						if (data.HasValue)
							Assert.AreEqual(data.Value, item, $"Data mismatch row: {row} column: {column}");
						else
							Assert.AreEqual(default(int), item, $"Data mismatch row: {row} column: {column}");
					}
				}
			}
		}
		#endregion
	}
}
