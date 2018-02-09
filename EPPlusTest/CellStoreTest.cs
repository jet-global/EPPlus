using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class CellStoreTest : TestBase
	{
		[TestMethod]
		public void CellStorePagingBreaksOnceAFullyFormedPageNeedsToBeSplit()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["A1"].Value = "Has some text in A1";
				sheet.Cells["A2"].Value = "Has some text in A2";
				sheet.Cells["B1"].Value = "Has some text in B1";
				sheet.Cells["B2"].Value = "Has some text in B2";
				sheet.Cells["C1"].Value = "Has some text in C1";
				sheet.Cells["C2"].Value = "Has some text in C2";
				sheet.Cells["D1"].Value = "Has some text in D1";
				sheet.Cells["D2"].Value = "Has some text in D2";

				sheet.Cells["E22"].Value = "Has some text to be copied.";
				sheet.Cells["F22"].Value = "Has some other text to be copied.";

				sheet.Cells["H23"].Value = "Has even more text to be copied.";
				sheet.Cells["I23"].Value = "Has even more distinct text to be copied.";

				sheet.Cells["K24"].Value = "Has even more text that could be copied if necessary.";
				sheet.Cells["L24"].Value = "Has even more distinct text that could be copied.";

				int row = 22;
				this.InsertTheRightAmountOfSpace(sheet, row, 5);
				row += 67;
				for (int i = row; i > row - 67; i--)
				{
					this.InsertTheRightAmountOfSpace(sheet, i, 6);
				}

				row = 4579;
				this.InsertTheRightAmountOfSpace(sheet, row, 8);
				row += 67;
				for (int i = row; i > row - 67; i--)
				{
					this.InsertTheRightAmountOfSpace(sheet, i, 9);
				}
				sheet.Calculate();
				for (int i = 22; i < 4579; i++)
				{
					Assert.IsTrue(!string.IsNullOrEmpty(sheet.Cells[i, 5].Value?.ToString()), $"Cell at [{i},5] is null or empty.");
					Assert.IsTrue(!string.IsNullOrEmpty(sheet.Cells[i, 6].Value?.ToString()), $"Cell at [{i},6] is null or empty.");
				}

				for (int i = 4579; i < 9136; i++)
				{
					Assert.IsTrue(!string.IsNullOrEmpty(sheet.Cells[i, 8].Value?.ToString()), $"Cell at [{i},8] is null or empty.");
					Assert.IsTrue(!string.IsNullOrEmpty(sheet.Cells[i, 9].Value?.ToString()), $"Cell at [{i},9] is null or empty.");
				}

				Assert.AreEqual("Has even more text that could be copied if necessary.", sheet.Cells[9136, 11].Value);
				Assert.AreEqual("Has even more distinct text that could be copied.", sheet.Cells[9136, 12].Value);
			}
		}

		private void InsertTheRightAmountOfSpace(ExcelWorksheet sheet, int row, int column)
		{
			sheet.InsertRow(row + 1, 67);
			for (int i = 1; i < 68; i++)
			{
				sheet.Cells[$"{row}:{row}"].Copy(sheet.Cells[$"{row + i}:{row + i}"]);
				sheet.Cells[row + i, column].Formula = $"\"Updated Unique Value {Guid.NewGuid().ToString()}\"";
			}
		}

		[TestMethod]
		public void DeleteRowsHandlesDifferentShapesOfColumns()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				for (int row = 4; row <= 8; row++)
				{
					sheet.Cells[row, 2].Formula = $"{row}";
				}
				for (int row = 1; row <= 8; row++)
				{
					sheet.Cells[row, 3].Formula = $"{row}";
				}
				sheet.DeleteRow(2, 4);
				Assert.AreEqual("1", sheet.Cells[1, 3].Formula);
				Assert.AreEqual("6", sheet.Cells[2, 2].Formula);
				Assert.AreEqual("6", sheet.Cells[2, 3].Formula);
				Assert.AreEqual("7", sheet.Cells[3, 2].Formula);
				Assert.AreEqual("7", sheet.Cells[3, 3].Formula);
				Assert.AreEqual("8", sheet.Cells[4, 2].Formula);
				Assert.AreEqual("8", sheet.Cells[4, 3].Formula);
			}
		}

		[TestMethod]
		public void DeleteRowsAcrossMultipleCellStorePages()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				for (int i = 1; i < 1024; i++)
				{
					sheet.Cells[i, 2].Value = i;
				}
				for (int i = 1500; i < 4096; i++)
				{
					sheet.Cells[i, 2].Value = i;
				}
				int splitPoint = 400;
				sheet.DeleteRow(splitPoint, 600);
				sheet.DeleteRow(splitPoint, 600);
				for (int i = 1; i < splitPoint; i++)
				{
					Assert.AreEqual(i, sheet.Cells[i, 2].Value);
				}
				for (int i = splitPoint; i < 2872; i++)
				{
					Assert.AreEqual(i + 1200, sheet.Cells[i, 2].Value);
				}
			}
		}

		[TestMethod]
		public void Insert1()
		{
			var ws = _pck.Workbook.Worksheets.Add("Insert1");
			LoadData(ws);

			ws.InsertRow(2, 1000);
			Assert.AreEqual(ws.GetValue(1002, 1), "1,0");
			ws.InsertRow(1003, 1000);
			Assert.AreEqual(ws.GetValue(2003, 1), "2,0");
			ws.InsertRow(2004, 1000);
			Assert.AreEqual(ws.GetValue(3004, 1), "3,0");
			ws.InsertRow(2006, 1000);
			Assert.AreEqual(ws.GetValue(4005, 1), "4,0");
			ws.InsertRow(4500, 500);
			Assert.AreEqual(ws.GetValue(5000, 1), "499,0");

			ws.InsertRow(1, 1);
			Assert.AreEqual(ws.GetValue(1003, 1), "1,0");
			Assert.AreEqual(ws.GetValue(5001, 1), "499,0");

			ws.InsertRow(1, 15);
			Assert.AreEqual(ws.GetValue(4020, 1), "3,0");
			Assert.AreEqual(ws.GetValue(5016, 1), "499,0");
		}

		[TestMethod]
		public void Insert2()
		{
			var ws = _pck.Workbook.Worksheets.Add("Insert2-1");
			LoadData(ws);

			for (int i = 0; i < 32; i++)
			{
				ws.InsertRow(1, 1);
			}
			Assert.AreEqual(ws.GetValue(33, 1), "0,0");

			ws = _pck.Workbook.Worksheets.Add("Insert2-2");
			LoadData(ws);

			for (int i = 0; i < 32; i++)
			{
				ws.InsertRow(15, 1);
			}
			Assert.AreEqual(ws.GetValue(1, 1), "0,0");
			Assert.AreEqual(ws.GetValue(47, 1), "14,0");
		}

		[TestMethod]
		public void Insert3()
		{
			var ws = _pck.Workbook.Worksheets.Add("Insert3");
			LoadData(ws);

			for (int i = 0; i < 500; i += 4)
			{
				ws.InsertRow(i + 1, 2);
			}
		}

		[TestMethod]
		public void InsertRandomTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("Insert4-1");

			LoadData(ws, 5000);

			for (int i = 5000; i > 0; i -= 2)
			{
				ws.InsertRow(i, 1);
			}
		}

		[TestMethod]
		public void EnumCellstore()
		{
			var ws = _pck.Workbook.Worksheets.Add("enum");

			LoadData(ws, 5000);

			var o = ws._values.GetEnumerator(2, 1, 5, 3);
			foreach (var i in o)
			{
				Console.WriteLine(i);
			}
		}

		[TestMethod]
		public void DeleteCells()
		{
			var ws = _pck.Workbook.Worksheets.Add("Delete");
			LoadData(ws, 5000);

			ws.DeleteRow(2, 2);
			Assert.AreEqual("3,0", ws.GetValue(2, 1));
			ws.DeleteRow(10, 10);
			Assert.AreEqual("21,0", ws.GetValue(10, 1));
			ws.DeleteRow(50, 40);
			Assert.AreEqual("101,0", ws.GetValue(50, 1));
			ws.DeleteRow(100, 100);
			Assert.AreEqual("251,0", ws.GetValue(100, 1));
			ws.DeleteRow(1, 31);
			Assert.AreEqual("43,0", ws.GetValue(1, 1));
		}

		[TestMethod]
		public void DeleteCellsFirst()
		{
			var ws = _pck.Workbook.Worksheets.Add("DeleteFirst");
			LoadData(ws, 5000);

			ws.DeleteRow(32, 30);
			for (int i = 1; i < 50; i++)
			{
				ws.DeleteRow(1, 1);
			}
		}

		[TestMethod]
		public void DeleteInsert()
		{
			var ws = _pck.Workbook.Worksheets.Add("DeleteInsert");
			LoadData(ws, 5000);

			ws.DeleteRow(2, 33);
			ws.InsertRow(2, 38);

			for (int i = 0; i < 33; i++)
			{
				ws.SetValue(i + 2, 1, i + 2);
			}
		}

		private void LoadData(ExcelWorksheet ws)
		{
			LoadData(ws, 1000);
		}

		private void LoadData(ExcelWorksheet ws, int rows, int cols = 1, bool isNumeric = false)
		{
			for (int r = 0; r < rows; r++)
			{
				for (int c = 0; c < cols; c++)
				{
					if (isNumeric)
						ws.SetValue(r + 1, c + 1, r + c);
					else
						ws.SetValue(r + 1, c + 1, r.ToString() + "," + c.ToString());
				}
			}
		}

		[TestMethod]
		public void FillInsertTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("FillInsert");

			LoadData(ws, 500);

			var r = 1;
			for (int i = 1; i <= 500; i++)
			{
				ws.InsertRow(r, i);
				Assert.AreEqual((i - 1).ToString() + ",0", ws.GetValue(r + i, 1).ToString());
				r += i + 1;
			}
		}

		[TestMethod]
		public void CopyCellsTest()
		{
			var ws = _pck.Workbook.Worksheets.Add("CopyCells");

			LoadData(ws, 100, isNumeric: true);
			ws.Cells["B1"].Formula = "SUM(A1:A500)";
			ws.Calculate();
			ws.Cells["B1"].Copy(ws.Cells["C1"]);
			ws.Cells["B1"].Copy(ws.Cells["D1"], ExcelRangeCopyOptionFlags.ExcludeFormulas);

			Assert.AreEqual(ws.Cells["B1"].Value, ws.Cells["C1"].Value);
			Assert.AreEqual("SUM(B1:B500)", ws.Cells["C1"].Formula);

			Assert.AreEqual(ws.Cells["B1"].Value, ws.Cells["D1"].Value);
			Assert.AreNotEqual(ws.Cells["B1"].Formula, ws.Cells["D1"].Formula);
		}
	}
}
