using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[DeploymentItem("Workbooks", "targetFolder")]
	[TestClass]
	public class Calculation
	{
		#region Calculate Tests
		[TestMethod]
		public void CalulationTestDatatypes()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Calc1");
			ws.SetValue("A1", (short)1);
			ws.SetValue("A2", (long)2);
			ws.SetValue("A3", (Single)3);
			ws.SetValue("A4", (double)4);
			ws.SetValue("A5", (Decimal)5);
			ws.SetValue("A6", (byte)6);
			ws.SetValue("A7", null);
			ws.Cells["A10"].Formula = "Sum(A1:A8)";
			ws.Cells["A11"].Formula = "SubTotal(9,A1:A8)";
			ws.Cells["A12"].Formula = "Average(A1:A8)";

			ws.Calculate();
			Assert.AreEqual(21D, ws.Cells["a10"].Value);
			Assert.AreEqual(21D, ws.Cells["a11"].Value);
			Assert.AreEqual(21D / 6, ws.Cells["a12"].Value);
		}

		[TestMethod]
		public void CalculateTest()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Calc1");

			ws.SetValue("A1", (short)1);
			var v = ws.Calculate("2.5-A1+ABS(-3.0)-SIN(3)");
			Assert.AreEqual(4.3589, Math.Round((double)v, 4));

			ws.Row(1).Hidden = true;
			v = ws.Calculate("subtotal(109,a1:a10)");
			Assert.AreEqual(0D, v);

			v = ws.Calculate("-subtotal(9,a1:a3)");
			Assert.AreEqual(-1D, v);
		}

		[TestMethod]
		public void CalculateTestIsFunctions()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Calc1");

			ws.SetValue(1, 1, 1.0D);
			ws.SetFormula(1, 2, "isblank(A1:A5)");
			ws.SetFormula(1, 3, "concatenate(a1,a2,a3)");
			ws.SetFormula(1, 4, "Row()");
			ws.SetFormula(1, 5, "Row(a3)");
			ws.Calculate();
		}

		[TestMethod]
		[DeploymentItem(@"Workbooks\FormulaTest.xlsx")]
		public void Calulation4()
		{
			var file = new FileInfo("FormulaTest.xlsx");
			Assert.IsTrue(file.Exists);
			var pck = new ExcelPackage(file);
			pck.Workbook.Calculate();
			Assert.AreEqual(490D, pck.Workbook.Worksheets[1].Cells["D5"].Value);
		}

		[TestMethod]
		[DeploymentItem(@"Workbooks\FormulaTest.xlsx")]
		public void CalulationValidationExcel()
		{
			var file = new FileInfo("FormulaTest.xlsx");
			Assert.IsTrue(file.Exists);

			var pck = new ExcelPackage(file);

			var ws = pck.Workbook.Worksheets["ValidateFormulas"];
			var fr = new Dictionary<string, object>();
			foreach (var cell in ws.Cells)
			{
				if (!string.IsNullOrEmpty(cell.Formula))
				{
					fr.Add(cell.Address, cell.Value);
				}
			}
			pck.Workbook.Calculate();
			var nErrors = 0;
			var errors = new List<Tuple<string, object, object>>();
			foreach (var adr in fr.Keys)
			{
				try
				{
					if (fr[adr] is double && ws.Cells[adr].Value is double)
					{
						var d1 = Convert.ToDouble(fr[adr]);
						var d2 = Convert.ToDouble(ws.Cells[adr].Value);
						if (Math.Abs(d1 - d2) < 0.0001)
						{
							continue;
						}
						else
						{
							Assert.AreEqual(fr[adr], ws.Cells[adr].Value);
						}
					}
					else
					{
						Assert.AreEqual(fr[adr], ws.Cells[adr].Value);
					}
				}
				catch
				{
					errors.Add(new Tuple<string, object, object>(adr, fr[adr], ws.Cells[adr].Value));
					nErrors++;
				}
			}
		}

		[TestMethod]
		public void CalcTwiceError()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("CalcTest");

			ws.Names.Add("PRICE", "10");
			ws.Names.Add("QUANTITY", "30");
			ws.Cells["A1"].Formula = "PRICE*QUANTITY";
			ws.Names.Add("AMOUNT", "PRICE*QUANTITY");
			ws.Calculate();
			Assert.AreEqual(300D, ws.Cells["A1"].Value);

			ws.Names["PRICE"].NameFormula = "40";
			ws.Names["QUANTITY"].NameFormula = "20";
			ws.Calculate();
			Assert.AreEqual(800D, ws.Cells["A1"].Value);
		}

		[TestMethod]
		public void IfError()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("CalcTest");
			ws.Cells["A1"].Value = "test1";
			ws.Cells["A5"].Value = "test2";
			ws.Cells["A2"].Value = "Sant";
			ws.Cells["A3"].Value = "Falskt";
			ws.Cells["A4"].Formula = "if(A1>=A5,true,A3)";
			ws.Cells["B1"].Formula = "isText(a1)";
			ws.Cells["B2"].Formula = "isText(\"Test\")";
			ws.Cells["B3"].Formula = "isText(1)";
			ws.Cells["B4"].Formula = "isText(true)";
			ws.Cells["c1"].Formula = "mid(a1,4,15)";

			ws.Calculate();
		}

		[TestMethod]
		public void LeftRightFunctionTest()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("CalcTest");
			ws.SetValue("A1", "asdf");
			ws.Cells["A2"].Formula = "Left(A1, 3)";
			ws.Cells["A3"].Formula = "Left(A1, 10)";
			ws.Cells["A4"].Formula = "Right(A1, 3)";
			ws.Cells["A5"].Formula = "Right(A1, 10)";

			ws.Calculate();
			Assert.AreEqual("asd", ws.Cells["A2"].Value);
			Assert.AreEqual("asdf", ws.Cells["A3"].Value);
			Assert.AreEqual("sdf", ws.Cells["A4"].Value);
			Assert.AreEqual("asdf", ws.Cells["A5"].Value);
		}

		[TestMethod]
		public void IfFunctionTest()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("CalcTest");
			ws.SetValue("A1", 123);
			ws.Cells["A2"].Formula = "IF(A1 = 123, 1, -1)";
			ws.Cells["A3"].Formula = "IF(A1 = 1, 1)";
			ws.Cells["A4"].Formula = "IF(A1 = 1, 1, -1)";
			ws.Cells["A5"].Formula = "IF(A1 = 123, 5)";

			ws.Calculate();
			Assert.AreEqual(1d, ws.Cells["A2"].Value);
			Assert.AreEqual(false, ws.Cells["A3"].Value);
			Assert.AreEqual(-1d, ws.Cells["A4"].Value);
			Assert.AreEqual(5d, ws.Cells["A5"].Value);
		}

		[TestMethod]
		public void INTFunctionTest()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("CalcTest");
			var currentDate = DateTime.UtcNow.Date;
			ws.SetValue("A1", currentDate.ToString("MM/dd/yyyy"));
			ws.SetValue("A2", currentDate.Date);
			ws.SetValue("A3", "31.1");
			ws.SetValue("A4", 31.1);
			ws.Cells["A5"].Formula = "INT(A1)";
			ws.Cells["A6"].Formula = "INT(A2)";
			ws.Cells["A7"].Formula = "INT(A3)";
			ws.Cells["A8"].Formula = "INT(A4)";

			ws.Calculate();
			Assert.AreEqual((int)currentDate.ToOADate(), ws.Cells["A5"].Value);
			Assert.AreEqual((int)currentDate.ToOADate(), ws.Cells["A6"].Value);
			Assert.AreEqual(31, ws.Cells["A7"].Value);
			Assert.AreEqual(31, ws.Cells["A8"].Value);
		}

		[TestMethod]
		public void CalculateDateMath()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Test");
				var dateCell = worksheet.Cells[2, 2];
				var date = new DateTime(2013, 1, 1);
				dateCell.Value = date;
				var quotedDateCell = worksheet.Cells[2, 3];
				quotedDateCell.Formula = $"\"{date.ToString("d")}\"";
				var dateFormula = "B2";
				var dateFormulaWithMath = "B2+1";
				var quotedDateFormulaWithMath = $"\"{date.ToString("d")}\"+1";
				var quotedDateReferenceFormulaWithMath = "C2+1";
				var expectedDate = DateTime.FromOADate(41275.0); // January 1, 2013
				var expectedDateDecimalWithMath = 41276.0;
				var expectedDateWithMath = DateTime.FromOADate(expectedDateDecimalWithMath); // January 2, 2013
				Assert.AreEqual(expectedDate, worksheet.Calculate(dateFormula));
				Assert.AreEqual(expectedDateWithMath, worksheet.Calculate(dateFormulaWithMath));
				Assert.AreEqual(expectedDateDecimalWithMath, worksheet.Calculate(quotedDateFormulaWithMath));
				Assert.AreEqual(expectedDateDecimalWithMath, worksheet.Calculate(quotedDateReferenceFormulaWithMath));
				var formulaCell = worksheet.Cells[2, 4];
				formulaCell.Formula = dateFormulaWithMath;
				formulaCell.Calculate();
				Assert.AreEqual(expectedDateWithMath, formulaCell.Value);
				formulaCell.Formula = quotedDateReferenceFormulaWithMath;
				formulaCell.Calculate();
				Assert.AreEqual(expectedDateDecimalWithMath, formulaCell.Value);
			}
		}

		[TestMethod]
		public void CalculateWithInvalidDateValue()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = "YEAR(\"30/12/15\")";
				sheet.Calculate();
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void CalculateHandlesNestedBrackets()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("sheet");
				worksheet.Cells["C3"].Formula = "=\"\"\"[Cust - Bill-to].[b\"&\"y Country City].[Country]\"\"\"";
				worksheet.Cells["C3"].Calculate();
				Assert.AreEqual("\"[Cust - Bill-to].[by Country City].[Country]\"", worksheet.Cells["C3"].Value);
			}
		}
		#endregion

		#region Private Methods
		private string GetOutput(string file)
		{
			using (var pck = new ExcelPackage(new FileInfo(file)))
			{
				var fr = new Dictionary<string, object>();
				foreach (var ws in pck.Workbook.Worksheets)
				{
					if (!(ws is ExcelChartsheet))
					{
						foreach (var cell in ws.Cells)
						{
							if (!string.IsNullOrEmpty(cell.Formula))
							{
								fr.Add(ws.PositionID.ToString() + "," + cell.Address, cell.Value);
								ws.SetValueInner(cell.Start.Row, cell.Start.Column, null);
							}
						}
					}
				}

				pck.Workbook.Calculate();
				var nErrors = 0;
				var errors = new List<Tuple<string, object, object>>();
				ExcelWorksheet sheet = null;
				string adr = "";
				var fileErr = new System.IO.StreamWriter("c:\\temp\\err.txt");
				foreach (var cell in fr.Keys)
				{
					try
					{
						var spl = cell.Split(',');
						var ix = int.Parse(spl[0]);
						sheet = pck.Workbook.Worksheets[ix];
						adr = spl[1];
						if (fr[cell] is double && (sheet.Cells[adr].Value is double || sheet.Cells[adr].Value is decimal || sheet.Cells[adr].Value.GetType().IsPrimitive))
						{
							var d1 = Convert.ToDouble(fr[cell]);
							var d2 = Convert.ToDouble(sheet.Cells[adr].Value);
							//if (Math.Abs(d1 - d2) < double.Epsilon)
							if (double.Equals(d1, d2))
							{
								continue;
							}
							else
							{
								//errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
								fileErr.WriteLine("Diff cell " + sheet.Name + "!" + adr + "\t" + d1.ToString("R15") + "\t" + d2.ToString("R15"));
							}
						}
						else
						{
							if ((fr[cell] ?? "").ToString() != (sheet.Cells[adr].Value ?? "").ToString())
							{
								fileErr.WriteLine("String?  cell " + sheet.Name + "!" + adr + "\t" + (fr[cell] ?? "").ToString() + "\t" + (sheet.Cells[adr].Value ?? "").ToString());
							}
							//errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
						}
					}
					catch (Exception e)
					{
						fileErr.WriteLine("Exception cell " + sheet.Name + "!" + adr + "\t" + fr[cell].ToString() + "\t" + sheet.Cells[adr].Value + "\t" + e.Message);
						fileErr.WriteLine("***************************");
						fileErr.WriteLine(e.ToString());
						nErrors++;
					}
				}
				fileErr.Close();
				return nErrors.ToString();
			}
		}
		#endregion
	}
}
