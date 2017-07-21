/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.

* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class AverageIfsTests
	{
		#region AverageIfs Tests
		[TestMethod]
		public void AverageIfsWithVariedRangeArgumentsAndCriteriaAndConstantAverageRange()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 4;
				worksheet.Cells["D2"].Value = 2;
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(C2:D2,C3:D3,\"<>-10\")";
				worksheet.Cells["C3"].Value = 4;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(C2:D2,C4:D4,\">0\")";
				worksheet.Cells["C4"].Value = "4";
				worksheet.Cells["D4"].Value = 2;
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(C2:D2,C5:D5,\"<>-10\")";
				worksheet.Cells["C5"].Value = "4";
				worksheet.Cells["D5"].Value = 2;
				worksheet.Cells["B6"].Formula = "AVERAGEIFS(C2:D2,C6:D6,\"word\")";
				worksheet.Cells["C6"].Value = "word";
				worksheet.Cells["D6"].Value = 2;
				worksheet.Cells["B7"].Formula = "AVERAGEIFS(C2:D2,C7:D7,\"<>-10\")";
				worksheet.Cells["C7"].Value = "word";
				worksheet.Cells["D7"].Value = 2;
				worksheet.Cells["B8"].Formula = "AVERAGEIFS(C2:D2,C8:D8,\"TRUE\")";
				worksheet.Cells["C8"].Value = true;
				worksheet.Cells["D8"].Value = 2;
				worksheet.Cells["B9"].Formula = "AVERAGEIFS(C2:D2,C9:D9,\"<>-10\")";
				worksheet.Cells["C9"].Value = false;
				worksheet.Cells["D9"].Value = 2;
				worksheet.Cells["B10"].Formula = "AVERAGEIFS(C2:D2,C10:D10,\"<6/23/2017\")";
				worksheet.Cells["C10"].Value = (new DateTime(2017, 6, 22)).ToOADate();
				worksheet.Cells["D10"].Value = 2;
				worksheet.Cells["B11"].Formula = "AVERAGEIFS(C2:D2,C11:D11,\"<>-10\")";
				worksheet.Cells["C11"].Value = (new DateTime(2017, 6, 22)).ToOADate();
				worksheet.Cells["D11"].Value = 2;
				worksheet.Cells["B12"].Formula = "AVERAGEIFS(C2:D2,C12:D12,\"6/22/2017\")";
				worksheet.Cells["C12"].Value = "6/22/2017";
				worksheet.Cells["D12"].Value = 2;
				worksheet.Cells["B13"].Formula = "AVERAGEIFS(C2:D2,C13:D13,\"\")";
				worksheet.Cells["C13"].Value = null;
				worksheet.Cells["D13"].Value = 2;
				worksheet.Cells["B14"].Formula = "AVERAGEIFS(C2:D2,C14:D14,\"<>-10\")";
				worksheet.Cells["C14"].Value = null;
				worksheet.Cells["D14"].Value = 2;
				worksheet.Cells["B15"].Formula = "AVERAGEIFS(C2:D2,C15:D15,\"<>-10\")";
				worksheet.Cells["C15"].Formula = "notAValidFormula"; // Evaluates to #NAME.
				worksheet.Cells["D15"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(4d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(4d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B11"].Value);
				Assert.AreEqual(4d, worksheet.Cells["B12"].Value);
				Assert.AreEqual(4d, worksheet.Cells["B13"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B14"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B15"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithDifferentCriteriaInputs()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "Monday";
				worksheet.Cells["D2"].Value = "Tuesday";
				worksheet.Cells["E2"].Value = "Thursday";
				worksheet.Cells["F2"].Value = "Friday";
				worksheet.Cells["G2"].Value = "Thursday";
				worksheet.Cells["H2"].Value = 5;
				worksheet.Cells["I2"].Value = 2;
				worksheet.Cells["J2"].Value = 3.5;
				worksheet.Cells["K2"].Value = 6;
				worksheet.Cells["L2"].Value = 1;
				worksheet.Cells["M2"].Value = "5";
				worksheet.Cells["N2"].Value = true;
				worksheet.Cells["O2"].Value = "True";
				worksheet.Cells["P2"].Value = false;
				worksheet.Cells["Q2"].Value = (new DateTime(2017, 6, 22)).ToOADate();
				worksheet.Cells["R2"].Value = (new DateTime(2017, 6, 23)).ToOADate();
				worksheet.Cells["S2"].Value = (new DateTime(2017, 6, 24)).ToOADate();
				worksheet.Cells["T2"].Value = 0.0;
				worksheet.Cells["U2"].Value = 0.25;
				worksheet.Cells["V2"].Value = 0.5;
				worksheet.Cells["W2"].Value = 0.75;

				worksheet.Cells["C3"].Value = 1;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["E3"].Value = 3;
				worksheet.Cells["F3"].Value = 4;
				worksheet.Cells["G3"].Value = 7;
				worksheet.Cells["H3"].Value = 6;
				worksheet.Cells["I3"].Value = 7;
				worksheet.Cells["J3"].Value = 8;
				worksheet.Cells["K3"].Value = 9;
				worksheet.Cells["L3"].Value = 10;
				worksheet.Cells["M3"].Value = 11;
				worksheet.Cells["N3"].Value = 12;
				worksheet.Cells["O3"].Value = 13;
				worksheet.Cells["P3"].Value = 14;
				worksheet.Cells["Q3"].Value = 15;
				worksheet.Cells["R3"].Value = 16;
				worksheet.Cells["S3"].Value = 17;
				worksheet.Cells["T3"].Value = 18;
				worksheet.Cells["U3"].Value = 19;
				worksheet.Cells["V3"].Value = 20;
				worksheet.Cells["W3"].Value = 21;

				worksheet.Cells["C11"].Value = 1;
				worksheet.Cells["D11"].Value = 2;
				worksheet.Cells["E11"].Value = 1;
				worksheet.Cells["F11"].Value = 1;
				worksheet.Cells["G11"].Value = 2;
				worksheet.Cells["H11"].Value = 3;

				worksheet.Cells["B2"].Formula = "AVERAGEIFS(C3:W3,C2:W2,5)";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\"5\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\"=5\")";
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(C3:W3,C2:W2,3.5)";
				worksheet.Cells["B6"].Formula = "AVERAGEIFS(C3:W3,C2:W2,TRUE)";
				worksheet.Cells["B7"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\"TRUE\")";
				worksheet.Cells["B8"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\"6/23/2017\")";
				worksheet.Cells["B9"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\"6:00 PM\")";
				worksheet.Cells["B10"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\"T*sday\")";
				worksheet.Cells["B11"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\">=1\")";
				worksheet.Cells["B12"].Formula = "AVERAGEIFS(C3:W3,C2:W2,\">6/22/2017\")";
				worksheet.Cells["B13"].Formula = "AVERAGEIFS(F11:H11,C11:E11,1)";
				worksheet.Calculate();
				Assert.AreEqual(8.5, worksheet.Cells["B2"].Value);
				Assert.AreEqual(8.5, worksheet.Cells["B3"].Value);
				Assert.AreEqual(8.5, worksheet.Cells["B4"].Value);
				Assert.AreEqual(8d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(12d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(12d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(16d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(21d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(4d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(11d, worksheet.Cells["B11"].Value);
				Assert.AreEqual(16.5, worksheet.Cells["B12"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B13"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithVariedAverageRangeArgumentsAndConstantRangeAndCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(C3:D3,C2:D2,\">0\")";
				worksheet.Cells["C3"].Value = 4;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(C4:D4,C2:D2,\">0\")";
				worksheet.Cells["C4"].Value = "4";
				worksheet.Cells["D4"].Value = 2;
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(C5:D5,C2:D2,\">0\")";
				worksheet.Cells["C5"].Value = "word";
				worksheet.Cells["D5"].Value = 2;
				worksheet.Cells["B6"].Formula = "AVERAGEIFS(C6:D6,C2:D2,\">0\")";
				worksheet.Cells["C6"].Value = true;
				worksheet.Cells["D6"].Value = 2;
				worksheet.Cells["B7"].Formula = "AVERAGEIFS(C7:D7,C2:D2,\">0\")";
				worksheet.Cells["C7"].Value = (new DateTime(2017, 6, 22)).ToOADate();
				worksheet.Cells["D7"].Value = 2;
				worksheet.Cells["B8"].Formula = "AVERAGEIFS(C8:D8,C2:D2,\">0\")";
				worksheet.Cells["C8"].Value = "6/22/2017";
				worksheet.Cells["D8"].Value = 2;
				worksheet.Cells["B9"].Formula = "AVERAGEIFS(C9:D9,C2:D2,\">0\")";
				worksheet.Cells["C9"].Value = null;
				worksheet.Cells["D9"].Value = 2;
				worksheet.Cells["B10"].Formula = "AVERAGEIFS(C10:D10,C2:D2,\">0\")";
				worksheet.Cells["C10"].Formula = "notAValidFormula"; // Evaluates to #NAME.
				worksheet.Cells["D10"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(21455d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void AverageIfsCriteriaIgnoresCase()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "aa";
				worksheet.Cells["D2"].Value = "ab";
				worksheet.Cells["C3"].Value = 3;
				worksheet.Cells["D3"].Value = 1;
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(C3:D3,C2:D2,\"?b\")";
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(C3:D3,C2:D2,\"*B\")";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithCriteriaAsCellRange()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "aa";
				worksheet.Cells["D2"].Value = "ab";
				worksheet.Cells["E2"].Value = 3;
				worksheet.Cells["F2"].Value = 1;
				worksheet.Cells["G2"].Value = "*b";
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(E2:F2,C2:D2,G2)";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(E2:F2,C2:D2,C2:D2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void AverageIfsWithRangeAsCellRangeAndCriteriaAsArray()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["D2"].Value = 5;
				worksheet.Cells["B3"].Formula = "{1,2,3}";
				worksheet.Cells["C3"].Formula = "{1,2,3}";
				worksheet.Cells["D3"].Formula = "{1,2,3}";
				worksheet.Cells["D4"].Formula = "AVERAGEIFS(B2:D2,B3:D3,{1,2,3})";
				worksheet.Cells["D5"].Formula = "AVERAGEIFS(B2:D2,B3:D3,B3)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["D4"].Value);
				Assert.AreEqual(3d, worksheet.Cells["D5"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsEmptyStringCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = null;
				worksheet.Cells["C3"].Value = string.Empty;
				worksheet.Cells["C4"].Value = "Not Empty";
				worksheet.Cells["D2"].Value = 1d;
				worksheet.Cells["D3"].Value = 3d;
				worksheet.Cells["D4"].Value = 5d;
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2:D4,C2:C4,\"\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithInequalitiesOnBooleanCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = true;
				worksheet.Cells["C3"].Value = false;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2,C2,\">TRUE\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(D2,C3,\"<TRUE\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(D2,C2,\">FALSE\")";
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(D2,C3,\">TRUE\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void AverageIfsWithCriteriaAsExpressionCharacter()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "=";
				worksheet.Cells["C3"].Value = "";
				worksheet.Cells["C4"].Value = null;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2,C2,\"=\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(D2,C3,\"=\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(D2,C4,\"=\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithInequalitiesOnStrings()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "ay";
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2,C2,\"<axz\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(D2,C2,\"<aya\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(D2,C2,\"<az\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithTextComparisonsWithWildcardCharacter()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "ay";
				worksheet.Cells["C3"].Value = "Modday";
				worksheet.Cells["C4"].Value = "Monnnnday";
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["D3"].Value = 3;
				worksheet.Cells["D4"].Value = 5;

				worksheet.Cells["B"].Formula = "AVERAGEIFS(D2:D4,C2:C4,\"=Mo*day\")";

				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2:D4,C2:C4,\"=Mo*day\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(D2:D4,C2:C4,\">Mo*day\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(D2:D4,C2:C4,\"<Mo*day\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(4d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithTextComparisonsWithEscapedWildcardCharacter()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "Mon?ay";
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["C3"].Value = "Monday";
				worksheet.Cells["D3"].Value = 3;
				worksheet.Cells["C4"].Value = "Mon*ay";
				worksheet.Cells["D4"].Value = 5;
				worksheet.Cells["C5"].Value = "Monddday";
				worksheet.Cells["D5"].Value = 7;
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2:D3,C2:C3,\"Mon?ay\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(D2:D3,C2:C3,\"Mon~?ay\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(D4:D5,C4:C5,\"Mon*ay\")";
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(D4:D5,C4:C5,\"Mon~*ay\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(6d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(5d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithStringInequalityCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = null;
				worksheet.Cells["C3"].Value = "";
				worksheet.Cells["C4"].Value = "zzz";
				worksheet.Cells["C5"].Value = 1;
				worksheet.Cells["C6"].Value = "1";
				worksheet.Cells["C7"].Value = true;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(D2,C2,\">a\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(D2,C2,\"<a\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(D2,C3,\">a\")";
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(D2,C3,\"<a\")";
				worksheet.Cells["B6"].Formula = "AVERAGEIFS(D2,C4,\">a\")";
				worksheet.Cells["B7"].Formula = "AVERAGEIFS(D2,C4,\"<a\")";
				worksheet.Cells["B8"].Formula = "AVERAGEIFS(D2,C5,\">a\")";
				worksheet.Cells["B9"].Formula = "AVERAGEIFS(D2,C5,\"<a\")";
				worksheet.Cells["B10"].Formula = "AVERAGEIFS(D2,C6,\">a\")";
				worksheet.Cells["B11"].Formula = "AVERAGEIFS(D2,C6,\"<a\")";
				worksheet.Cells["B12"].Formula = "AVERAGEIFS(D2,C7,\">a\")";
				worksheet.Cells["B13"].Formula = "AVERAGEIFS(D2,C7,\"<a\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B7"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B8"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B9"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B11"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B12"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B13"].Value).Type);
			}
		}

		[TestMethod]
		public void AverageIfsWithVariedExpressionCharacters()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C8"].Value = ">1";
				worksheet.Cells["C9"].Value = "<1";
				worksheet.Cells["C10"].Value = "=1";

				worksheet.Cells["B"].Formula = "AVERAGEIFS(C2,C,\"1\")";

				worksheet.Cells["B2"].Formula = "AVERAGEIFS(C2,C2,\"=1\")";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(C2,C2,\">1\")";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(C2,C2,\"<1\")";
				worksheet.Cells["B5"].Formula = "AVERAGEIFS(C2,C2,\">=1\")";
				worksheet.Cells["B6"].Formula = "AVERAGEIFS(C2,C2,\"<=1\")";
				worksheet.Cells["B7"].Formula = "AVERAGEIFS(C2,C2,\"<>1\")";
				worksheet.Cells["B8"].Formula = "AVERAGEIFS(C2,C8,\"=>1\")";
				worksheet.Cells["B9"].Formula = "AVERAGEIFS(C2,C9,\"=<1\")";
				worksheet.Cells["B10"].Formula = "AVERAGEIFS(C2,C10,\"==1\")";
				worksheet.Cells["B11"].Formula = "AVERAGEIFS(C2,C2,\">>1\")";
				worksheet.Cells["B12"].Formula = "AVERAGEIFS(C2,C2,\"><1\")";
				worksheet.Cells["B13"].Formula = "AVERAGEIFS(C2,C2,\"<<1\")";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B7"].Value).Type);
				Assert.AreEqual(1d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B11"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B12"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B13"].Value).Type);
			}
		}

		[TestMethod]
		public void AverageIfsWithDifferentRangeSizesReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(E2,C2:D4,\">0\")";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 2;
				worksheet.Cells["C4"].Value = 3;
				worksheet.Cells["D2"].Value = 4;
				worksheet.Cells["D3"].Value = 5;
				worksheet.Cells["D4"].Value = 6;
				worksheet.Cells["E2"].Value = 1;
				worksheet.Cells["E3"].Value = 3;
				worksheet.Cells["E4"].Value = 5;
				worksheet.Cells["F2"].Value = 7;
				worksheet.Cells["F3"].Value = 9;
				worksheet.Cells["F4"].Value = 11;
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void AverageIfsWithMultipleDifferentCriterias()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(C2:C4,D2:D4,\"<>2\",E2:E4,true)";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 7;
				worksheet.Cells["C4"].Value = 5;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["D4"].Value = 3;
				worksheet.Cells["E2"].Value = false;
				worksheet.Cells["E3"].Value = false;
				worksheet.Cells["E4"].Value = true;
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithNullCriteriaReturnsDiv0()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(C2:C3,D2:D3,1)";
				worksheet.Cells["B3"].Formula = "AVERAGEIFS(C2:C3,D2:D3,1,E2:E3,)";
				worksheet.Cells["B4"].Formula = "AVERAGEIFS(C2:C3,D2:D3,1,E2:E3,E3)";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 3;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["D3"].Value = 1;
				worksheet.Cells["E2"].Value = 0;
				worksheet.Cells["E3"].Value = null;
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void AverageIfsWithUnsetEmptyCellsInCriteria()
		{
			// This test exists to ensure that cells that have never been set are still 
			// being compared against the criterion.
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "AVERAGEIFS(C2:C3,D2:D3,\"\")";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 3;
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}
		#endregion
	}
}
