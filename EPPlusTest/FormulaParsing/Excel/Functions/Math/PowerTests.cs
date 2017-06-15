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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class PowerTests : MathFunctionsTestBase
	{
		#region Power Function (Execute) Tests
		[TestMethod]
		public void PowerWithPositiveIntegerArgumentsReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, 2), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void PowerWithZeroAsBaseReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, 5), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void PowerWithZeroAsPowerReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, 0), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void PowerWithPositiveBaseAndNegativePowerReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, -3), this.ParsingContext);
			Assert.AreEqual(0.125d, result.Result);
		}

		[TestMethod]
		public void PowerWithNegativeBaseAndEvenPowerReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(-2, 2), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void PowerWithNegativeBaseAndOddPowerReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(-2, 3), this.ParsingContext);
			Assert.AreEqual(-8d, result.Result);
		}

		[TestMethod]
		public void PowerWithNegativeBaseAndNegativePowerReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(-2, -5), this.ParsingContext);
			Assert.AreEqual(-0.03125d, (double)result.Result, 0.00001);
		}

		[TestMethod]
		public void	PowerWithFractionBaseReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "POWER((2/3), 5)";
				ws.Calculate();
				Assert.AreEqual(0.131687243d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void PowerWithFractionPowerReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "POWER(2, (1/5))";
				ws.Calculate();
				Assert.AreEqual(1.148698355d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void PowerWithFractionBaseAndPowerReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "POWER((2/3), (1/5))";
				ws.Calculate();
				Assert.AreEqual(0.922107911d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void PowerWithBaseAsResultOfDateFunctionReturnsCorrectValue()
		{
			var function = new Power();
			var dateArgument = new DateTime(2017, 5, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(dateArgument, 2), this.ParsingContext);
			Assert.AreEqual(1836636736d, result.Result);
		}

		[TestMethod]
		public void PowerWithPowerAsResultOfDateFunctionReturnsPoundNum()
		{
			var function = new Power();
			var dateArgument = new DateTime(2017, 5, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(2, dateArgument), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void PowerWithBaseDateAsStringReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/1/2017", 2), this.ParsingContext);
			Assert.AreEqual(1836636736d, result.Result);
		}

		[TestMethod]
		public void PowerWithPowerDateAsStringReturnsPoundNum()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, "5/1/2017"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void PowerWithDoubleAsBaseReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(2.3, 2), this.ParsingContext);
			Assert.AreEqual(5.29d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void PowerWithDoubleAsPowerReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(3, 5.5), this.ParsingContext);
			Assert.AreEqual(420.8883462d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void PowerWithGeneralStringAsBaseReturnsPoundValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void PowerWithGeneralStringAsPowerReturnsPoundValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void PowerWithNullFirstArgumentReturnsZero()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 2), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void PowerWithNullOrMissingSecondArgumentReturnsOne()
		{
			var function = new Power();
			var resultWithNullArg = function.Execute(FunctionsHelper.CreateArgs(2, null), this.ParsingContext);
			Assert.AreEqual(1d, resultWithNullArg.Result);
			var resultWithMissingArg = function.Execute(FunctionsHelper.CreateArgs(2), this.ParsingContext);
			Assert.AreEqual(1d, resultWithMissingArg.Result);
		}

		[TestMethod]
		public void PowerWithArgumentsAsStringsReturnsCorrectValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs("2", "3"), this.ParsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void PowerWithNoArgumentsReturnsPoundValue()
		{
			var function = new Power();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void PowerWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Power();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void PowerFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Power();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),1);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),1);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),1);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),1);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),1);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),1);
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}
		#endregion
	}
}
