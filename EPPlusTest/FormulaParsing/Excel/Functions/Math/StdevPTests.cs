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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{

	[TestClass]
	public class StdevPTests : MathFunctionsTestBase
	{
		#region StddevP Function(Execute) Tests

		[TestMethod]
		public void StddevPIsGivenBooleanInputs()
		{
			var function = new StdevP();
			var boolInputTrue = true;
			var boolInputFalse = false;
			var result1 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputFalse), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputTrue), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputFalse, boolInputFalse), this.ParsingContext);

			Assert.AreEqual(0.471404521, result1.ResultNumeric, .00001);
			Assert.AreEqual(0, result2.ResultNumeric, .00001);
			Assert.AreEqual(0.471404521, result3.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StddevPIsGivenBooleanInputsFromCellRefrence()
		{
			var function = new StdevP();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Formula = "=stdev.p(B1,B1,B2)";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B1,B1)";
				worksheet.Cells["B5"].Formula = "=stdev.p(B1,B2,B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, worksheet.Cells["B3"].Value);
				Assert.AreEqual(eErrorType.Div0, worksheet.Cells["B4"].Value);
				Assert.AreEqual(eErrorType.Div0, worksheet.Cells["B5"].Value);
			}
		}




		#endregion
	}
}
