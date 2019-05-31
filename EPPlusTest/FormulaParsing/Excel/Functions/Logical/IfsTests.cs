/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Evan Schallerer and others as noted in the source history.
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
using System.Collections.Generic;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
	[TestClass]
	public class IfsTests
	{
		#region Class Variables
		private ParsingContext _parsingContext = ParsingContext.Create();
		#endregion

		#region TestMethods
		[TestMethod]
		public void IfsFunctionSingleConditionTrue()
		{
			string expected = "ASDAS";
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs(true, expected);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expected, (string)result.Result);
		}

		[TestMethod]
		public void IfsFunctionSingleConditionTrueString()
		{
			string expected = "ASDAS";
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs("true", expected);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expected, (string)result.Result);
		}

		[TestMethod]
		public void IfsFunctionSingleConditionFalse()
		{
			string expected = "ASDAS";
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs(false, expected);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfsFunctionZeroArguments()
		{
			var func = new Ifs();
			var result = func.Execute(new List<FunctionArgument>(), _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfsFunctionOddArgumentCount()
		{
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 5);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion

		#region Integration Tests
		[TestMethod]
		public void IfsFunctionIntegrationTrueTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"IFS(2=1, ""badbadbad"", 2=2, ""good"")";
				sheet.Calculate();
				Assert.AreEqual("good", sheet.Cells[2, 2].Value);
			}
		}
		#endregion
	}
}
