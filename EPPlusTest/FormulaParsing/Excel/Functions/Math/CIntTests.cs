﻿/*******************************************************************************
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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class CIntTests : MathFunctionsTestBase
	{
		#region CInt Function (Execute) Tests
		[TestMethod]
		public void CIntShouldConvertTextToInteger()
		{
			var func = new CInt();
			var args = FunctionsHelper.CreateArgs("2");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void CIntithInvalidArgumentReturnsPoundValue()
		{
			var func = new CInt();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CIntShouldConvertDecimalToInteger()
		{
			var func = new CInt();
			var args = FunctionsHelper.CreateArgs(2.88m);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void CIntShouldConvertNegativeDecimalToInteger()
		{
			var func = new CInt();
			var args = FunctionsHelper.CreateArgs(-2.88m);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(-3, result.Result);
		}

		[TestMethod]
		public void CIntShouldConvertStringToInteger()
		{
			var func = new CInt();
			var args = FunctionsHelper.CreateArgs("-2.88");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(-3, result.Result);
		}
		#endregion
	}
}
