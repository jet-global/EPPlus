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
	public class MedianTests : MathFunctionsTestBase
	{
		#region Median Function (Execute) Tests

		[TestMethod]
		public void MedianWithNoArgumentsReturnsPoundValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MedianWithMaxArgumentsReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithOneInputReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithNumericInputsReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithGenericStringInputReturnsPoundValue()
		{

		}

		[TestMethod]
		public void MedianWithNumericStringInputReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithReferenceToNumbersReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithReferencesTypedOutReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithReferenceToNumericStringsReturnsPoundNum()
		{

		}

		[TestMethod]
		public void MedianWithReferencesToGeneralStringsReturnsPoundNum()
		{

		}

		[TestMethod]
		public void MedianWithLogicInputsReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithReferenceToLogicInputsReturnsPoundNum()
		{

		}

		[TestMethod]
		public void MedianWithReferenceToCellsWithZeroReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithReferenceToEmptyCellsReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithReferenceToStringsAndNumbersReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithDateObjectAsInputReturnsCorrectValue()
		{

		}
		
		[TestMethod]
		public void MedianWithDateAsStringReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithDoubleInputsReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithFractionInputReturnsCorrectValue()
		{

		}
		#endregion
	}
}
