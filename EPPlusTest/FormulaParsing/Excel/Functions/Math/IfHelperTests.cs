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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class IfHelperTests : MathFunctionsTestBase 
	{
		private ExcelPackage _package;
		private EpplusExcelDataProvider _provider;
		private ParsingContext _parsingContext;
		private ExcelWorksheet _worksheet;

		[TestInitialize]
		public void Initialize()
		{
			_package = new ExcelPackage();
			_provider = new EpplusExcelDataProvider(_package);
			_parsingContext = ParsingContext.Create();
			_parsingContext.Scopes.NewScope(RangeAddress.Empty);
			_worksheet = _package.Workbook.Worksheets.Add("TestSheet");
		}

		[TestCleanup]
		public void Cleanup()
		{
			_package.Dispose();
		}

		[TestMethod]
		public void CalculateCriteriaWithCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var provider = new EpplusExcelDataProvider(package);
				var worksheet = package.Workbook.Worksheets.Add("Sheet2");
				this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);

				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 10;
				worksheet.Cells["B3"].Value = 15;

				IRangeInfo testRange = provider.GetRange(worksheet.Name, 1, 1, 3, 1);
				IRangeInfo secondRnage = provider.GetRange(worksheet.Name, 2, 2, 2, 2);

				var res = IfHelper.CalculateCriteria(FunctionsHelper.CreateArgs(secondRnage, testRange), this.ParsingContext);
				Assert.AreEqual(3, res);
			}

		}
	}
}
