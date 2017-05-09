/* Copyright (C) 2011  Jan Källman
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	[TestClass]
	public class IndexTests
	{
		#region Properties
		private ParsingContext ParsingContext { get; set; }
		private ExcelPackage Package { get; set; }
		private ExcelWorksheet Worksheet { get; set; }
		#endregion

		#region TestInitialize/TestCleanup
		[TestInitialize]
		public void Initialize()
		{
			this.ParsingContext = ParsingContext.Create();
			this.Package = new ExcelPackage(new MemoryStream());
			this.Worksheet = this.Package.Workbook.Worksheets.Add("test");
		}

		[TestCleanup]
		public void Cleanup()
		{
			this.Package.Dispose();
		}
		#endregion

		#region Index Tests
		[TestMethod]
		public void IndexReturnsPoundValueWhenTooFewArgumentsAreSupplied()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Formula = "INDEX(A1:A3,213)";
			this.Worksheet.Calculate();
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), this.Worksheet.Cells["A4"].Value);
		}

		[TestMethod]
		public void IndexReturnsValueByIndex()
		{
			var func = new Index();
			var result = func.Execute(FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 5), 3), this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void IndexHandlesSingleRange()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Formula = "INDEX(A1:A3,3)";
			this.Worksheet.Calculate();
			Assert.AreEqual(5d, this.Worksheet.Cells["A4"].Value);
		}

		[TestMethod]
		public void IndexHandlesNAError()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Value = ExcelErrorValue.Create(eErrorType.NA);
			this.Worksheet.Cells["A5"].Formula = "INDEX(A1:A3,A4)";
			this.Worksheet.Calculate();
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), this.Worksheet.Cells["A5"].Value);
		}
		#endregion
	}
}
