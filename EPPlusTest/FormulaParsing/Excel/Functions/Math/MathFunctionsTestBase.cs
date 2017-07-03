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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public abstract class MathFunctionsTestBase
	{
		#region Properties
		protected ParsingContext ParsingContext { get; } = ParsingContext.Create();
		#endregion

		#region Protected Methods
		protected ExcelPackage CreateTestingPackage()
		{
			var package = new ExcelPackage();
			var worksheet = package.Workbook.Worksheets.Add("Sheet1");
			worksheet.Cells["A2"].Value = 1;
			worksheet.Cells["A3"].Value = 2;
			worksheet.Cells["A4"].Value = 1;
			worksheet.Cells["A5"].Value = 1;
			worksheet.Cells["A6"].Value = 2;
			worksheet.Cells["A7"].Value = 3;
			worksheet.Cells["B1"].Value = "Monday";
			worksheet.Cells["B2"].Value = "Tuesday";
			worksheet.Cells["B3"].Value = "Thursday";
			worksheet.Cells["B4"].Value = "Friday";
			worksheet.Cells["B5"].Value = "Thursday";
			worksheet.Cells["B6"].Value = "5";
			worksheet.Cells["B7"].Value = "2";
			worksheet.Cells["B8"].Value = "3.5";
			worksheet.Cells["B9"].Value = "6";
			worksheet.Cells["B10"].Value = "1";
			worksheet.Cells["B11"].Value = "\"5\"";
			worksheet.Cells["B12"].Value = true;
			worksheet.Cells["B13"].Value = "TRUE";
			worksheet.Cells["B14"].Value = false;
			worksheet.Cells["B15"].Formula = "DATE(2017, 6, 22)";
			worksheet.Cells["B16"].Formula = "DATE(2017, 6, 23)";
			worksheet.Cells["B17"].Formula = "DATE(2017, 6, 24)";
			worksheet.Cells["B18"].Value = "12:00:00 AM";
			worksheet.Cells["B19"].Value = "6:00:00 AM";
			worksheet.Cells["B20"].Value = "12:00:00 PM";
			worksheet.Cells["B21"].Value = "6:00:00 PM";
			worksheet.Cells["C1"].Value = 1;
			worksheet.Cells["C2"].Value = 2;
			worksheet.Cells["C3"].Value = 3;
			worksheet.Cells["C4"].Value = 4;
			worksheet.Cells["C5"].Value = 5;
			worksheet.Cells["C6"].Value = 6;
			worksheet.Cells["C7"].Value = 7;
			worksheet.Cells["C8"].Value = 8;
			worksheet.Cells["C9"].Value = 9;
			worksheet.Cells["C10"].Value = 10;
			worksheet.Cells["C11"].Value = 11;
			worksheet.Cells["C12"].Value = 12;
			worksheet.Cells["C13"].Value = 13;
			worksheet.Cells["C14"].Value = 14;
			worksheet.Cells["C15"].Value = 15;
			worksheet.Cells["C16"].Value = 16;
			worksheet.Cells["C17"].Value = 17;
			worksheet.Cells["C18"].Value = 18;
			worksheet.Cells["C19"].Value = 19;
			worksheet.Cells["C20"].Value = 20;
			worksheet.Cells["C21"].Value = 21;
			return package;
		}
		#endregion
	}
}
