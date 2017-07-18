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
			return package;
		}
		#endregion
	}
}
