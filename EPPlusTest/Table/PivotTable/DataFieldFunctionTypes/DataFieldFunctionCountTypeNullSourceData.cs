/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Michelle Lau and others as noted in the source history.
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
using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable.DataFieldFunctionTypes
{
	[TestClass]
	public class DataFieldFunctionCountTypeNullSourceData
	{
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableDataCalculationCountType.xlsx")]
		public void PivotTableRefreshCountFunctionTypeOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableDataCalculationCountType.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:D11"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Count of Currency Code"),
					new ExpectedCellValue(sheetName, 2, 1, "Row Labels"),
					new ExpectedCellValue(sheetName, 3, 1, "Autohaus Mielberg KG 2018"),
					new ExpectedCellValue(sheetName, 4, 1, "Beef House 2018"),
					new ExpectedCellValue(sheetName, 5, 1, "Credit Memo 104001"),
					new ExpectedCellValue(sheetName, 6, 1, "Opening Entries, Customers"),
					new ExpectedCellValue(sheetName, 7, 1, "Order 101008"),
					new ExpectedCellValue(sheetName, 8, 1, "Order 101014"),
					new ExpectedCellValue(sheetName, 9, 1, "Order 101021"),
					new ExpectedCellValue(sheetName, 10, 1, "Payment 2018"),
					new ExpectedCellValue(sheetName, 11, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, "Column Labels"),
					new ExpectedCellValue(sheetName, 2, 2, "Jan"),
					new ExpectedCellValue(sheetName, 3, 2, 4d),
					new ExpectedCellValue(sheetName, 4, 2, 1d),
					new ExpectedCellValue(sheetName, 5, 2, 2d),
					new ExpectedCellValue(sheetName, 6, 2, null),
					new ExpectedCellValue(sheetName, 7, 2, 1d),
					new ExpectedCellValue(sheetName, 8, 2, 1d),
					new ExpectedCellValue(sheetName, 9, 2, 1d),
					new ExpectedCellValue(sheetName, 10, 2, 2d),
					new ExpectedCellValue(sheetName, 11, 2, 12d),
					new ExpectedCellValue(sheetName, 1, 3, null),
					new ExpectedCellValue(sheetName, 2, 3, "Dec"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 6, 3, 3d),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 11, 3, 3d),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 4, 4d),
					new ExpectedCellValue(sheetName, 4, 4, 1d),
					new ExpectedCellValue(sheetName, 5, 4, 2d),
					new ExpectedCellValue(sheetName, 6, 4, 3d),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 1d),
					new ExpectedCellValue(sheetName, 9, 4, 1d),
					new ExpectedCellValue(sheetName, 10, 4, 2d),
					new ExpectedCellValue(sheetName, 11, 4, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableDataCalculationCountType.xlsx")]
		public void PivotTableRefreshCountFunctionTypeTwoRowFieldsTwoColumnFieldsLeafDataField()
		{
			var file = new FileInfo("PivotTableDataCalculationCountType.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A15:H27"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 15, 1, null),
					new ExpectedCellValue(sheetName, 16, 1, null),
					new ExpectedCellValue(sheetName, 17, 1, "Description"),
					new ExpectedCellValue(sheetName, 18, 1, "Autohaus Mielberg KG 2018"),
					new ExpectedCellValue(sheetName, 19, 1, "Beef House 2018"),
					new ExpectedCellValue(sheetName, 20, 1, "Credit Memo 104001"),
					new ExpectedCellValue(sheetName, 21, 1, "Opening Entries, Customers"),
					new ExpectedCellValue(sheetName, 22, 1, null),
					new ExpectedCellValue(sheetName, 23, 1, "Order 101008"),
					new ExpectedCellValue(sheetName, 24, 1, "Order 101014"),
					new ExpectedCellValue(sheetName, 25, 1, "Order 101021"),
					new ExpectedCellValue(sheetName, 26, 1, "Payment 2018"),
					new ExpectedCellValue(sheetName, 27, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 2, null),
					new ExpectedCellValue(sheetName, 16, 2, null),
					new ExpectedCellValue(sheetName, 17, 2, "Customer No."),
					new ExpectedCellValue(sheetName, 18, 2, 49633663),
					new ExpectedCellValue(sheetName, 19, 2, 49525252),
					new ExpectedCellValue(sheetName, 20, 2, 10000),
					new ExpectedCellValue(sheetName, 21, 2, 10000),
					new ExpectedCellValue(sheetName, 22, 2, 30000),
					new ExpectedCellValue(sheetName, 23, 2, 35451236),
					new ExpectedCellValue(sheetName, 24, 2, 47563218),
					new ExpectedCellValue(sheetName, 25, 2, 35963852),
					new ExpectedCellValue(sheetName, 26, 2, 10000),
					new ExpectedCellValue(sheetName, 27, 2, null),
					new ExpectedCellValue(sheetName, 15, 3, "Posting Date"),
					new ExpectedCellValue(sheetName, 16, 3, "Jan"),
					new ExpectedCellValue(sheetName, 17, 3, "Count of Currency Code"),
					new ExpectedCellValue(sheetName, 18, 3, 4d),
					new ExpectedCellValue(sheetName, 19, 3, 1d),
					new ExpectedCellValue(sheetName, 20, 3, 2d),
					new ExpectedCellValue(sheetName, 21, 3, null),
					new ExpectedCellValue(sheetName, 22, 3, null),
					new ExpectedCellValue(sheetName, 23, 3, 1d),
					new ExpectedCellValue(sheetName, 24, 3, 1d),
					new ExpectedCellValue(sheetName, 25, 3, 1d),
					new ExpectedCellValue(sheetName, 26, 3, 2d),
					new ExpectedCellValue(sheetName, 27, 3, 12d),
					new ExpectedCellValue(sheetName, 15, 4, "Values"),
					new ExpectedCellValue(sheetName, 16, 4, null),
					new ExpectedCellValue(sheetName, 17, 4, "Sum of Profit ($)"),
					new ExpectedCellValue(sheetName, 18, 4, 0),
					new ExpectedCellValue(sheetName, 19, 4, 0),
					new ExpectedCellValue(sheetName, 20, 4, -129.98),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 22, 4, null),
					new ExpectedCellValue(sheetName, 23, 4, 259.97),
					new ExpectedCellValue(sheetName, 24, 4, 6349.7),
					new ExpectedCellValue(sheetName, 25, 4, 521.17),
					new ExpectedCellValue(sheetName, 26, 4, 0),
					new ExpectedCellValue(sheetName, 27, 4, 7000.86),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 16, 5, "Dec"),
					new ExpectedCellValue(sheetName, 17, 5, "Count of Currency Code"),
					new ExpectedCellValue(sheetName, 18, 5, null),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 20, 5, null),
					new ExpectedCellValue(sheetName, 21, 5, 2d),
					new ExpectedCellValue(sheetName, 22, 5, 1d),
					new ExpectedCellValue(sheetName, 23, 5, null),
					new ExpectedCellValue(sheetName, 24, 5, null),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, 3d),
					new ExpectedCellValue(sheetName, 15, 6, null),
					new ExpectedCellValue(sheetName, 16, 6, null),
					new ExpectedCellValue(sheetName, 17, 6, "Sum of Profit ($)"),
					new ExpectedCellValue(sheetName, 18, 6, null),
					new ExpectedCellValue(sheetName, 19, 6, null),
					new ExpectedCellValue(sheetName, 20, 6, null),
					new ExpectedCellValue(sheetName, 21, 6, 0),
					new ExpectedCellValue(sheetName, 22, 6, 0),
					new ExpectedCellValue(sheetName, 23, 6, null),
					new ExpectedCellValue(sheetName, 24, 6, null),
					new ExpectedCellValue(sheetName, 25, 6, null),
					new ExpectedCellValue(sheetName, 26, 6, null),
					new ExpectedCellValue(sheetName, 27, 6, 0),
					new ExpectedCellValue(sheetName, 15, 7, null),
					new ExpectedCellValue(sheetName, 16, 7, "Total Count of Currency Code"),
					new ExpectedCellValue(sheetName, 17, 7, null),
					new ExpectedCellValue(sheetName, 18, 7, 4d),
					new ExpectedCellValue(sheetName, 19, 7, 1d),
					new ExpectedCellValue(sheetName, 20, 7, 2d),
					new ExpectedCellValue(sheetName, 21, 7, 2d),
					new ExpectedCellValue(sheetName, 22, 7, 1d),
					new ExpectedCellValue(sheetName, 23, 7, 1d),
					new ExpectedCellValue(sheetName, 24, 7, 1d),
					new ExpectedCellValue(sheetName, 25, 7, 1d),
					new ExpectedCellValue(sheetName, 26, 7, 2d),
					new ExpectedCellValue(sheetName, 27, 7, 15d),
					new ExpectedCellValue(sheetName, 15, 8, null),
					new ExpectedCellValue(sheetName, 16, 8, "Total Sum of Profit ($)"),
					new ExpectedCellValue(sheetName, 17, 8, null),
					new ExpectedCellValue(sheetName, 18, 8, 0),
					new ExpectedCellValue(sheetName, 19, 8, 0),
					new ExpectedCellValue(sheetName, 20, 8, -129.98),
					new ExpectedCellValue(sheetName, 21, 8, 0),
					new ExpectedCellValue(sheetName, 22, 8, 0),
					new ExpectedCellValue(sheetName, 23, 8, 259.97),
					new ExpectedCellValue(sheetName, 24, 8, 6349.7),
					new ExpectedCellValue(sheetName, 25, 8, 521.17),
					new ExpectedCellValue(sheetName, 26, 8, 0),
					new ExpectedCellValue(sheetName, 27, 8, 7000.86)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableDataCalculationCountType.xlsx")]
		public void PivotTableRefreshCountFunctionTypeTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableDataCalculationCountType.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A31:E51"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 1, "Count of Currency Code"),
					new ExpectedCellValue(sheetName, 32, 1, null),
					new ExpectedCellValue(sheetName, 33, 1, "Description"),
					new ExpectedCellValue(sheetName, 34, 1, "Autohaus Mielberg KG 2018"),
					new ExpectedCellValue(sheetName, 35, 1, null),
					new ExpectedCellValue(sheetName, 36, 1, "Beef House 2018"),
					new ExpectedCellValue(sheetName, 37, 1, null),
					new ExpectedCellValue(sheetName, 38, 1, "Credit Memo 104001"),
					new ExpectedCellValue(sheetName, 39, 1, null),
					new ExpectedCellValue(sheetName, 40, 1, "Opening Entries, Customers"),
					new ExpectedCellValue(sheetName, 41, 1, null),
					new ExpectedCellValue(sheetName, 42, 1, null),
					new ExpectedCellValue(sheetName, 43, 1, "Order 101008"),
					new ExpectedCellValue(sheetName, 44, 1, null),
					new ExpectedCellValue(sheetName, 45, 1, "Order 101014"),
					new ExpectedCellValue(sheetName, 46, 1, null),
					new ExpectedCellValue(sheetName, 47, 1, "Order 101021"),
					new ExpectedCellValue(sheetName, 48, 1, null),
					new ExpectedCellValue(sheetName, 49, 1, "Payment 2018"),
					new ExpectedCellValue(sheetName, 50, 1, null),
					new ExpectedCellValue(sheetName, 51, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 32, 2, null),
					new ExpectedCellValue(sheetName, 33, 2, "Customer No."),
					new ExpectedCellValue(sheetName, 34, 2, null),
					new ExpectedCellValue(sheetName, 35, 2, 49633663),
					new ExpectedCellValue(sheetName, 36, 2, null),
					new ExpectedCellValue(sheetName, 37, 2, 49525252),
					new ExpectedCellValue(sheetName, 38, 2, null),
					new ExpectedCellValue(sheetName, 39, 2, 10000),
					new ExpectedCellValue(sheetName, 40, 2, null),
					new ExpectedCellValue(sheetName, 41, 2, 10000),
					new ExpectedCellValue(sheetName, 42, 2, 30000),
					new ExpectedCellValue(sheetName, 43, 2, null),
					new ExpectedCellValue(sheetName, 44, 2, 35451236),
					new ExpectedCellValue(sheetName, 45, 2, null),
					new ExpectedCellValue(sheetName, 46, 2, 47563218),
					new ExpectedCellValue(sheetName, 47, 2, null),
					new ExpectedCellValue(sheetName, 48, 2, 35963852),
					new ExpectedCellValue(sheetName, 49, 2, null),
					new ExpectedCellValue(sheetName, 50, 2, 10000),
					new ExpectedCellValue(sheetName, 51, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Years"),
					new ExpectedCellValue(sheetName, 32, 3, 2017),
					new ExpectedCellValue(sheetName, 33, 3, "Dec"),
					new ExpectedCellValue(sheetName, 34, 3, null),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 40, 3, 3d),
					new ExpectedCellValue(sheetName, 41, 3, 2d),
					new ExpectedCellValue(sheetName, 42, 3, 1d),
					new ExpectedCellValue(sheetName, 43, 3, null),
					new ExpectedCellValue(sheetName, 44, 3, null),
					new ExpectedCellValue(sheetName, 45, 3, null),
					new ExpectedCellValue(sheetName, 46, 3, null),
					new ExpectedCellValue(sheetName, 47, 3, null),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 50, 3, null),
					new ExpectedCellValue(sheetName, 51, 3, 3d),
					new ExpectedCellValue(sheetName, 31, 4, "Posting Date"),
					new ExpectedCellValue(sheetName, 32, 4, 2018),
					new ExpectedCellValue(sheetName, 33, 4, "Jan"),
					new ExpectedCellValue(sheetName, 34, 4, 4d),
					new ExpectedCellValue(sheetName, 35, 4, 4d),
					new ExpectedCellValue(sheetName, 36, 4, 1d),
					new ExpectedCellValue(sheetName, 37, 4, 1d),
					new ExpectedCellValue(sheetName, 38, 4, 2d),
					new ExpectedCellValue(sheetName, 39, 4, 2d),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, 1d),
					new ExpectedCellValue(sheetName, 44, 4, 1d),
					new ExpectedCellValue(sheetName, 45, 4, 1d),
					new ExpectedCellValue(sheetName, 46, 4, 1d),
					new ExpectedCellValue(sheetName, 47, 4, 1d),
					new ExpectedCellValue(sheetName, 48, 4, 1d),
					new ExpectedCellValue(sheetName, 49, 4, 2d),
					new ExpectedCellValue(sheetName, 50, 4, 2d),
					new ExpectedCellValue(sheetName, 51, 4, 12d),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 32, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 34, 5, 4d),
					new ExpectedCellValue(sheetName, 35, 5, 4d),
					new ExpectedCellValue(sheetName, 36, 5, 1d),
					new ExpectedCellValue(sheetName, 37, 5, 1d),
					new ExpectedCellValue(sheetName, 38, 5, 2d),
					new ExpectedCellValue(sheetName, 39, 5, 2d),
					new ExpectedCellValue(sheetName, 40, 5, 3d),
					new ExpectedCellValue(sheetName, 41, 5, 2d),
					new ExpectedCellValue(sheetName, 42, 5, 1d),
					new ExpectedCellValue(sheetName, 43, 5, 1d),
					new ExpectedCellValue(sheetName, 44, 5, 1d),
					new ExpectedCellValue(sheetName, 45, 5, 1d),
					new ExpectedCellValue(sheetName, 46, 5, 1d),
					new ExpectedCellValue(sheetName, 47, 5, 1d),
					new ExpectedCellValue(sheetName, 48, 5, 1d),
					new ExpectedCellValue(sheetName, 49, 5, 2d),
					new ExpectedCellValue(sheetName, 50, 5, 2d),
					new ExpectedCellValue(sheetName, 51, 5, 15d)
				});
			}
		}
	}
}
