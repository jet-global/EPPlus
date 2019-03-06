/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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

namespace EPPlusTest.Table.PivotTable.Filters
{
	[TestClass]
	public class PivotTableFiltersTest
	{
		#region Label Filters Tests
		#region CaptionEquals Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 1194d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 2857d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsColumnFilterOnlyRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B10:D14"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 10, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 11, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 12, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 14, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 10, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 11, 3, "February"),
					new ExpectedCellValue(sheetName, 12, 3, 1194d),
					new ExpectedCellValue(sheetName, 13, 3, 99d),
					new ExpectedCellValue(sheetName, 14, 3, 1293d),
					new ExpectedCellValue(sheetName, 10, 4, null),
					new ExpectedCellValue(sheetName, 11, 4, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 4, 1194d),
					new ExpectedCellValue(sheetName, 13, 4, 99d),
					new ExpectedCellValue(sheetName, 14, 4, 1293d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsRowAndColumnFilterRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B18:D21"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 18, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 19, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 20, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 21, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 18, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 19, 3, "February"),
					new ExpectedCellValue(sheetName, 20, 3, 99d),
					new ExpectedCellValue(sheetName, 21, 3, 99d),
					new ExpectedCellValue(sheetName, 18, 4, null),
					new ExpectedCellValue(sheetName, 19, 4, "Grand Total"),
					new ExpectedCellValue(sheetName, 20, 4, 99d),
					new ExpectedCellValue(sheetName, 21, 4, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsRowFilterEnabledForAllRowFieldsTwoRowFieldsOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I3:K7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 9, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 9, "March"),
					new ExpectedCellValue(sheetName, 7, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 10, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 5, 10, 24.99),
					new ExpectedCellValue(sheetName, 6, 10, 24.99),
					new ExpectedCellValue(sheetName, 7, 10, 24.99),
					new ExpectedCellValue(sheetName, 3, 11, null),
					new ExpectedCellValue(sheetName, 4, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 6, 11, 24.99),
					new ExpectedCellValue(sheetName, 7, 11, 24.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsColumnFilterEnabledForAllColumnFieldsTwoColumnFieldsOneRowField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I11:L15"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 11, 9, "Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 9, null),
					new ExpectedCellValue(sheetName, 13, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 14, 9, "Nashville"),
					new ExpectedCellValue(sheetName, 15, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 11, 10, "Column Labels"),
					new ExpectedCellValue(sheetName, 12, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 13, 10, "March"),
					new ExpectedCellValue(sheetName, 14, 10, 831.5),
					new ExpectedCellValue(sheetName, 15, 10, 831.5),
					new ExpectedCellValue(sheetName, 11, 11, null),
					new ExpectedCellValue(sheetName, 12, 11, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 13, 11, null),
					new ExpectedCellValue(sheetName, 14, 11, 831.5),
					new ExpectedCellValue(sheetName, 15, 11, 831.5),
					new ExpectedCellValue(sheetName, 11, 12, null),
					new ExpectedCellValue(sheetName, 12, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 12, null),
					new ExpectedCellValue(sheetName, 14, 12, 831.5),
					new ExpectedCellValue(sheetName, 15, 12, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsRowAndColumnFiltersEnabledForAllFieldsTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I19:L24"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 9, "Sum of Total"),
					new ExpectedCellValue(sheetName, 20, 9, null),
					new ExpectedCellValue(sheetName, 21, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 22, 9, "Nashville"),
					new ExpectedCellValue(sheetName, 23, 9, 20100090),
					new ExpectedCellValue(sheetName, 24, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 19, 10, "Column Labels"),
					new ExpectedCellValue(sheetName, 20, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 21, 10, "January"),
					new ExpectedCellValue(sheetName, 22, 10, 831.5),
					new ExpectedCellValue(sheetName, 23, 10, 831.5),
					new ExpectedCellValue(sheetName, 24, 10, 831.5),
					new ExpectedCellValue(sheetName, 19, 11, null),
					new ExpectedCellValue(sheetName, 20, 11, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 21, 11, null),
					new ExpectedCellValue(sheetName, 22, 11, 831.5),
					new ExpectedCellValue(sheetName, 23, 11, 831.5),
					new ExpectedCellValue(sheetName, 24, 11, 831.5),
					new ExpectedCellValue(sheetName, 19, 12, null),
					new ExpectedCellValue(sheetName, 20, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 21, 12, null),
					new ExpectedCellValue(sheetName, 22, 12, 831.5),
					new ExpectedCellValue(sheetName, 23, 12, 831.5),
					new ExpectedCellValue(sheetName, 24, 12, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsMultipleRowDataFieldsRowFiltersEnabledOneRowFieldOneColumnFieldFirstDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B27:E34"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 27, 2, null),
					new ExpectedCellValue(sheetName, 28, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 29, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 31, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 33, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 28, 3, "January"),
					new ExpectedCellValue(sheetName, 29, 3, null),
					new ExpectedCellValue(sheetName, 30, 3, 415.75),
					new ExpectedCellValue(sheetName, 31, 3, null),
					new ExpectedCellValue(sheetName, 32, 3, 831.5),
					new ExpectedCellValue(sheetName, 33, 3, 415.75),
					new ExpectedCellValue(sheetName, 34, 3, 831.5),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, "March"),
					new ExpectedCellValue(sheetName, 29, 4, null),
					new ExpectedCellValue(sheetName, 30, 4, 24.99),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 32, 4, 24.99),
					new ExpectedCellValue(sheetName, 33, 4, 24.99),
					new ExpectedCellValue(sheetName, 34, 4, 24.99),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, 440.74),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 32, 5, 856.49),
					new ExpectedCellValue(sheetName, 33, 5, 440.74),
					new ExpectedCellValue(sheetName, 34, 5, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsMultipleRowDataFieldsRowFiltersEnabledOneRowFieldOneColumnFieldLeafDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B40:E46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 2, null),
					new ExpectedCellValue(sheetName, 41, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 42, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 43, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 44, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 45, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 46, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 40, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 41, 3, "January"),
					new ExpectedCellValue(sheetName, 42, 3, null),
					new ExpectedCellValue(sheetName, 43, 3, 415.75),
					new ExpectedCellValue(sheetName, 44, 3, 415.75),
					new ExpectedCellValue(sheetName, 45, 3, 415.75),
					new ExpectedCellValue(sheetName, 46, 3, 415.75),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 41, 4, "February"),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, 99d),
					new ExpectedCellValue(sheetName, 44, 4, 99d),
					new ExpectedCellValue(sheetName, 45, 4, 99d),
					new ExpectedCellValue(sheetName, 46, 4, 99d),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 41, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, 514.75),
					new ExpectedCellValue(sheetName, 44, 5, 514.75),
					new ExpectedCellValue(sheetName, 45, 5, 514.75),
					new ExpectedCellValue(sheetName, 46, 5, 514.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsMultipleColumnDataFieldsColumnFiltersEnabledTwoRowFieldsOneColumnFieldFirstDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable9"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I28:M35"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 28, 9, null),
					new ExpectedCellValue(sheetName, 29, 9, null),
					new ExpectedCellValue(sheetName, 30, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 31, 9, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 9, "Tent"),
					new ExpectedCellValue(sheetName, 33, 9, "San Francisco"),
					new ExpectedCellValue(sheetName, 34, 9, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 35, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 28, 10, "Column Labels"),
					new ExpectedCellValue(sheetName, 29, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 10, "February"),
					new ExpectedCellValue(sheetName, 31, 10, 199d),
					new ExpectedCellValue(sheetName, 32, 10, 199d),
					new ExpectedCellValue(sheetName, 33, 10, 99d),
					new ExpectedCellValue(sheetName, 34, 10, 99d),
					new ExpectedCellValue(sheetName, 35, 10, 298d),
					new ExpectedCellValue(sheetName, 28, 11, null),
					new ExpectedCellValue(sheetName, 29, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 11, "February"),
					new ExpectedCellValue(sheetName, 31, 11, 1194d),
					new ExpectedCellValue(sheetName, 32, 11, 1194d),
					new ExpectedCellValue(sheetName, 33, 11, 99d),
					new ExpectedCellValue(sheetName, 34, 11, 99d),
					new ExpectedCellValue(sheetName, 35, 11, 1293d),
					new ExpectedCellValue(sheetName, 28, 12, null),
					new ExpectedCellValue(sheetName, 29, 12, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 12, null),
					new ExpectedCellValue(sheetName, 31, 12, 199d),
					new ExpectedCellValue(sheetName, 32, 12, 199d),
					new ExpectedCellValue(sheetName, 33, 12, 99d),
					new ExpectedCellValue(sheetName, 34, 12, 99d),
					new ExpectedCellValue(sheetName, 35, 12, 298d),
					new ExpectedCellValue(sheetName, 28, 13, null),
					new ExpectedCellValue(sheetName, 29, 13, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 13, null),
					new ExpectedCellValue(sheetName, 31, 13, 1194d),
					new ExpectedCellValue(sheetName, 32, 13, 1194d),
					new ExpectedCellValue(sheetName, 33, 13, 99d),
					new ExpectedCellValue(sheetName, 34, 13, 99d),
					new ExpectedCellValue(sheetName, 35, 13, 1293d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsMultipleColumnDataFieldsColumnFiltersEnabledOneRowFieldTwoColumnFieldsInnerDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable10"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I40:O45"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 9, null),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 42, 9, null),
					new ExpectedCellValue(sheetName, 43, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 9, "Nashville"),
					new ExpectedCellValue(sheetName, 45, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 10, "Column Labels"),
					new ExpectedCellValue(sheetName, 41, 10, "Tent"),
					new ExpectedCellValue(sheetName, 42, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 43, 10, "February"),
					new ExpectedCellValue(sheetName, 44, 10, 199d),
					new ExpectedCellValue(sheetName, 45, 10, 199d),
					new ExpectedCellValue(sheetName, 40, 11, null),
					new ExpectedCellValue(sheetName, 41, 11, null),
					new ExpectedCellValue(sheetName, 42, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 11, "February"),
					new ExpectedCellValue(sheetName, 44, 11, 1194d),
					new ExpectedCellValue(sheetName, 45, 11, 1194d),
					new ExpectedCellValue(sheetName, 40, 12, null),
					new ExpectedCellValue(sheetName, 41, 12, "Tent Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 42, 12, null),
					new ExpectedCellValue(sheetName, 43, 12, null),
					new ExpectedCellValue(sheetName, 44, 12, 199d),
					new ExpectedCellValue(sheetName, 45, 12, 199d),
					new ExpectedCellValue(sheetName, 40, 13, null),
					new ExpectedCellValue(sheetName, 41, 13, "Tent Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 13, null),
					new ExpectedCellValue(sheetName, 43, 13, null),
					new ExpectedCellValue(sheetName, 44, 13, 1194d),
					new ExpectedCellValue(sheetName, 45, 13, 1194d),
					new ExpectedCellValue(sheetName, 40, 14, null),
					new ExpectedCellValue(sheetName, 41, 14, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 42, 14, null),
					new ExpectedCellValue(sheetName, 43, 14, null),
					new ExpectedCellValue(sheetName, 44, 14, 199d),
					new ExpectedCellValue(sheetName, 45, 14, 199d),
					new ExpectedCellValue(sheetName, 40, 15, null),
					new ExpectedCellValue(sheetName, 41, 15, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 15, null),
					new ExpectedCellValue(sheetName, 43, 15, null),
					new ExpectedCellValue(sheetName, 44, 15, 1194d),
					new ExpectedCellValue(sheetName, 45, 15, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsWithRegularExpressionRowFiltersEnabledTwoRowFieldsOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable11"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B51:E58"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 51, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 52, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 53, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 54, 2, "January"),
					new ExpectedCellValue(sheetName, 55, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 56, 2, "January"),
					new ExpectedCellValue(sheetName, 57, 2, "February"),
					new ExpectedCellValue(sheetName, 58, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 51, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 52, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 53, 3, 831.5),
					new ExpectedCellValue(sheetName, 54, 3, 831.5),
					new ExpectedCellValue(sheetName, 55, 3, 415.75),
					new ExpectedCellValue(sheetName, 56, 3, 415.75),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 58, 3, 1247.25),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 52, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 53, 4, null),
					new ExpectedCellValue(sheetName, 54, 4, null),
					new ExpectedCellValue(sheetName, 55, 4, 99d),
					new ExpectedCellValue(sheetName, 56, 4, null),
					new ExpectedCellValue(sheetName, 57, 4, 99d),
					new ExpectedCellValue(sheetName, 58, 4, 99d),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 53, 5, 831.5),
					new ExpectedCellValue(sheetName, 54, 5, 831.5),
					new ExpectedCellValue(sheetName, 55, 5, 514.75),
					new ExpectedCellValue(sheetName, 56, 5, 415.75),
					new ExpectedCellValue(sheetName, 57, 5, 99d),
					new ExpectedCellValue(sheetName, 58, 5, 1346.25)
				});
			}
		}
	
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsWithRegularExpressionColumnFiltersEnabledOneRowFieldTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable12"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I51:L57"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 51, 9, "Sum of Total"),
					new ExpectedCellValue(sheetName, 52, 9, null),
					new ExpectedCellValue(sheetName, 53, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 54, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 55, 9, "Nashville"),
					new ExpectedCellValue(sheetName, 56, 9, "San Francisco"),
					new ExpectedCellValue(sheetName, 57, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 51, 10, "Column Labels"),
					new ExpectedCellValue(sheetName, 52, 10, "January"),
					new ExpectedCellValue(sheetName, 53, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 54, 10, 831.5),
					new ExpectedCellValue(sheetName, 55, 10, 831.5),
					new ExpectedCellValue(sheetName, 56, 10, 415.75),
					new ExpectedCellValue(sheetName, 57, 10, 2078.75),
					new ExpectedCellValue(sheetName, 51, 11, null),
					new ExpectedCellValue(sheetName, 52, 11, "January Total"),
					new ExpectedCellValue(sheetName, 53, 11, null),
					new ExpectedCellValue(sheetName, 54, 11, 831.5),
					new ExpectedCellValue(sheetName, 55, 11, 831.5),
					new ExpectedCellValue(sheetName, 56, 11, 415.75),
					new ExpectedCellValue(sheetName, 57, 11, 2078.75),
					new ExpectedCellValue(sheetName, 51, 12, null),
					new ExpectedCellValue(sheetName, 52, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 53, 12, null),
					new ExpectedCellValue(sheetName, 54, 12, 831.5),
					new ExpectedCellValue(sheetName, 55, 12, 831.5),
					new ExpectedCellValue(sheetName, 56, 12, 415.75),
					new ExpectedCellValue(sheetName, 57, 12, 2078.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable13"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B63:G76"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 63, 2, null),
					new ExpectedCellValue(sheetName, 64, 2, null),
					new ExpectedCellValue(sheetName, 65, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 66, 2, "February"),
					new ExpectedCellValue(sheetName, 67, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 68, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 69, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 70, 2, "Tent"),
					new ExpectedCellValue(sheetName, 71, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 72, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 73, 2, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 74, 2, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 75, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 76, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 63, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 64, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 65, 3, 20100070),
					new ExpectedCellValue(sheetName, 66, 3, null),
					new ExpectedCellValue(sheetName, 67, 3, null),
					new ExpectedCellValue(sheetName, 68, 3, null),
					new ExpectedCellValue(sheetName, 69, 3, null),
					new ExpectedCellValue(sheetName, 70, 3, null),
					new ExpectedCellValue(sheetName, 71, 3, 199d),
					new ExpectedCellValue(sheetName, 72, 3, 1194d),
					new ExpectedCellValue(sheetName, 73, 3, 199d),
					new ExpectedCellValue(sheetName, 74, 3, 1194d),
					new ExpectedCellValue(sheetName, 75, 3, 199d),
					new ExpectedCellValue(sheetName, 76, 3, 1194d),
					new ExpectedCellValue(sheetName, 63, 4, null),
					new ExpectedCellValue(sheetName, 64, 4, "Nashville Total"),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 66, 4, null),
					new ExpectedCellValue(sheetName, 67, 4, null),
					new ExpectedCellValue(sheetName, 68, 4, null),
					new ExpectedCellValue(sheetName, 69, 4, null),
					new ExpectedCellValue(sheetName, 70, 4, null),
					new ExpectedCellValue(sheetName, 71, 4, 199d),
					new ExpectedCellValue(sheetName, 72, 4, 1194d),
					new ExpectedCellValue(sheetName, 73, 4, 199d),
					new ExpectedCellValue(sheetName, 74, 4, 1194d),
					new ExpectedCellValue(sheetName, 75, 4, 199d),
					new ExpectedCellValue(sheetName, 76, 4, 1194d),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 64, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 65, 5, 20100085),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 67, 5, null),
					new ExpectedCellValue(sheetName, 68, 5, 99d),
					new ExpectedCellValue(sheetName, 69, 5, 99d),
					new ExpectedCellValue(sheetName, 70, 5, null),
					new ExpectedCellValue(sheetName, 71, 5, null),
					new ExpectedCellValue(sheetName, 72, 5, null),
					new ExpectedCellValue(sheetName, 73, 5, 99d),
					new ExpectedCellValue(sheetName, 74, 5, 99d),
					new ExpectedCellValue(sheetName, 75, 5, 99d),
					new ExpectedCellValue(sheetName, 76, 5, 99d),
					new ExpectedCellValue(sheetName, 63, 6, null),
					new ExpectedCellValue(sheetName, 64, 6, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 65, 6, null),
					new ExpectedCellValue(sheetName, 66, 6, null),
					new ExpectedCellValue(sheetName, 67, 6, null),
					new ExpectedCellValue(sheetName, 68, 6, 99d),
					new ExpectedCellValue(sheetName, 69, 6, 99d),
					new ExpectedCellValue(sheetName, 70, 6, null),
					new ExpectedCellValue(sheetName, 71, 6, null),
					new ExpectedCellValue(sheetName, 72, 6, null),
					new ExpectedCellValue(sheetName, 73, 6, 99d),
					new ExpectedCellValue(sheetName, 74, 6, 99d),
					new ExpectedCellValue(sheetName, 75, 6, 99d),
					new ExpectedCellValue(sheetName, 76, 6, 99d),
					new ExpectedCellValue(sheetName, 63, 7, null),
					new ExpectedCellValue(sheetName, 64, 7, "Grand Total"),
					new ExpectedCellValue(sheetName, 65, 7, null),
					new ExpectedCellValue(sheetName, 66, 7, null),
					new ExpectedCellValue(sheetName, 67, 7, null),
					new ExpectedCellValue(sheetName, 68, 7, 99d),
					new ExpectedCellValue(sheetName, 69, 7, 99d),
					new ExpectedCellValue(sheetName, 70, 7, null),
					new ExpectedCellValue(sheetName, 71, 7, 199d),
					new ExpectedCellValue(sheetName, 72, 7, 1194d),
					new ExpectedCellValue(sheetName, 73, 7, 298d),
					new ExpectedCellValue(sheetName, 74, 7, 1293d),
					new ExpectedCellValue(sheetName, 75, 7, 298d),
					new ExpectedCellValue(sheetName, 76, 7, 1293d)

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEqualsWithRegularExpressionRowAndColumnFiltersEnabledNoMatchTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable14"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B81:D84"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 81, 2, null),
					new ExpectedCellValue(sheetName, 82, 2, null),
					new ExpectedCellValue(sheetName, 83, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 84, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 81, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 82, 3, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 83, 3, null),
					new ExpectedCellValue(sheetName, 84, 3, null),
					new ExpectedCellValue(sheetName, 81, 4, null),
					new ExpectedCellValue(sheetName, 82, 4, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 83, 4, null),
					new ExpectedCellValue(sheetName, 84, 4, null),
				});
			}
		}
		#endregion

		#region CaptionNotEquals Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 7, 3, 1663d),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, null),
					new ExpectedCellValue(sheetName, 6, 4, 1194d),
					new ExpectedCellValue(sheetName, 7, 4, 1194d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 24.99),
					new ExpectedCellValue(sheetName, 6, 5, 831.5),
					new ExpectedCellValue(sheetName, 7, 5, 856.49),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 856.49),
					new ExpectedCellValue(sheetName, 6, 6, 2857d),
					new ExpectedCellValue(sheetName, 7, 6, 3713.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsColumnFilterOnlyRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:M8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 8, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "January"),
					new ExpectedCellValue(sheetName, 5, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 2078.75),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "March"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, null),
					new ExpectedCellValue(sheetName, 8, 12, 856.49),
					new ExpectedCellValue(sheetName, 3, 13, null),
					new ExpectedCellValue(sheetName, 4, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 13, 856.49),
					new ExpectedCellValue(sheetName, 6, 13, 1663d),
					new ExpectedCellValue(sheetName, 7, 13, 415.75),
					new ExpectedCellValue(sheetName, 8, 13, 2935.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsRowAndColumnFilterRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:E16"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 12, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 14, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 16, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 13, 3, "January"),
					new ExpectedCellValue(sheetName, 14, 3, 831.5),
					new ExpectedCellValue(sheetName, 15, 3, 831.5),
					new ExpectedCellValue(sheetName, 16, 3, 1663d),
					new ExpectedCellValue(sheetName, 12, 4, null),
					new ExpectedCellValue(sheetName, 13, 4, "February"),
					new ExpectedCellValue(sheetName, 14, 4, null),
					new ExpectedCellValue(sheetName, 15, 4, 1194d),
					new ExpectedCellValue(sheetName, 16, 4, 1194d),
					new ExpectedCellValue(sheetName, 12, 5, null),
					new ExpectedCellValue(sheetName, 13, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 14, 5, 831.5),
					new ExpectedCellValue(sheetName, 15, 5, 2025.5),
					new ExpectedCellValue(sheetName, 16, 5, 2857d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsRowFilterEnabledForAllRowFieldsTwoRowFieldsOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J12:M18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 12, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 14, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 16, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 10, "Tent"),
					new ExpectedCellValue(sheetName, 18, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 13, 11, "February"),
					new ExpectedCellValue(sheetName, 14, 11, null),
					new ExpectedCellValue(sheetName, 15, 11, null),
					new ExpectedCellValue(sheetName, 16, 11, 1194d),
					new ExpectedCellValue(sheetName, 17, 11, 1194d),
					new ExpectedCellValue(sheetName, 18, 11, 1194d),
					new ExpectedCellValue(sheetName, 12, 12, null),
					new ExpectedCellValue(sheetName, 13, 12, "March"),
					new ExpectedCellValue(sheetName, 14, 12, 24.99),
					new ExpectedCellValue(sheetName, 15, 12, 24.99),
					new ExpectedCellValue(sheetName, 16, 12, null),
					new ExpectedCellValue(sheetName, 17, 12, null),
					new ExpectedCellValue(sheetName, 18, 12, 24.99),
					new ExpectedCellValue(sheetName, 12, 13, null),
					new ExpectedCellValue(sheetName, 13, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 14, 13, 24.99),
					new ExpectedCellValue(sheetName, 15, 13, 24.99),
					new ExpectedCellValue(sheetName, 16, 13, 1194d),
					new ExpectedCellValue(sheetName, 17, 13, 1194d),
					new ExpectedCellValue(sheetName, 18, 13, 1218.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsColumnFilterEnabledForAllColumnFieldsTwoColumnFieldsOneRowField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B21:G26"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 22, 2, null),
					new ExpectedCellValue(sheetName, 23, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 24, 2, "February"),
					new ExpectedCellValue(sheetName, 25, 2, "March"),
					new ExpectedCellValue(sheetName, 26, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 21, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 22, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 23, 3, "Headlamp"),
					new ExpectedCellValue(sheetName, 24, 3, null),
					new ExpectedCellValue(sheetName, 25, 3, 24.99),
					new ExpectedCellValue(sheetName, 26, 3, 24.99),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 22, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 23, 4, null),
					new ExpectedCellValue(sheetName, 24, 4, null),
					new ExpectedCellValue(sheetName, 25, 4, 24.99),
					new ExpectedCellValue(sheetName, 26, 4, 24.99),
					new ExpectedCellValue(sheetName, 21, 5, null),
					new ExpectedCellValue(sheetName, 22, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 23, 5, "Tent"),
					new ExpectedCellValue(sheetName, 24, 5, 1194d),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, 1194d),
					new ExpectedCellValue(sheetName, 21, 6, null),
					new ExpectedCellValue(sheetName, 22, 6, "Nashville Total"),
					new ExpectedCellValue(sheetName, 23, 6, null),
					new ExpectedCellValue(sheetName, 24, 6, 1194d),
					new ExpectedCellValue(sheetName, 25, 6, null),
					new ExpectedCellValue(sheetName, 26, 6, 1194d),
					new ExpectedCellValue(sheetName, 21, 7, null),
					new ExpectedCellValue(sheetName, 22, 7, "Grand Total"),
					new ExpectedCellValue(sheetName, 23, 7, null),
					new ExpectedCellValue(sheetName, 24, 7, 1194d),
					new ExpectedCellValue(sheetName, 25, 7, 24.99),
					new ExpectedCellValue(sheetName, 26, 7, 1218.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsRowAndColumnFiltersEnabledForAllFieldsTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J21:M26"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 22, 10, null),
					new ExpectedCellValue(sheetName, 23, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 24, 10, "January"),
					new ExpectedCellValue(sheetName, 25, 10, 20100007),
					new ExpectedCellValue(sheetName, 26, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 21, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 22, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 23, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 24, 11, 831.5),
					new ExpectedCellValue(sheetName, 25, 11, 831.5),
					new ExpectedCellValue(sheetName, 26, 11, 831.5),
					new ExpectedCellValue(sheetName, 21, 12, null),
					new ExpectedCellValue(sheetName, 22, 12, "Chicago Total"),
					new ExpectedCellValue(sheetName, 23, 12, null),
					new ExpectedCellValue(sheetName, 24, 12, 831.5),
					new ExpectedCellValue(sheetName, 25, 12, 831.5),
					new ExpectedCellValue(sheetName, 26, 12, 831.5),
					new ExpectedCellValue(sheetName, 21, 13, null),
					new ExpectedCellValue(sheetName, 22, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 23, 13, null),
					new ExpectedCellValue(sheetName, 24, 13, 831.5),
					new ExpectedCellValue(sheetName, 25, 13, 831.5),
					new ExpectedCellValue(sheetName, 26, 13, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsMultipleRowDataFieldsRowFiltersEnabledOneRowFieldOneColumnFieldFirstDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B30:F39"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 30, 2, null),
					new ExpectedCellValue(sheetName, 31, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 34, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 35, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 36, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 37, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 38, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 39, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 31, 3, "January"),
					new ExpectedCellValue(sheetName, 32, 3, null),
					new ExpectedCellValue(sheetName, 33, 3, 415.75),
					new ExpectedCellValue(sheetName, 34, 3, 415.75),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 36, 3, 831.5),
					new ExpectedCellValue(sheetName, 37, 3, 415.75),
					new ExpectedCellValue(sheetName, 38, 3, 831.5),
					new ExpectedCellValue(sheetName, 39, 3, 1247.25),
					new ExpectedCellValue(sheetName, 30, 4, null),
					new ExpectedCellValue(sheetName, 31, 4, "February"),
					new ExpectedCellValue(sheetName, 32, 4, null),
					new ExpectedCellValue(sheetName, 33, 4, 199d),
					new ExpectedCellValue(sheetName, 34, 4, 99d),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 36, 4, 1194d),
					new ExpectedCellValue(sheetName, 37, 4, 99d),
					new ExpectedCellValue(sheetName, 38, 4, 298d),
					new ExpectedCellValue(sheetName, 39, 4, 1293d),
					new ExpectedCellValue(sheetName, 30, 5, null),
					new ExpectedCellValue(sheetName, 31, 5, "March"),
					new ExpectedCellValue(sheetName, 32, 5, null),
					new ExpectedCellValue(sheetName, 33, 5, 415.75),
					new ExpectedCellValue(sheetName, 34, 5, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 36, 5, 831.5),
					new ExpectedCellValue(sheetName, 37, 5, null),
					new ExpectedCellValue(sheetName, 38, 5, 415.75),
					new ExpectedCellValue(sheetName, 39, 5, 831.5),
					new ExpectedCellValue(sheetName, 30, 6, null),
					new ExpectedCellValue(sheetName, 31, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 32, 6, null),
					new ExpectedCellValue(sheetName, 33, 6, 1030.5),
					new ExpectedCellValue(sheetName, 34, 6, 514.75),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 36, 6, 2857d),
					new ExpectedCellValue(sheetName, 37, 6, 514.75),
					new ExpectedCellValue(sheetName, 38, 6, 1545.25),
					new ExpectedCellValue(sheetName, 39, 6, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsMultipleRowDataFieldsRowFiltersEnabledOneRowFieldOneColumnFieldLeafDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J30:N39"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 30, 10, null),
					new ExpectedCellValue(sheetName, 31, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 33, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 35, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 36, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 37, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 39, 10, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 31, 11, "January"),
					new ExpectedCellValue(sheetName, 32, 11, null),
					new ExpectedCellValue(sheetName, 33, 11, 415.75),
					new ExpectedCellValue(sheetName, 34, 11, 831.5),
					new ExpectedCellValue(sheetName, 35, 11, null),
					new ExpectedCellValue(sheetName, 36, 11, 415.75),
					new ExpectedCellValue(sheetName, 37, 11, 415.75),
					new ExpectedCellValue(sheetName, 38, 11, 831.5),
					new ExpectedCellValue(sheetName, 39, 11, 1247.25),
					new ExpectedCellValue(sheetName, 30, 12, null),
					new ExpectedCellValue(sheetName, 31, 12, "February"),
					new ExpectedCellValue(sheetName, 32, 12, null),
					new ExpectedCellValue(sheetName, 33, 12, 199d),
					new ExpectedCellValue(sheetName, 34, 12, 1194d),
					new ExpectedCellValue(sheetName, 35, 12, null),
					new ExpectedCellValue(sheetName, 36, 12, 99d),
					new ExpectedCellValue(sheetName, 37, 12, 99d),
					new ExpectedCellValue(sheetName, 38, 12, 298d),
					new ExpectedCellValue(sheetName, 39, 12, 1293d),
					new ExpectedCellValue(sheetName, 30, 13, null),
					new ExpectedCellValue(sheetName, 31, 13, "March"),
					new ExpectedCellValue(sheetName, 32, 13, null),
					new ExpectedCellValue(sheetName, 33, 13, 415.75),
					new ExpectedCellValue(sheetName, 34, 13, 831.5),
					new ExpectedCellValue(sheetName, 35, 13, null),
					new ExpectedCellValue(sheetName, 36, 13, null),
					new ExpectedCellValue(sheetName, 37, 13, null),
					new ExpectedCellValue(sheetName, 38, 13, 415.75),
					new ExpectedCellValue(sheetName, 39, 13, 831.5),
					new ExpectedCellValue(sheetName, 30, 14, null),
					new ExpectedCellValue(sheetName, 31, 14, "Grand Total"),
					new ExpectedCellValue(sheetName, 32, 14, null),
					new ExpectedCellValue(sheetName, 33, 14, 1030.5),
					new ExpectedCellValue(sheetName, 34, 14, 2857d),
					new ExpectedCellValue(sheetName, 35, 14, null),
					new ExpectedCellValue(sheetName, 36, 14, 514.75),
					new ExpectedCellValue(sheetName, 37, 14, 514.75),
					new ExpectedCellValue(sheetName, 38, 14, 1545.25),
					new ExpectedCellValue(sheetName, 39, 14, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsMultipleColumnDataFieldsColumnFiltersEnabledTwoRowFieldsOneColumnFieldFirstDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable9"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B44:H50"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 44, 2, null),
					new ExpectedCellValue(sheetName, 45, 2, null),
					new ExpectedCellValue(sheetName, 46, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 2, "January"),
					new ExpectedCellValue(sheetName, 48, 2, "February"),
					new ExpectedCellValue(sheetName, 49, 2, "March"),
					new ExpectedCellValue(sheetName, 50, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 44, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 45, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 46, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 3, 415.75),
					new ExpectedCellValue(sheetName, 48, 3, 199d),
					new ExpectedCellValue(sheetName, 49, 3, 415.75),
					new ExpectedCellValue(sheetName, 50, 3, 1030.5),
					new ExpectedCellValue(sheetName, 44, 4, null),
					new ExpectedCellValue(sheetName, 45, 4, null),
					new ExpectedCellValue(sheetName, 46, 4, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 4, 415.75),
					new ExpectedCellValue(sheetName, 48, 4, 99d),
					new ExpectedCellValue(sheetName, 49, 4, null),
					new ExpectedCellValue(sheetName, 50, 4, 514.75),
					new ExpectedCellValue(sheetName, 44, 5, null),
					new ExpectedCellValue(sheetName, 45, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 46, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, 831.5),
					new ExpectedCellValue(sheetName, 48, 5, 1194d),
					new ExpectedCellValue(sheetName, 49, 5, 831.5),
					new ExpectedCellValue(sheetName, 50, 5, 2857d),
					new ExpectedCellValue(sheetName, 44, 6, null),
					new ExpectedCellValue(sheetName, 45, 6, null),
					new ExpectedCellValue(sheetName, 46, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, 415.75),
					new ExpectedCellValue(sheetName, 48, 6, 99d),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 6, 514.75),
					new ExpectedCellValue(sheetName, 44, 7, null),
					new ExpectedCellValue(sheetName, 45, 7, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 46, 7, null),
					new ExpectedCellValue(sheetName, 47, 7, 831.5),
					new ExpectedCellValue(sheetName, 48, 7, 298d),
					new ExpectedCellValue(sheetName, 49, 7, 415.75),
					new ExpectedCellValue(sheetName, 50, 7, 1545.25),
					new ExpectedCellValue(sheetName, 44, 8, null),
					new ExpectedCellValue(sheetName, 45, 8, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 46, 8, null),
					new ExpectedCellValue(sheetName, 47, 8, 1247.25),
					new ExpectedCellValue(sheetName, 48, 8, 1293d),
					new ExpectedCellValue(sheetName, 49, 8, 831.5),
					new ExpectedCellValue(sheetName, 50, 8, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsMultipleColumnDataFieldsColumnFiltersEnabledTwoRowFieldsOneColumnFieldLeafDataField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable10"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J44:P50"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 44, 10, null),
					new ExpectedCellValue(sheetName, 45, 10, null),
					new ExpectedCellValue(sheetName, 46, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 10, "January"),
					new ExpectedCellValue(sheetName, 48, 10, "February"),
					new ExpectedCellValue(sheetName, 49, 10, "March"),
					new ExpectedCellValue(sheetName, 50, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 44, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 45, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 46, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 11, 415.75),
					new ExpectedCellValue(sheetName, 48, 11, 199d),
					new ExpectedCellValue(sheetName, 49, 11, 415.75),
					new ExpectedCellValue(sheetName, 50, 11, 1030.5),
					new ExpectedCellValue(sheetName, 45, 12, null),
					new ExpectedCellValue(sheetName, 45, 12, null),
					new ExpectedCellValue(sheetName, 46, 12, "Sum of Total"),
					new ExpectedCellValue(sheetName, 47, 12, 831.5),
					new ExpectedCellValue(sheetName, 48, 12, 1194d),
					new ExpectedCellValue(sheetName, 49, 12, 831.5),
					new ExpectedCellValue(sheetName, 50, 12, 2857d),
					new ExpectedCellValue(sheetName, 44, 13, null),
					new ExpectedCellValue(sheetName, 45, 13, "San Francisco"),
					new ExpectedCellValue(sheetName, 46, 13, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 13, 415.75),
					new ExpectedCellValue(sheetName, 48, 13, 99d),
					new ExpectedCellValue(sheetName, 49, 13, null),
					new ExpectedCellValue(sheetName, 50, 13, 514.75),
					new ExpectedCellValue(sheetName, 44, 14, null),
					new ExpectedCellValue(sheetName, 45, 14, null),
					new ExpectedCellValue(sheetName, 46, 14, "Sum of Total"),
					new ExpectedCellValue(sheetName, 47, 14, 415.75),
					new ExpectedCellValue(sheetName, 48, 14, 99d),
					new ExpectedCellValue(sheetName, 49, 14, null),
					new ExpectedCellValue(sheetName, 50, 14, 514.75),
					new ExpectedCellValue(sheetName, 44, 15, null),
					new ExpectedCellValue(sheetName, 45, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 46, 15, null),
					new ExpectedCellValue(sheetName, 47, 15, 831.5),
					new ExpectedCellValue(sheetName, 48, 15, 298d),
					new ExpectedCellValue(sheetName, 49, 15, 415.75),
					new ExpectedCellValue(sheetName, 50, 15, 1545.25),
					new ExpectedCellValue(sheetName, 44, 16, null),
					new ExpectedCellValue(sheetName, 45, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 46, 16, null),
					new ExpectedCellValue(sheetName, 47, 16, 1247.25),
					new ExpectedCellValue(sheetName, 48, 16, 1293d),
					new ExpectedCellValue(sheetName, 49, 16, 831.5),
					new ExpectedCellValue(sheetName, 50, 16, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsWithRegularExpressionRowFiltersEnabledTwoRowFieldsOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable11"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B54:D58"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 54, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 55, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 56, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 57, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 54, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 55, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 56, 3, 831.5),
					new ExpectedCellValue(sheetName, 57, 3, 831.5),
					new ExpectedCellValue(sheetName, 58, 3, 831.5),
					new ExpectedCellValue(sheetName, 54, 4, null),
					new ExpectedCellValue(sheetName, 55, 4, "Grand Total"),
					new ExpectedCellValue(sheetName, 56, 4, 831.5),
					new ExpectedCellValue(sheetName, 57, 4, 831.5),
					new ExpectedCellValue(sheetName, 58, 4, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsWithRegularExpressionColumnFiltersEnabledOneRowFieldTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable12"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J54:M58"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 54, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 55, 10, null),
					new ExpectedCellValue(sheetName, 56, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 57, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 58, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 54, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 55, 11, "Tent"),
					new ExpectedCellValue(sheetName, 56, 11, "February"),
					new ExpectedCellValue(sheetName, 57, 11, 1194d),
					new ExpectedCellValue(sheetName, 58, 11, 1194d),
					new ExpectedCellValue(sheetName, 54, 12, null),
					new ExpectedCellValue(sheetName, 55, 12, "Tent Total"),
					new ExpectedCellValue(sheetName, 56, 12, null),
					new ExpectedCellValue(sheetName, 57, 12, 1194d),
					new ExpectedCellValue(sheetName, 58, 12, 1194d),
					new ExpectedCellValue(sheetName, 54, 13, null),
					new ExpectedCellValue(sheetName, 55, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 56, 13, null),
					new ExpectedCellValue(sheetName, 57, 13, 1194d),
					new ExpectedCellValue(sheetName, 58, 13, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable13"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B62:E72"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 62, 2, null),
					new ExpectedCellValue(sheetName, 63, 2, null),
					new ExpectedCellValue(sheetName, 64, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 65, 2, "March"),
					new ExpectedCellValue(sheetName, 66, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 67, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 68, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 69, 2, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 70, 2, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 71, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 72, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 62, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 63, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 64, 3, 20100083),
					new ExpectedCellValue(sheetName, 65, 3, null),
					new ExpectedCellValue(sheetName, 66, 3, null),
					new ExpectedCellValue(sheetName, 67, 3, 24.99),
					new ExpectedCellValue(sheetName, 68, 3, 24.99),
					new ExpectedCellValue(sheetName, 69, 3, 24.99),
					new ExpectedCellValue(sheetName, 70, 3, 24.99),
					new ExpectedCellValue(sheetName, 71, 3, 24.99),
					new ExpectedCellValue(sheetName, 72, 3, 24.99),
					new ExpectedCellValue(sheetName, 62, 4, null),
					new ExpectedCellValue(sheetName, 63, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 64, 4, null),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 66, 4, null),
					new ExpectedCellValue(sheetName, 67, 4, 24.99),
					new ExpectedCellValue(sheetName, 68, 4, 24.99),
					new ExpectedCellValue(sheetName, 69, 4, 24.99),
					new ExpectedCellValue(sheetName, 70, 4, 24.99),
					new ExpectedCellValue(sheetName, 71, 4, 24.99),
					new ExpectedCellValue(sheetName, 72, 4, 24.99),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 63, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 67, 5, 24.99),
					new ExpectedCellValue(sheetName, 68, 5, 24.99),
					new ExpectedCellValue(sheetName, 69, 5, 24.99),
					new ExpectedCellValue(sheetName, 70, 5, 24.99),
					new ExpectedCellValue(sheetName, 71, 5, 24.99),
					new ExpectedCellValue(sheetName, 72, 5, 24.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEqualsWithRegularExpressionRowAndColumnFiltersEnabledNoMatchTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEquals";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable14"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J62:K64"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 62, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 63, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 64, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 62, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 63, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 64, 11, null),
				});
			}
		}
		#endregion

		#region CaptionBeginsWith Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:E6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "March"),
					new ExpectedCellValue(sheetName, 5, 4, 24.99),
					new ExpectedCellValue(sheetName, 6, 4, 24.99),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 856.49),
					new ExpectedCellValue(sheetName, 6, 5, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithColumnFilterOnlyRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "March"),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 856.49),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithRowAndColumnFilterRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:E15"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 12, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 14, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 15, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 13, 3, 20100076),
					new ExpectedCellValue(sheetName, 14, 3, 415.75),
					new ExpectedCellValue(sheetName, 15, 3, 415.75),
					new ExpectedCellValue(sheetName, 12, 4, null),
					new ExpectedCellValue(sheetName, 13, 4, 20100085),
					new ExpectedCellValue(sheetName, 14, 4, 99d),
					new ExpectedCellValue(sheetName, 15, 4, 99d),
					new ExpectedCellValue(sheetName, 12, 5, null),
					new ExpectedCellValue(sheetName, 13, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 14, 5, 514.75),
					new ExpectedCellValue(sheetName, 15, 5, 514.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithRowAndColumnFiltersEnabledForAllFieldsTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B20:F26"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 20, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 2, null),
					new ExpectedCellValue(sheetName, 22, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 23, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 24, 2, 20100017),
					new ExpectedCellValue(sheetName, 25, 2, 20100090),
					new ExpectedCellValue(sheetName, 26, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 20, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 21, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 22, 3, "January"),
					new ExpectedCellValue(sheetName, 23, 3, 831.5),
					new ExpectedCellValue(sheetName, 24, 3, null),
					new ExpectedCellValue(sheetName, 25, 3, 831.5),
					new ExpectedCellValue(sheetName, 26, 3, 831.5),
					new ExpectedCellValue(sheetName, 20, 4, null),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 22, 4, "March"),
					new ExpectedCellValue(sheetName, 23, 4, 831.5),
					new ExpectedCellValue(sheetName, 24, 4, 831.5),
					new ExpectedCellValue(sheetName, 25, 4, null),
					new ExpectedCellValue(sheetName, 26, 4, 831.5),
					new ExpectedCellValue(sheetName, 20, 5, null),
					new ExpectedCellValue(sheetName, 21, 5, "Nashville Total"),
					new ExpectedCellValue(sheetName, 22, 5, null),
					new ExpectedCellValue(sheetName, 23, 5, 1663d),
					new ExpectedCellValue(sheetName, 24, 5, 831.5),
					new ExpectedCellValue(sheetName, 25, 5, 831.5),
					new ExpectedCellValue(sheetName, 26, 5, 1663d),
					new ExpectedCellValue(sheetName, 20, 6, null),
					new ExpectedCellValue(sheetName, 21, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 22, 6, null),
					new ExpectedCellValue(sheetName, 23, 6, 1663d),
					new ExpectedCellValue(sheetName, 24, 6, 831.5),
					new ExpectedCellValue(sheetName, 25, 6, 831.5),
					new ExpectedCellValue(sheetName, 26, 6, 1663d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithMultipleRowDataFieldsRowFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B31:E42"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 32, 2, null),
					new ExpectedCellValue(sheetName, 33, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 35, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 2, 20100017),
					new ExpectedCellValue(sheetName, 37, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 2, 20100017),
					new ExpectedCellValue(sheetName, 39, 2, "Car Rack Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 40, 2, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 42, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 33, 3, "March"),
					new ExpectedCellValue(sheetName, 34, 3, null),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 36, 3, 415.75),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 38, 3, 831.5),
					new ExpectedCellValue(sheetName, 39, 3, 415.75),
					new ExpectedCellValue(sheetName, 40, 3, 831.5),
					new ExpectedCellValue(sheetName, 41, 3, 415.75),
					new ExpectedCellValue(sheetName, 42, 3, 831.5),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville Total"),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 34, 4, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 36, 4, 415.75),
					new ExpectedCellValue(sheetName, 37, 4, null),
					new ExpectedCellValue(sheetName, 38, 4, 831.5),
					new ExpectedCellValue(sheetName, 39, 4, 415.75),
					new ExpectedCellValue(sheetName, 40, 4, 831.5),
					new ExpectedCellValue(sheetName, 41, 4, 415.75),
					new ExpectedCellValue(sheetName, 42, 4, 831.5),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 32, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 34, 5, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 36, 5, 415.75),
					new ExpectedCellValue(sheetName, 37, 5, null),
					new ExpectedCellValue(sheetName, 38, 5, 831.5),
					new ExpectedCellValue(sheetName, 39, 5, 415.75),
					new ExpectedCellValue(sheetName, 40, 5, 831.5),
					new ExpectedCellValue(sheetName, 41, 5, 415.75),
					new ExpectedCellValue(sheetName, 42, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithMultipleColumnDataFieldsColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J31:P37"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 10, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 35, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 10, 20100085),
					new ExpectedCellValue(sheetName, 37, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 31, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 32, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 33, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 11, "February"),
					new ExpectedCellValue(sheetName, 35, 11, 99d),
					new ExpectedCellValue(sheetName, 36, 11, 99d),
					new ExpectedCellValue(sheetName, 37, 11, 99d),
					new ExpectedCellValue(sheetName, 31, 12, null),
					new ExpectedCellValue(sheetName, 32, 12, null),
					new ExpectedCellValue(sheetName, 33, 12, "Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 12, "February"),
					new ExpectedCellValue(sheetName, 35, 12, 99d),
					new ExpectedCellValue(sheetName, 36, 12, 99d),
					new ExpectedCellValue(sheetName, 37, 12, 99d),
					new ExpectedCellValue(sheetName, 31, 13, null),
					new ExpectedCellValue(sheetName, 32, 13, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 13, null),
					new ExpectedCellValue(sheetName, 34, 13, null),
					new ExpectedCellValue(sheetName, 35, 13, 99d),
					new ExpectedCellValue(sheetName, 36, 13, 99d),
					new ExpectedCellValue(sheetName, 37, 13, 99d),
					new ExpectedCellValue(sheetName, 31, 14, null),
					new ExpectedCellValue(sheetName, 32, 14, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 14, null),
					new ExpectedCellValue(sheetName, 34, 14, null),
					new ExpectedCellValue(sheetName, 35, 14, 99d),
					new ExpectedCellValue(sheetName, 36, 14, 99d),
					new ExpectedCellValue(sheetName, 37, 14, 99d),
					new ExpectedCellValue(sheetName, 31, 15, null),
					new ExpectedCellValue(sheetName, 32, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 15, null),
					new ExpectedCellValue(sheetName, 34, 15, null),
					new ExpectedCellValue(sheetName, 35, 15, 99d),
					new ExpectedCellValue(sheetName, 36, 15, 99d),
					new ExpectedCellValue(sheetName, 37, 15, 99d),
					new ExpectedCellValue(sheetName, 31, 16, null),
					new ExpectedCellValue(sheetName, 32, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 16, null),
					new ExpectedCellValue(sheetName, 34, 16, null),
					new ExpectedCellValue(sheetName, 35, 16, 99d),
					new ExpectedCellValue(sheetName, 36, 16, 99d),
					new ExpectedCellValue(sheetName, 37, 16, 99d)
				});
			}
		}
		
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithWithRegularExpressionRowFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B47:G54"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 48, 2, null),
					new ExpectedCellValue(sheetName, 49, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 50, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 51, 2, 20100070),
					new ExpectedCellValue(sheetName, 52, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 53, 2, 20100076),
					new ExpectedCellValue(sheetName, 54, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 47, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 48, 3, "January"),
					new ExpectedCellValue(sheetName, 49, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, null),
					new ExpectedCellValue(sheetName, 51, 3, null),
					new ExpectedCellValue(sheetName, 52, 3, 415.75),
					new ExpectedCellValue(sheetName, 53, 3, 415.75),
					new ExpectedCellValue(sheetName, 54, 3, 415.75),
					new ExpectedCellValue(sheetName, 47, 4, null),
					new ExpectedCellValue(sheetName, 48, 4, "January Total"),
					new ExpectedCellValue(sheetName, 49, 4, null),
					new ExpectedCellValue(sheetName, 50, 4, null),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 52, 4, 415.75),
					new ExpectedCellValue(sheetName, 53, 4, 415.75),
					new ExpectedCellValue(sheetName, 54, 4, 415.75),
					new ExpectedCellValue(sheetName, 47, 5, null),
					new ExpectedCellValue(sheetName, 48, 5, "February"),
					new ExpectedCellValue(sheetName, 49, 5, "Tent"),
					new ExpectedCellValue(sheetName, 50, 5, 1194d),
					new ExpectedCellValue(sheetName, 51, 5, 1194d),
					new ExpectedCellValue(sheetName, 52, 5, null),
					new ExpectedCellValue(sheetName, 53, 5, null),
					new ExpectedCellValue(sheetName, 54, 5, 1194d),
					new ExpectedCellValue(sheetName, 47, 6, null),
					new ExpectedCellValue(sheetName, 48, 6, "February Total"),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 6, 1194d),
					new ExpectedCellValue(sheetName, 51, 6, 1194d),
					new ExpectedCellValue(sheetName, 52, 6, null),
					new ExpectedCellValue(sheetName, 53, 6, null),
					new ExpectedCellValue(sheetName, 54, 6, 1194d),
					new ExpectedCellValue(sheetName, 47, 7, null),
					new ExpectedCellValue(sheetName, 48, 7, "Grand Total"),
					new ExpectedCellValue(sheetName, 49, 7, null),
					new ExpectedCellValue(sheetName, 50, 7, 1194d),
					new ExpectedCellValue(sheetName, 51, 7, 1194d),
					new ExpectedCellValue(sheetName, 52, 7, 415.75),
					new ExpectedCellValue(sheetName, 53, 7, 415.75),
					new ExpectedCellValue(sheetName, 54, 7, 1609.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithWithRegularExpressionColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J47:M52"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 48, 10, null),
					new ExpectedCellValue(sheetName, 49, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 50, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 51, 10, 20100083),
					new ExpectedCellValue(sheetName, 52, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 47, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 48, 11, "March"),
					new ExpectedCellValue(sheetName, 49, 11, "Headlamp"),
					new ExpectedCellValue(sheetName, 50, 11, 24.99),
					new ExpectedCellValue(sheetName, 51, 11, 24.99),
					new ExpectedCellValue(sheetName, 52, 11, 24.99),
					new ExpectedCellValue(sheetName, 47, 12, null),
					new ExpectedCellValue(sheetName, 48, 12, "March Total"),
					new ExpectedCellValue(sheetName, 49, 12, null),
					new ExpectedCellValue(sheetName, 50, 12, 24.99),
					new ExpectedCellValue(sheetName, 51, 12, 24.99),
					new ExpectedCellValue(sheetName, 52, 12, 24.99),
					new ExpectedCellValue(sheetName, 47, 13, null),
					new ExpectedCellValue(sheetName, 48, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 49, 13, null),
					new ExpectedCellValue(sheetName, 50, 13, 24.99),
					new ExpectedCellValue(sheetName, 51, 13, 24.99),
					new ExpectedCellValue(sheetName, 52, 13, 24.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable9"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B59:E64"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 59, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 60, 2, null),
					new ExpectedCellValue(sheetName, 61, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 62, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 59, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 60, 3, "March"),
					new ExpectedCellValue(sheetName, 61, 3, 20100017),
					new ExpectedCellValue(sheetName, 62, 3, 831.5),
					new ExpectedCellValue(sheetName, 63, 3, 831.5),
					new ExpectedCellValue(sheetName, 64, 3, 831.5),
					new ExpectedCellValue(sheetName, 59, 4, null),
					new ExpectedCellValue(sheetName, 60, 4, "March Total"),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 62, 4, 831.5),
					new ExpectedCellValue(sheetName, 63, 4, 831.5),
					new ExpectedCellValue(sheetName, 64, 4, 831.5),
					new ExpectedCellValue(sheetName, 59, 5, null),
					new ExpectedCellValue(sheetName, 60, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 61, 5, null),
					new ExpectedCellValue(sheetName, 62, 5, 831.5),
					new ExpectedCellValue(sheetName, 63, 5, 831.5),
					new ExpectedCellValue(sheetName, 64, 5, 831.5)

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBeginsWithWithRegularExpressionRowAndColumnFiltersEnabledNoMatchTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable10"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J59:K62"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 59, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 60, 10, null),
					new ExpectedCellValue(sheetName, 61, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 62, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 59, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 60, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 61, 11, null),
					new ExpectedCellValue(sheetName, 62, 11, null),
				});
			}
		}
		#endregion

		#region CaptionNotBeginsWith Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 99d),
					new ExpectedCellValue(sheetName, 7, 4, 1293d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 514.75),
					new ExpectedCellValue(sheetName, 7, 6, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithColumnFilterOnlyRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:M8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 8, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "January"),
					new ExpectedCellValue(sheetName, 5, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 2078.75),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "February"),
					new ExpectedCellValue(sheetName, 5, 12, null),
					new ExpectedCellValue(sheetName, 6, 12, 1194d),
					new ExpectedCellValue(sheetName, 7, 12, 99d),
					new ExpectedCellValue(sheetName, 8, 12, 1293d),
					new ExpectedCellValue(sheetName, 3, 13, null),
					new ExpectedCellValue(sheetName, 4, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 13, 831.5),
					new ExpectedCellValue(sheetName, 6, 13, 2025.5),
					new ExpectedCellValue(sheetName, 7, 13, 514.75),
					new ExpectedCellValue(sheetName, 8, 13, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithRowAndColumnFilterRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E17"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 15, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "February"),
					new ExpectedCellValue(sheetName, 15, 3, null),
					new ExpectedCellValue(sheetName, 16, 3, 1194d),
					new ExpectedCellValue(sheetName, 17, 3, 1194d),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "March"),
					new ExpectedCellValue(sheetName, 15, 4, 24.99),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 856.49),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, 24.99),
					new ExpectedCellValue(sheetName, 16, 5, 2025.5),
					new ExpectedCellValue(sheetName, 17, 5, 2050.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithRowAndColumnFiltersEnabledForAllFieldsTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B22:E27"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 2, null),
					new ExpectedCellValue(sheetName, 24, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 25, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 26, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 27, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 22, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 23, 3, "February"),
					new ExpectedCellValue(sheetName, 24, 3, 20100085),
					new ExpectedCellValue(sheetName, 25, 3, 99d),
					new ExpectedCellValue(sheetName, 26, 3, 99d),
					new ExpectedCellValue(sheetName, 27, 3, 99d),
					new ExpectedCellValue(sheetName, 22, 4, null),
					new ExpectedCellValue(sheetName, 23, 4, "February Total"),
					new ExpectedCellValue(sheetName, 24, 4, null),
					new ExpectedCellValue(sheetName, 25, 4, 99d),
					new ExpectedCellValue(sheetName, 26, 4, 99d),
					new ExpectedCellValue(sheetName, 27, 4, 99d),
					new ExpectedCellValue(sheetName, 22, 5, null),
					new ExpectedCellValue(sheetName, 23, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 24, 5, null),
					new ExpectedCellValue(sheetName, 25, 5, 99d),
					new ExpectedCellValue(sheetName, 26, 5, 99d),
					new ExpectedCellValue(sheetName, 27, 5, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithMultipleRowDataFieldsRowFiltersEnabledTwoRowFieldsOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B32:F47"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 32, 2, null),
					new ExpectedCellValue(sheetName, 33, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 34, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 36, 2, 20100007),
					new ExpectedCellValue(sheetName, 37, 2, 20100083),
					new ExpectedCellValue(sheetName, 38, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 39, 2, 20100085),
					new ExpectedCellValue(sheetName, 40, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 42, 2, 20100007),
					new ExpectedCellValue(sheetName, 43, 2, 20100083),
					new ExpectedCellValue(sheetName, 44, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 45, 2, 20100085),
					new ExpectedCellValue(sheetName, 46, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 33, 3, "January"),
					new ExpectedCellValue(sheetName, 34, 3, null),
					new ExpectedCellValue(sheetName, 35, 3, 415.75),
					new ExpectedCellValue(sheetName, 36, 3, 415.75),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 40, 3, null),
					new ExpectedCellValue(sheetName, 41, 3, 831.5),
					new ExpectedCellValue(sheetName, 42, 3, 831.5),
					new ExpectedCellValue(sheetName, 43, 3, null),
					new ExpectedCellValue(sheetName, 44, 3, null),
					new ExpectedCellValue(sheetName, 45, 3, null),
					new ExpectedCellValue(sheetName, 46, 3, 415.75),
					new ExpectedCellValue(sheetName, 47, 3, 831.5),
					new ExpectedCellValue(sheetName, 32, 4, null),
					new ExpectedCellValue(sheetName, 33, 4, "February"),
					new ExpectedCellValue(sheetName, 34, 4, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 37, 4, null),
					new ExpectedCellValue(sheetName, 38, 4, 99d),
					new ExpectedCellValue(sheetName, 39, 4, 99d),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, null),
					new ExpectedCellValue(sheetName, 44, 4, 99d),
					new ExpectedCellValue(sheetName, 45, 4, 99d),
					new ExpectedCellValue(sheetName, 46, 4, 99d),
					new ExpectedCellValue(sheetName, 47, 4, 99d),
					new ExpectedCellValue(sheetName, 32, 5, null),
					new ExpectedCellValue(sheetName, 33, 5, "March"),
					new ExpectedCellValue(sheetName, 34, 5, null),
					new ExpectedCellValue(sheetName, 35, 5, 24.99),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 37, 5, 24.99),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 39, 5, null),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 41, 5, 24.99),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, 24.99),
					new ExpectedCellValue(sheetName, 44, 5, null),
					new ExpectedCellValue(sheetName, 45, 5, null),
					new ExpectedCellValue(sheetName, 46, 5, 24.99),
					new ExpectedCellValue(sheetName, 47, 5, 24.99),
					new ExpectedCellValue(sheetName, 32, 6, null),
					new ExpectedCellValue(sheetName, 33, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 34, 6, null),
					new ExpectedCellValue(sheetName, 35, 6, 440.74),
					new ExpectedCellValue(sheetName, 36, 6, 415.75),
					new ExpectedCellValue(sheetName, 37, 6, 24.99),
					new ExpectedCellValue(sheetName, 38, 6, 99d),
					new ExpectedCellValue(sheetName, 39, 6, 99d),
					new ExpectedCellValue(sheetName, 40, 6, null),
					new ExpectedCellValue(sheetName, 41, 6, 856.49),
					new ExpectedCellValue(sheetName, 42, 6, 831.5),
					new ExpectedCellValue(sheetName, 43, 6, 24.99),
					new ExpectedCellValue(sheetName, 44, 6, 99d),
					new ExpectedCellValue(sheetName, 45, 6, 99d),
					new ExpectedCellValue(sheetName, 46, 6, 539.74),
					new ExpectedCellValue(sheetName, 47, 6, 955.49),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithMultipleColumnDataFieldsColumnFiltersEnabledOneRowFieldTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J32:V39"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 10, null),
					new ExpectedCellValue(sheetName, 35, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 36, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 37, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 38, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 39, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 32, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 33, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 11, "January"),
					new ExpectedCellValue(sheetName, 35, 11, 20100007),
					new ExpectedCellValue(sheetName, 36, 11, 415.75),
					new ExpectedCellValue(sheetName, 37, 11, null),
					new ExpectedCellValue(sheetName, 38, 11, null),
					new ExpectedCellValue(sheetName, 39, 11, 415.75),
					new ExpectedCellValue(sheetName, 32, 12, null),
					new ExpectedCellValue(sheetName, 33, 12, null),
					new ExpectedCellValue(sheetName, 34, 12, null),
					new ExpectedCellValue(sheetName, 35, 12, 20100090),
					new ExpectedCellValue(sheetName, 36, 12, null),
					new ExpectedCellValue(sheetName, 37, 12, 415.75),
					new ExpectedCellValue(sheetName, 38, 12, null),
					new ExpectedCellValue(sheetName, 39, 12, 415.75),
					new ExpectedCellValue(sheetName, 32, 13, null),
					new ExpectedCellValue(sheetName, 33, 13, null),
					new ExpectedCellValue(sheetName, 34, 13, "January Total"),
					new ExpectedCellValue(sheetName, 35, 13, null),
					new ExpectedCellValue(sheetName, 36, 13, 415.75),
					new ExpectedCellValue(sheetName, 37, 13, 415.75),
					new ExpectedCellValue(sheetName, 38, 13, null),
					new ExpectedCellValue(sheetName, 39, 13, 831.5),
					new ExpectedCellValue(sheetName, 32, 14, null),
					new ExpectedCellValue(sheetName, 33, 14, null),
					new ExpectedCellValue(sheetName, 34, 14, "February"),
					new ExpectedCellValue(sheetName, 35, 14, 20100085),
					new ExpectedCellValue(sheetName, 36, 14, null),
					new ExpectedCellValue(sheetName, 37, 14, null),
					new ExpectedCellValue(sheetName, 38, 14, 99d),
					new ExpectedCellValue(sheetName, 39, 14, 99d),
					new ExpectedCellValue(sheetName, 32, 15, null),
					new ExpectedCellValue(sheetName, 33, 15, null),
					new ExpectedCellValue(sheetName, 34, 15, "February Total"),
					new ExpectedCellValue(sheetName, 35, 15, null),
					new ExpectedCellValue(sheetName, 36, 15, null),
					new ExpectedCellValue(sheetName, 37, 15, null),
					new ExpectedCellValue(sheetName, 38, 15, 99d),
					new ExpectedCellValue(sheetName, 39, 15, 99d),
					new ExpectedCellValue(sheetName, 32, 16, null),
					new ExpectedCellValue(sheetName, 33, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 16, "January"),
					new ExpectedCellValue(sheetName, 35, 16, 20100007),
					new ExpectedCellValue(sheetName, 36, 16, 831.5),
					new ExpectedCellValue(sheetName, 37, 16, null),
					new ExpectedCellValue(sheetName, 38, 16, null),
					new ExpectedCellValue(sheetName, 39, 16, 831.5),
					new ExpectedCellValue(sheetName, 32, 17, null),
					new ExpectedCellValue(sheetName, 33, 17, null),
					new ExpectedCellValue(sheetName, 34, 17, null),
					new ExpectedCellValue(sheetName, 35, 17, 20100090),
					new ExpectedCellValue(sheetName, 36, 17, null),
					new ExpectedCellValue(sheetName, 37, 17, 831.5),
					new ExpectedCellValue(sheetName, 38, 17, null),
					new ExpectedCellValue(sheetName, 39, 17, 831.5),
					new ExpectedCellValue(sheetName, 32, 18, null),
					new ExpectedCellValue(sheetName, 33, 18, null),
					new ExpectedCellValue(sheetName, 34, 18, "January Total"),
					new ExpectedCellValue(sheetName, 35, 18, null),
					new ExpectedCellValue(sheetName, 36, 18, 831.5),
					new ExpectedCellValue(sheetName, 37, 18, 831.5),
					new ExpectedCellValue(sheetName, 38, 18, null),
					new ExpectedCellValue(sheetName, 39, 18, 1663d),
					new ExpectedCellValue(sheetName, 32, 19, null),
					new ExpectedCellValue(sheetName, 33, 19, null),
					new ExpectedCellValue(sheetName, 34, 19, "February"),
					new ExpectedCellValue(sheetName, 35, 19, 20100085),
					new ExpectedCellValue(sheetName, 36, 19, null),
					new ExpectedCellValue(sheetName, 37, 19, null),
					new ExpectedCellValue(sheetName, 38, 19, 99d),
					new ExpectedCellValue(sheetName, 39, 19, 99d),
					new ExpectedCellValue(sheetName, 32, 20, null),
					new ExpectedCellValue(sheetName, 33, 20, null),
					new ExpectedCellValue(sheetName, 34, 20, "February Total"),
					new ExpectedCellValue(sheetName, 35, 20, null),
					new ExpectedCellValue(sheetName, 36, 20, null),
					new ExpectedCellValue(sheetName, 37, 20, null),
					new ExpectedCellValue(sheetName, 38, 20, 99d),
					new ExpectedCellValue(sheetName, 39, 20, 99d),
					new ExpectedCellValue(sheetName, 32, 21, null),
					new ExpectedCellValue(sheetName, 33, 21, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 21, null),
					new ExpectedCellValue(sheetName, 35, 21, null),
					new ExpectedCellValue(sheetName, 36, 21, 415.75),
					new ExpectedCellValue(sheetName, 37, 21, 415.75),
					new ExpectedCellValue(sheetName, 38, 21, 99d),
					new ExpectedCellValue(sheetName, 39, 21, 930.5),
					new ExpectedCellValue(sheetName, 32, 22, null),
					new ExpectedCellValue(sheetName, 33, 22, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 22, null),
					new ExpectedCellValue(sheetName, 35, 22, null),
					new ExpectedCellValue(sheetName, 36, 22, 831.5),
					new ExpectedCellValue(sheetName, 37, 22, 831.5),
					new ExpectedCellValue(sheetName, 38, 22, 99d),
					new ExpectedCellValue(sheetName, 39, 22, 1762d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithWithRegularExpressionRowFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B52:I59"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 52, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 53, 2, null),
					new ExpectedCellValue(sheetName, 54, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 55, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 56, 2, 20100017),
					new ExpectedCellValue(sheetName, 57, 2, 20100070),
					new ExpectedCellValue(sheetName, 58, 2, 20100090),
					new ExpectedCellValue(sheetName, 59, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 52, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 53, 3, "January"),
					new ExpectedCellValue(sheetName, 54, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 831.5),
					new ExpectedCellValue(sheetName, 56, 3, null),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 58, 3, 831.5),
					new ExpectedCellValue(sheetName, 59, 3, 831.5),
					new ExpectedCellValue(sheetName, 52, 4, null),
					new ExpectedCellValue(sheetName, 53, 4, "January Total"),
					new ExpectedCellValue(sheetName, 54, 4, null),
					new ExpectedCellValue(sheetName, 55, 4, 831.5),
					new ExpectedCellValue(sheetName, 56, 4, null),
					new ExpectedCellValue(sheetName, 57, 4, null),
					new ExpectedCellValue(sheetName, 58, 4, 831.5),
					new ExpectedCellValue(sheetName, 59, 4, 831.5),
					new ExpectedCellValue(sheetName, 52, 5, null),
					new ExpectedCellValue(sheetName, 53, 5, "February"),
					new ExpectedCellValue(sheetName, 54, 5, "Tent"),
					new ExpectedCellValue(sheetName, 55, 5, 1194d),
					new ExpectedCellValue(sheetName, 56, 5, null),
					new ExpectedCellValue(sheetName, 57, 5, 1194d),
					new ExpectedCellValue(sheetName, 58, 5, null),
					new ExpectedCellValue(sheetName, 59, 5, 1194d),
					new ExpectedCellValue(sheetName, 52, 6, null),
					new ExpectedCellValue(sheetName, 53, 6, "February Total"),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 6, 1194d),
					new ExpectedCellValue(sheetName, 56, 6, null),
					new ExpectedCellValue(sheetName, 57, 6, 1194d),
					new ExpectedCellValue(sheetName, 58, 6, null),
					new ExpectedCellValue(sheetName, 59, 6, 1194d),
					new ExpectedCellValue(sheetName, 52, 7, null),
					new ExpectedCellValue(sheetName, 53, 7, "March"),
					new ExpectedCellValue(sheetName, 54, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 7, 831.5),
					new ExpectedCellValue(sheetName, 56, 7, 831.5),
					new ExpectedCellValue(sheetName, 57, 7, null),
					new ExpectedCellValue(sheetName, 58, 7, null),
					new ExpectedCellValue(sheetName, 59, 7, 831.5),
					new ExpectedCellValue(sheetName, 52, 8, null),
					new ExpectedCellValue(sheetName, 53, 8, "March Total"),
					new ExpectedCellValue(sheetName, 54, 8, null),
					new ExpectedCellValue(sheetName, 55, 8, 831.5),
					new ExpectedCellValue(sheetName, 56, 8, 831.5),
					new ExpectedCellValue(sheetName, 57, 8, null),
					new ExpectedCellValue(sheetName, 58, 8, null),
					new ExpectedCellValue(sheetName, 59, 8, 831.5),
					new ExpectedCellValue(sheetName, 52, 9, null),
					new ExpectedCellValue(sheetName, 53, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 54, 9, null),
					new ExpectedCellValue(sheetName, 55, 9, 2857d),
					new ExpectedCellValue(sheetName, 56, 9, 831.5),
					new ExpectedCellValue(sheetName, 57, 9, 1194d),
					new ExpectedCellValue(sheetName, 58, 9, 831.5),
					new ExpectedCellValue(sheetName, 59, 9, 2857d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithWithRegularExpressionColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("K52:N57"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 52, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 53, 11, null),
					new ExpectedCellValue(sheetName, 54, 11, "Row Labels"),
					new ExpectedCellValue(sheetName, 55, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 56, 11, 20100085),
					new ExpectedCellValue(sheetName, 57, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 52, 12, "Column Labels"),
					new ExpectedCellValue(sheetName, 53, 12, "February"),
					new ExpectedCellValue(sheetName, 54, 12, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 55, 12, 99d),
					new ExpectedCellValue(sheetName, 56, 12, 99d),
					new ExpectedCellValue(sheetName, 57, 12, 99d),
					new ExpectedCellValue(sheetName, 52, 13, null),
					new ExpectedCellValue(sheetName, 53, 13, "February Total"),
					new ExpectedCellValue(sheetName, 54, 13, null),
					new ExpectedCellValue(sheetName, 55, 13, 99d),
					new ExpectedCellValue(sheetName, 56, 13, 99d),
					new ExpectedCellValue(sheetName, 57, 13, 99d),
					new ExpectedCellValue(sheetName, 52, 14, null),
					new ExpectedCellValue(sheetName, 53, 14, "Grand Total"),
					new ExpectedCellValue(sheetName, 54, 14, null),
					new ExpectedCellValue(sheetName, 55, 14, 99d),
					new ExpectedCellValue(sheetName, 56, 14, 99d),
					new ExpectedCellValue(sheetName, 57, 14, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable9"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B72:E77"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 72, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 73, 2, null),
					new ExpectedCellValue(sheetName, 74, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 75, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 76, 2, 20100070),
					new ExpectedCellValue(sheetName, 77, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 72, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 73, 3, "February"),
					new ExpectedCellValue(sheetName, 74, 3, "Tent"),
					new ExpectedCellValue(sheetName, 75, 3, 1194d),
					new ExpectedCellValue(sheetName, 76, 3, 1194d),
					new ExpectedCellValue(sheetName, 77, 3, 1194d),
					new ExpectedCellValue(sheetName, 72, 4, null),
					new ExpectedCellValue(sheetName, 73, 4, "February Total"),
					new ExpectedCellValue(sheetName, 74, 4, null),
					new ExpectedCellValue(sheetName, 75, 4, 1194d),
					new ExpectedCellValue(sheetName, 76, 4, 1194d),
					new ExpectedCellValue(sheetName, 77, 4, 1194d),
					new ExpectedCellValue(sheetName, 72, 5, null),
					new ExpectedCellValue(sheetName, 73, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 74, 5, null),
					new ExpectedCellValue(sheetName, 75, 5, 1194d),
					new ExpectedCellValue(sheetName, 76, 5, 1194d),
					new ExpectedCellValue(sheetName, 77, 5, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBeginsWithWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBeginsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable10"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("K72:O79"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 72, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 73, 11, null),
					new ExpectedCellValue(sheetName, 74, 11, "Row Labels"),
					new ExpectedCellValue(sheetName, 75, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 76, 11, 20100070),
					new ExpectedCellValue(sheetName, 77, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 78, 11, 20100085),
					new ExpectedCellValue(sheetName, 79, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 72, 12, "Column Labels"),
					new ExpectedCellValue(sheetName, 73, 12, "February"),
					new ExpectedCellValue(sheetName, 74, 12, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 75, 12, null),
					new ExpectedCellValue(sheetName, 76, 12, null),
					new ExpectedCellValue(sheetName, 77, 12, 99d),
					new ExpectedCellValue(sheetName, 78, 12, 99d),
					new ExpectedCellValue(sheetName, 79, 12, 99d),
					new ExpectedCellValue(sheetName, 72, 13, null),
					new ExpectedCellValue(sheetName, 73, 13, null),
					new ExpectedCellValue(sheetName, 74, 13, "Tent"),
					new ExpectedCellValue(sheetName, 75, 13, 1194d),
					new ExpectedCellValue(sheetName, 76, 13, 1194d),
					new ExpectedCellValue(sheetName, 77, 13, null),
					new ExpectedCellValue(sheetName, 78, 13, null),
					new ExpectedCellValue(sheetName, 79, 13, 1194d),
					new ExpectedCellValue(sheetName, 72, 14, null),
					new ExpectedCellValue(sheetName, 73, 14, "February Total"),
					new ExpectedCellValue(sheetName, 74, 14, null),
					new ExpectedCellValue(sheetName, 75, 14, 1194d),
					new ExpectedCellValue(sheetName, 76, 14, 1194d),
					new ExpectedCellValue(sheetName, 77, 14, 99d),
					new ExpectedCellValue(sheetName, 78, 14, 99d),
					new ExpectedCellValue(sheetName, 79, 14, 1293d),
					new ExpectedCellValue(sheetName, 72, 15, null),
					new ExpectedCellValue(sheetName, 73, 15, "Grand Total"),
					new ExpectedCellValue(sheetName, 74, 15, null),
					new ExpectedCellValue(sheetName, 75, 15, 1194d),
					new ExpectedCellValue(sheetName, 76, 15, 1194d),
					new ExpectedCellValue(sheetName, 77, 15, 99d),
					new ExpectedCellValue(sheetName, 78, 15, 99d),
					new ExpectedCellValue(sheetName, 79, 15, 1293d)
				});
			}
		}
		#endregion

		#region CaptionEndsWith Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, null),
					new ExpectedCellValue(sheetName, 6, 4, 99d),
					new ExpectedCellValue(sheetName, 7, 4, 99d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 24.99),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 7, 5, 24.99),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 856.49),
					new ExpectedCellValue(sheetName, 6, 6, 514.75),
					new ExpectedCellValue(sheetName, 7, 6, 1371.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "March"),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 856.49),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithRowAndColumnFilterRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E17"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 15, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 17, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "January"),
					new ExpectedCellValue(sheetName, 15, 3, 831.5),
					new ExpectedCellValue(sheetName, 16, 3, 415.75),
					new ExpectedCellValue(sheetName, 17, 3, 1247.25),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "February"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 99d),
					new ExpectedCellValue(sheetName, 17, 4, 99d),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, 831.5),
					new ExpectedCellValue(sheetName, 16, 5, 514.75),
					new ExpectedCellValue(sheetName, 17, 5, 1346.25)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithRowAndColumnFiltersEnabledForAllFieldsTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B22:E27"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 2, null),
					new ExpectedCellValue(sheetName, 24, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 25, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 26, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 27, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 22, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 23, 3, "January"),
					new ExpectedCellValue(sheetName, 24, 3, 20100076),
					new ExpectedCellValue(sheetName, 25, 3, 415.75),
					new ExpectedCellValue(sheetName, 26, 3, 415.75),
					new ExpectedCellValue(sheetName, 27, 3, 415.75),
					new ExpectedCellValue(sheetName, 22, 4, null),
					new ExpectedCellValue(sheetName, 23, 4, "January Total"),
					new ExpectedCellValue(sheetName, 24, 4, null),
					new ExpectedCellValue(sheetName, 25, 4, 415.75),
					new ExpectedCellValue(sheetName, 26, 4, 415.75),
					new ExpectedCellValue(sheetName, 27, 4, 415.75),
					new ExpectedCellValue(sheetName, 22, 5, null),
					new ExpectedCellValue(sheetName, 23, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 24, 5, null),
					new ExpectedCellValue(sheetName, 25, 5, 415.75),
					new ExpectedCellValue(sheetName, 26, 5, 415.75),
					new ExpectedCellValue(sheetName, 27, 5, 415.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithMultipleRowDataFieldsRowFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B33:G46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, null),
					new ExpectedCellValue(sheetName, 34, 2, null),
					new ExpectedCellValue(sheetName, 35, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 36, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 37, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 38, 2, 20100007),
					new ExpectedCellValue(sheetName, 39, 2, 20100017),
					new ExpectedCellValue(sheetName, 40, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 2, 20100007),
					new ExpectedCellValue(sheetName, 42, 2, 20100017),
					new ExpectedCellValue(sheetName, 43, 2, "Car Rack Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 44, 2, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 45, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 46, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 34, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 35, 3, "January"),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 38, 3, 415.75),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 40, 3, null),
					new ExpectedCellValue(sheetName, 41, 3, 831.5),
					new ExpectedCellValue(sheetName, 42, 3, null),
					new ExpectedCellValue(sheetName, 43, 3, 415.75),
					new ExpectedCellValue(sheetName, 44, 3, 831.5),
					new ExpectedCellValue(sheetName, 45, 3, 415.75),
					new ExpectedCellValue(sheetName, 46, 3, 831.5),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 34, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 37, 4, null),
					new ExpectedCellValue(sheetName, 38, 4, 415.75),
					new ExpectedCellValue(sheetName, 39, 4, null),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 41, 4, 831.5),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, 415.75),
					new ExpectedCellValue(sheetName, 44, 4, 831.5),
					new ExpectedCellValue(sheetName, 45, 4, 415.75),
					new ExpectedCellValue(sheetName, 46, 4, 831.5),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 34, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 35, 5, "March"),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 37, 5, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 39, 5, 415.75),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 42, 5, 831.5),
					new ExpectedCellValue(sheetName, 43, 5, 415.75),
					new ExpectedCellValue(sheetName, 44, 5, 831.5),
					new ExpectedCellValue(sheetName, 45, 5, 415.75),
					new ExpectedCellValue(sheetName, 46, 5, 831.5),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 34, 6, "Nashville Total"),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 39, 6, 415.75),
					new ExpectedCellValue(sheetName, 40, 6, null),
					new ExpectedCellValue(sheetName, 41, 6, null),
					new ExpectedCellValue(sheetName, 42, 6, 831.5),
					new ExpectedCellValue(sheetName, 43, 6, 415.75),
					new ExpectedCellValue(sheetName, 44, 6, 831.5),
					new ExpectedCellValue(sheetName, 45, 6, 415.75),
					new ExpectedCellValue(sheetName, 46, 6, 831.5),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 34, 7, "Grand Total"),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 37, 7, null),
					new ExpectedCellValue(sheetName, 38, 7, 415.75),
					new ExpectedCellValue(sheetName, 39, 7, 415.75),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 41, 7, 831.5),
					new ExpectedCellValue(sheetName, 42, 7, 831.5),
					new ExpectedCellValue(sheetName, 43, 7, 831.5),
					new ExpectedCellValue(sheetName, 44, 7, 1663d),
					new ExpectedCellValue(sheetName, 45, 7, 831.5),
					new ExpectedCellValue(sheetName, 46, 7, 1663d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithMultipleColumnDataFieldsColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J33:M44"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 10, null),
					new ExpectedCellValue(sheetName, 35, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 36, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 37, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 38, 10, 20100017),
					new ExpectedCellValue(sheetName, 39, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 40, 10, 20100017),
					new ExpectedCellValue(sheetName, 41, 10, "Car Rack Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 42, 10, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 44, 10, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 34, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 35, 11, "March"),
					new ExpectedCellValue(sheetName, 36, 11, null),
					new ExpectedCellValue(sheetName, 37, 11, null),
					new ExpectedCellValue(sheetName, 38, 11, 415.75),
					new ExpectedCellValue(sheetName, 39, 11, null),
					new ExpectedCellValue(sheetName, 40, 11, 831.5),
					new ExpectedCellValue(sheetName, 41, 11, 415.75),
					new ExpectedCellValue(sheetName, 42, 11, 831.5),
					new ExpectedCellValue(sheetName, 43, 11, 415.75),
					new ExpectedCellValue(sheetName, 44, 11, 831.5),
					new ExpectedCellValue(sheetName, 33, 12, null),
					new ExpectedCellValue(sheetName, 34, 12, "Nashville Total"),
					new ExpectedCellValue(sheetName, 35, 12, null),
					new ExpectedCellValue(sheetName, 36, 12, null),
					new ExpectedCellValue(sheetName, 37, 12, null),
					new ExpectedCellValue(sheetName, 38, 12, 415.75),
					new ExpectedCellValue(sheetName, 39, 12, null),
					new ExpectedCellValue(sheetName, 40, 12, 831.5),
					new ExpectedCellValue(sheetName, 41, 12, 415.75),
					new ExpectedCellValue(sheetName, 42, 12, 831.5),
					new ExpectedCellValue(sheetName, 43, 12, 415.75),
					new ExpectedCellValue(sheetName, 44, 12, 831.5),
					new ExpectedCellValue(sheetName, 33, 13, null),
					new ExpectedCellValue(sheetName, 34, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 35, 13, null),
					new ExpectedCellValue(sheetName, 36, 13, null),
					new ExpectedCellValue(sheetName, 37, 13, null),
					new ExpectedCellValue(sheetName, 38, 13, 415.75),
					new ExpectedCellValue(sheetName, 39, 13, null),
					new ExpectedCellValue(sheetName, 40, 13, 831.5),
					new ExpectedCellValue(sheetName, 41, 13, 415.75),
					new ExpectedCellValue(sheetName, 42, 13, 831.5),
					new ExpectedCellValue(sheetName, 43, 13, 415.75),
					new ExpectedCellValue(sheetName, 44, 13, 831.5),

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithWithRegularExpressionRowFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B51:G57"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 51, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 52, 2, null),
					new ExpectedCellValue(sheetName, 53, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 54, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 55, 2, 20100017),
					new ExpectedCellValue(sheetName, 56, 2, 20100070),
					new ExpectedCellValue(sheetName, 57, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 51, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 52, 3, "February"),
					new ExpectedCellValue(sheetName, 53, 3, "Tent"),
					new ExpectedCellValue(sheetName, 54, 3, 1194d),
					new ExpectedCellValue(sheetName, 55, 3, null),
					new ExpectedCellValue(sheetName, 56, 3, 1194d),
					new ExpectedCellValue(sheetName, 57, 3, 1194d),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 52, 4, "February Total"),
					new ExpectedCellValue(sheetName, 53, 4, null),
					new ExpectedCellValue(sheetName, 54, 4, 1194d),
					new ExpectedCellValue(sheetName, 55, 4, null),
					new ExpectedCellValue(sheetName, 56, 4, 1194d),
					new ExpectedCellValue(sheetName, 57, 4, 1194d),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 5, "March"),
					new ExpectedCellValue(sheetName, 53, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 54, 5, 831.5),
					new ExpectedCellValue(sheetName, 55, 5, 831.5),
					new ExpectedCellValue(sheetName, 56, 5, null),
					new ExpectedCellValue(sheetName, 57, 5, 831.5),
					new ExpectedCellValue(sheetName, 51, 6, null),
					new ExpectedCellValue(sheetName, 52, 6, "March Total"),
					new ExpectedCellValue(sheetName, 53, 6, null),
					new ExpectedCellValue(sheetName, 54, 6, 831.5),
					new ExpectedCellValue(sheetName, 55, 6, 831.5),
					new ExpectedCellValue(sheetName, 56, 6, null),
					new ExpectedCellValue(sheetName, 57, 6, 831.5),
					new ExpectedCellValue(sheetName, 51, 7, null),
					new ExpectedCellValue(sheetName, 52, 7, "Grand Total"),
					new ExpectedCellValue(sheetName, 53, 7, null),
					new ExpectedCellValue(sheetName, 54, 7, 2025.5),
					new ExpectedCellValue(sheetName, 55, 7, 831.5),
					new ExpectedCellValue(sheetName, 56, 7, 1194d),
					new ExpectedCellValue(sheetName, 57, 7, 2025.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithWithRegularExpressionColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J51:M60"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 51, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 52, 10, null),
					new ExpectedCellValue(sheetName, 53, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 54, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 55, 10, 20100007),
					new ExpectedCellValue(sheetName, 56, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 57, 10, 20100090),
					new ExpectedCellValue(sheetName, 58, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 59, 10, 20100076),
					new ExpectedCellValue(sheetName, 60, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 51, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 52, 11, "January"),
					new ExpectedCellValue(sheetName, 53, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 54, 11, 831.5),
					new ExpectedCellValue(sheetName, 55, 11, 831.5),
					new ExpectedCellValue(sheetName, 56, 11, 831.5),
					new ExpectedCellValue(sheetName, 57, 11, 831.5),
					new ExpectedCellValue(sheetName, 58, 11, 415.75),
					new ExpectedCellValue(sheetName, 59, 11, 415.75),
					new ExpectedCellValue(sheetName, 60, 11, 2078.75),
					new ExpectedCellValue(sheetName, 51, 12, null),
					new ExpectedCellValue(sheetName, 52, 12, "January Total"),
					new ExpectedCellValue(sheetName, 53, 12, null),
					new ExpectedCellValue(sheetName, 54, 12, 831.5),
					new ExpectedCellValue(sheetName, 55, 12, 831.5),
					new ExpectedCellValue(sheetName, 56, 12, 831.5),
					new ExpectedCellValue(sheetName, 57, 12, 831.5),
					new ExpectedCellValue(sheetName, 58, 12, 415.75),
					new ExpectedCellValue(sheetName, 59, 12, 415.75),
					new ExpectedCellValue(sheetName, 60, 12, 2078.75),
					new ExpectedCellValue(sheetName, 51, 13, null),
					new ExpectedCellValue(sheetName, 52, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 53, 13, null),
					new ExpectedCellValue(sheetName, 54, 13, 831.5),
					new ExpectedCellValue(sheetName, 55, 13, 831.5),
					new ExpectedCellValue(sheetName, 56, 13, 831.5),
					new ExpectedCellValue(sheetName, 57, 13, 831.5),
					new ExpectedCellValue(sheetName, 58, 13, 415.75),
					new ExpectedCellValue(sheetName, 59, 13, 415.75),
					new ExpectedCellValue(sheetName, 60, 13, 2078.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable9"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B66:E71"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 66, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 67, 2, null),
					new ExpectedCellValue(sheetName, 68, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 69, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 70, 2, 20100007),
					new ExpectedCellValue(sheetName, 71, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 66, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 67, 3, "January"),
					new ExpectedCellValue(sheetName, 68, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 69, 3, 831.5),
					new ExpectedCellValue(sheetName, 70, 3, 831.5),
					new ExpectedCellValue(sheetName, 71, 3, 831.5),
					new ExpectedCellValue(sheetName, 66, 4, null),
					new ExpectedCellValue(sheetName, 67, 4, "January Total"),
					new ExpectedCellValue(sheetName, 68, 4, null),
					new ExpectedCellValue(sheetName, 69, 4, 831.5),
					new ExpectedCellValue(sheetName, 70, 4, 831.5),
					new ExpectedCellValue(sheetName, 71, 4, 831.5),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 67, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 68, 5, null),
					new ExpectedCellValue(sheetName, 69, 5, 831.5),
					new ExpectedCellValue(sheetName, 70, 5, 831.5),
					new ExpectedCellValue(sheetName, 71, 5, 831.5),

				});
			}
		}
		#endregion

		#region CaptionNotEndsWith Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 1194d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 2857d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "March"),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 856.49),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithRowAndColumnFilterRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B10:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 10, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 11, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 12, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 10, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 11, 3, "March"),
					new ExpectedCellValue(sheetName, 12, 3, 831.5),
					new ExpectedCellValue(sheetName, 13, 3, 831.5),
					new ExpectedCellValue(sheetName, 10, 4, null),
					new ExpectedCellValue(sheetName, 11, 4, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 4, 831.5),
					new ExpectedCellValue(sheetName, 13, 4, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithRowAndColumnFiltersEnabledForAllFieldsTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B17:E22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 17, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 18, 2, null),
					new ExpectedCellValue(sheetName, 19, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 20, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 21, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 22, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 17, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 18, 3, "February"),
					new ExpectedCellValue(sheetName, 19, 3, 20100085),
					new ExpectedCellValue(sheetName, 20, 3, 99d),
					new ExpectedCellValue(sheetName, 21, 3, 99d),
					new ExpectedCellValue(sheetName, 22, 3, 99d),
					new ExpectedCellValue(sheetName, 17, 4, null),
					new ExpectedCellValue(sheetName, 18, 4, "February Total"),
					new ExpectedCellValue(sheetName, 19, 4, null),
					new ExpectedCellValue(sheetName, 20, 4, 99d),
					new ExpectedCellValue(sheetName, 21, 4, 99d),
					new ExpectedCellValue(sheetName, 22, 4, 99d),
					new ExpectedCellValue(sheetName, 17, 5, null),
					new ExpectedCellValue(sheetName, 18, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 20, 5, 99d),
					new ExpectedCellValue(sheetName, 21, 5, 99d),
					new ExpectedCellValue(sheetName, 22, 5, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithMultipleRowDataFieldsRowFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B27:E37"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 27, 2, null),
					new ExpectedCellValue(sheetName, 28, 2, null),
					new ExpectedCellValue(sheetName, 29, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 30, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 31, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 32, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 2, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 35, 2, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 37, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 28, 3, "March"),
					new ExpectedCellValue(sheetName, 29, 3, 20100017),
					new ExpectedCellValue(sheetName, 30, 3, null),
					new ExpectedCellValue(sheetName, 31, 3, null),
					new ExpectedCellValue(sheetName, 32, 3, 831.5),
					new ExpectedCellValue(sheetName, 33, 3, 415.75),
					new ExpectedCellValue(sheetName, 34, 3, 831.5),
					new ExpectedCellValue(sheetName, 35, 3, 415.75),
					new ExpectedCellValue(sheetName, 36, 3, 831.5),
					new ExpectedCellValue(sheetName, 37, 3, 415.75),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, "March Total"),
					new ExpectedCellValue(sheetName, 29, 4, null),
					new ExpectedCellValue(sheetName, 30, 4, null),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 32, 4, 831.5),
					new ExpectedCellValue(sheetName, 33, 4, 415.75),
					new ExpectedCellValue(sheetName, 34, 4, 831.5),
					new ExpectedCellValue(sheetName, 35, 4, 415.75),
					new ExpectedCellValue(sheetName, 36, 4, 831.5),
					new ExpectedCellValue(sheetName, 37, 4, 415.75),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 32, 5, 831.5),
					new ExpectedCellValue(sheetName, 33, 5, 415.75),
					new ExpectedCellValue(sheetName, 34, 5, 831.5),
					new ExpectedCellValue(sheetName, 35, 5, 415.75),
					new ExpectedCellValue(sheetName, 36, 5, 831.5),
					new ExpectedCellValue(sheetName, 37, 5, 415.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithMultipleColumnDataFieldsColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J27:P33"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 27, 10, null),
					new ExpectedCellValue(sheetName, 28, 10, null),
					new ExpectedCellValue(sheetName, 29, 10, null),
					new ExpectedCellValue(sheetName, 30, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 31, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 33, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 27, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 28, 11, "March"),
					new ExpectedCellValue(sheetName, 29, 11, 20100017),
					new ExpectedCellValue(sheetName, 30, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 11, 831.5),
					new ExpectedCellValue(sheetName, 32, 11, 831.5),
					new ExpectedCellValue(sheetName, 33, 11, 831.5),
					new ExpectedCellValue(sheetName, 27, 12, null),
					new ExpectedCellValue(sheetName, 28, 12, null),
					new ExpectedCellValue(sheetName, 29, 12, null),
					new ExpectedCellValue(sheetName, 30, 12, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 12, 415.75),
					new ExpectedCellValue(sheetName, 32, 12, 415.75),
					new ExpectedCellValue(sheetName, 33, 12, 415.75),
					new ExpectedCellValue(sheetName, 27, 13, null),
					new ExpectedCellValue(sheetName, 28, 13, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 29, 13, null),
					new ExpectedCellValue(sheetName, 30, 13, null),
					new ExpectedCellValue(sheetName, 31, 13, 831.5),
					new ExpectedCellValue(sheetName, 32, 13, 831.5),
					new ExpectedCellValue(sheetName, 33, 13, 831.5),
					new ExpectedCellValue(sheetName, 27, 14, null),
					new ExpectedCellValue(sheetName, 28, 14, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 29, 14, null),
					new ExpectedCellValue(sheetName, 30, 14, null),
					new ExpectedCellValue(sheetName, 32, 14, 415.75),
					new ExpectedCellValue(sheetName, 31, 14, 415.75),
					new ExpectedCellValue(sheetName, 33, 14, 415.75),
					new ExpectedCellValue(sheetName, 27, 15, null),
					new ExpectedCellValue(sheetName, 28, 15, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 29, 15, null),
					new ExpectedCellValue(sheetName, 30, 15, null),
					new ExpectedCellValue(sheetName, 31, 15, 831.5),
					new ExpectedCellValue(sheetName, 32, 15, 831.5),
					new ExpectedCellValue(sheetName, 33, 15, 831.5),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 29, 16, null),
					new ExpectedCellValue(sheetName, 30, 16, null),
					new ExpectedCellValue(sheetName, 31, 16, 415.75),
					new ExpectedCellValue(sheetName, 32, 16, 415.75),
					new ExpectedCellValue(sheetName, 33, 16, 415.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B42:E47"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 42, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 2, null),
					new ExpectedCellValue(sheetName, 44, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 45, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 46, 2, "February"),
					new ExpectedCellValue(sheetName, 47, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 43, 3, "Tent"),
					new ExpectedCellValue(sheetName, 44, 3, 20100070),
					new ExpectedCellValue(sheetName, 45, 3, 1194d),
					new ExpectedCellValue(sheetName, 46, 3, 1194d),
					new ExpectedCellValue(sheetName, 47, 3, 1194d),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, "Tent Total"),
					new ExpectedCellValue(sheetName, 44, 4, null),
					new ExpectedCellValue(sheetName, 45, 4, 1194d),
					new ExpectedCellValue(sheetName, 46, 4, 1194d),
					new ExpectedCellValue(sheetName, 47, 4, 1194d),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 44, 5, null),
					new ExpectedCellValue(sheetName, 45, 5, 1194d),
					new ExpectedCellValue(sheetName, 46, 5, 1194d),
					new ExpectedCellValue(sheetName, 47, 5, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotEndsWithWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotEndsWith";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J42:M47"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 42, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 10, null),
					new ExpectedCellValue(sheetName, 44, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 45, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 46, 10, "January"),
					new ExpectedCellValue(sheetName, 47, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 43, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 44, 11, 20100090),
					new ExpectedCellValue(sheetName, 45, 11, 831.5),
					new ExpectedCellValue(sheetName, 46, 11, 831.5),
					new ExpectedCellValue(sheetName, 47, 11, 831.5),
					new ExpectedCellValue(sheetName, 42, 12, null),
					new ExpectedCellValue(sheetName, 43, 12, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 44, 12, null),
					new ExpectedCellValue(sheetName, 45, 12, 831.5),
					new ExpectedCellValue(sheetName, 46, 12, 831.5),
					new ExpectedCellValue(sheetName, 47, 12, 831.5),
					new ExpectedCellValue(sheetName, 42, 13, null),
					new ExpectedCellValue(sheetName, 43, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 44, 13, null),
					new ExpectedCellValue(sheetName, 45, 13, 831.5),
					new ExpectedCellValue(sheetName, 46, 13, 831.5),
					new ExpectedCellValue(sheetName, 47, 13, 831.5)
				});
			}
		}
		#endregion

		#region CaptionContains
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionContainsWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 99d),
					new ExpectedCellValue(sheetName, 7, 4, 1293d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 514.75),
					new ExpectedCellValue(sheetName, 7, 6, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionContainsWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:M8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 8, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "January"),
					new ExpectedCellValue(sheetName, 5, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 2078.75),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "February"),
					new ExpectedCellValue(sheetName, 5, 12, null),
					new ExpectedCellValue(sheetName, 6, 12, 1194d),
					new ExpectedCellValue(sheetName, 7, 12, 99d),
					new ExpectedCellValue(sheetName, 8, 12, 1293d),
					new ExpectedCellValue(sheetName, 3, 13, null),
					new ExpectedCellValue(sheetName, 4, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 13, 831.5),
					new ExpectedCellValue(sheetName, 6, 13, 2025.5),
					new ExpectedCellValue(sheetName, 7, 13, 514.75),
					new ExpectedCellValue(sheetName, 8, 13, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionContainsWithRowAndColumnFilterTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E20"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 16, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 17, 2, 20100007),
					new ExpectedCellValue(sheetName, 18, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 19, 2, 20100076),
					new ExpectedCellValue(sheetName, 20, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "January"),
					new ExpectedCellValue(sheetName, 15, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 18, 3, 415.75),
					new ExpectedCellValue(sheetName, 19, 3, 415.75),
					new ExpectedCellValue(sheetName, 20, 3, 1247.25),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "January Total"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 831.5),
					new ExpectedCellValue(sheetName, 18, 4, 415.75),
					new ExpectedCellValue(sheetName, 19, 4, 415.75),
					new ExpectedCellValue(sheetName, 20, 4, 1247.25),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5),
					new ExpectedCellValue(sheetName, 18, 5, 415.75),
					new ExpectedCellValue(sheetName, 19, 5, 415.75),
					new ExpectedCellValue(sheetName, 20, 5, 1247.25)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionEndsWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B25:E35"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, null),
					new ExpectedCellValue(sheetName, 27, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 28, 2, "February"),
					new ExpectedCellValue(sheetName, 29, 2, 20100070),
					new ExpectedCellValue(sheetName, 30, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 2, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 26, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 27, 3, "Tent"),
					new ExpectedCellValue(sheetName, 28, 3, null),
					new ExpectedCellValue(sheetName, 29, 3, null),
					new ExpectedCellValue(sheetName, 30, 3, 199d),
					new ExpectedCellValue(sheetName, 31, 3, 1194d),
					new ExpectedCellValue(sheetName, 32, 3, 199d),
					new ExpectedCellValue(sheetName, 33, 3, 1194d),
					new ExpectedCellValue(sheetName, 34, 3, 199d),
					new ExpectedCellValue(sheetName, 35, 3, 1194d),
					new ExpectedCellValue(sheetName, 25, 4, null),
					new ExpectedCellValue(sheetName, 26, 4, "Nashville Total"),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, null),
					new ExpectedCellValue(sheetName, 29, 4, null),
					new ExpectedCellValue(sheetName, 30, 4, 199d),
					new ExpectedCellValue(sheetName, 31, 4, 1194d),
					new ExpectedCellValue(sheetName, 32, 4, 199d),
					new ExpectedCellValue(sheetName, 33, 4, 1194d),
					new ExpectedCellValue(sheetName, 34, 4, 199d),
					new ExpectedCellValue(sheetName, 35, 4, 1194d),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, null),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, 199d),
					new ExpectedCellValue(sheetName, 31, 5, 1194d),
					new ExpectedCellValue(sheetName, 32, 5, 199d),
					new ExpectedCellValue(sheetName, 33, 5, 1194d),
					new ExpectedCellValue(sheetName, 34, 5, 199d),
					new ExpectedCellValue(sheetName, 35, 5, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionContainsWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J25:P31"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, null),
					new ExpectedCellValue(sheetName, 28, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 29, 10, "February"),
					new ExpectedCellValue(sheetName, 30, 10, 20100085),
					new ExpectedCellValue(sheetName, 31, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 25, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 26, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 27, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 28, 11, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 29, 11, 99d),
					new ExpectedCellValue(sheetName, 30, 11, 99d),
					new ExpectedCellValue(sheetName, 31, 11, 99d),
					new ExpectedCellValue(sheetName, 25, 12, null),
					new ExpectedCellValue(sheetName, 26, 12, null),
					new ExpectedCellValue(sheetName, 27, 12, "Sum of Total"),
					new ExpectedCellValue(sheetName, 28, 12, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 29, 12, 99d),
					new ExpectedCellValue(sheetName, 30, 12, 99d),
					new ExpectedCellValue(sheetName, 31, 12, 99d),
					new ExpectedCellValue(sheetName, 25, 13, null),
					new ExpectedCellValue(sheetName, 26, 13, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 13, null),
					new ExpectedCellValue(sheetName, 28, 13, null),
					new ExpectedCellValue(sheetName, 29, 13, 99d),
					new ExpectedCellValue(sheetName, 30, 13, 99d),
					new ExpectedCellValue(sheetName, 31, 13, 99d),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 14, null),
					new ExpectedCellValue(sheetName, 28, 14, null),
					new ExpectedCellValue(sheetName, 29, 14, 99d),
					new ExpectedCellValue(sheetName, 30, 14, 99d),
					new ExpectedCellValue(sheetName, 31, 14, 99d),
					new ExpectedCellValue(sheetName, 25, 15, null),
					new ExpectedCellValue(sheetName, 26, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 15, null),
					new ExpectedCellValue(sheetName, 28, 15, null),
					new ExpectedCellValue(sheetName, 29, 15, 99d),
					new ExpectedCellValue(sheetName, 30, 15, 99d),
					new ExpectedCellValue(sheetName, 31, 15, 99d),
					new ExpectedCellValue(sheetName, 25, 16, null),
					new ExpectedCellValue(sheetName, 26, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, null),
					new ExpectedCellValue(sheetName, 29, 16, 99d),
					new ExpectedCellValue(sheetName, 30, 16, 99d),
					new ExpectedCellValue(sheetName, 31, 16, 99d)

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionContainsWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B40:E47"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 2, null),
					new ExpectedCellValue(sheetName, 42, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 43, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 44, 2, 20100007),
					new ExpectedCellValue(sheetName, 45, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 46, 2, 20100076),
					new ExpectedCellValue(sheetName, 47, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 41, 3, "January"),
					new ExpectedCellValue(sheetName, 42, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 43, 3, 831.5),
					new ExpectedCellValue(sheetName, 44, 3, 831.5),
					new ExpectedCellValue(sheetName, 45, 3, 415.75),
					new ExpectedCellValue(sheetName, 46, 3, 415.75),
					new ExpectedCellValue(sheetName, 47, 3, 1247.25),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 41, 4, "January Total"),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, 831.5),
					new ExpectedCellValue(sheetName, 44, 4, 831.5),
					new ExpectedCellValue(sheetName, 45, 4, 415.75),
					new ExpectedCellValue(sheetName, 46, 4, 415.75),
					new ExpectedCellValue(sheetName, 47, 4, 1247.25),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 41, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, 831.5),
					new ExpectedCellValue(sheetName, 44, 5, 831.5),
					new ExpectedCellValue(sheetName, 45, 5, 415.75),
					new ExpectedCellValue(sheetName, 46, 5, 415.75),
					new ExpectedCellValue(sheetName, 47, 5, 1247.25)

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionContainsWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J40:M45"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 10, null),
					new ExpectedCellValue(sheetName, 42, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 43, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 44, 10, 20100083),
					new ExpectedCellValue(sheetName, 45, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 41, 11, "March"),
					new ExpectedCellValue(sheetName, 42, 11, "Headlamp"),
					new ExpectedCellValue(sheetName, 43, 11, 24.99),
					new ExpectedCellValue(sheetName, 44, 11, 24.99),
					new ExpectedCellValue(sheetName, 45, 11, 24.99),
					new ExpectedCellValue(sheetName, 40, 12, null),
					new ExpectedCellValue(sheetName, 41, 12, "March Total"),
					new ExpectedCellValue(sheetName, 42, 12, null),
					new ExpectedCellValue(sheetName, 43, 12, 24.99),
					new ExpectedCellValue(sheetName, 44, 12, 24.99),
					new ExpectedCellValue(sheetName, 45, 12, 24.99),
					new ExpectedCellValue(sheetName, 40, 13, null),
					new ExpectedCellValue(sheetName, 41, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 13, null),
					new ExpectedCellValue(sheetName, 43, 13, 24.99),
					new ExpectedCellValue(sheetName, 44, 13, 24.99),
					new ExpectedCellValue(sheetName, 45, 13, 24.99)
				});
			}
		}
		#endregion

		#region CaptionNotContains Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:E6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "March"),
					new ExpectedCellValue(sheetName, 5, 4, 24.99),
					new ExpectedCellValue(sheetName, 6, 4, 24.99),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 856.49),
					new ExpectedCellValue(sheetName, 6, 5, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "March"),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 856.49),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithRowAndColumnFilterTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 16, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 2, 20100090),
					new ExpectedCellValue(sheetName, 18, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "January"),
					new ExpectedCellValue(sheetName, 15, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 18, 3, 831.5),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "January Total"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 831.5),
					new ExpectedCellValue(sheetName, 18, 4, 831.5),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5),
					new ExpectedCellValue(sheetName, 18, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B25:E35"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, null),
					new ExpectedCellValue(sheetName, 27, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 28, 2, "March"),
					new ExpectedCellValue(sheetName, 29, 2, 20100017),
					new ExpectedCellValue(sheetName, 30, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 2, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 26, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 27, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 28, 3, null),
					new ExpectedCellValue(sheetName, 29, 3, null),
					new ExpectedCellValue(sheetName, 30, 3, 415.75),
					new ExpectedCellValue(sheetName, 31, 3, 831.5),
					new ExpectedCellValue(sheetName, 32, 3, 415.75),
					new ExpectedCellValue(sheetName, 33, 3, 831.5),
					new ExpectedCellValue(sheetName, 34, 3, 415.75),
					new ExpectedCellValue(sheetName, 35, 3, 831.5),
					new ExpectedCellValue(sheetName, 25, 4, null),
					new ExpectedCellValue(sheetName, 26, 4, "Nashville Total"),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, null),
					new ExpectedCellValue(sheetName, 29, 4, null),
					new ExpectedCellValue(sheetName, 30, 4, 415.75),
					new ExpectedCellValue(sheetName, 31, 4, 831.5),
					new ExpectedCellValue(sheetName, 32, 4, 415.75),
					new ExpectedCellValue(sheetName, 33, 4, 831.5),
					new ExpectedCellValue(sheetName, 34, 4, 415.75),
					new ExpectedCellValue(sheetName, 35, 4, 831.5),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, null),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, 415.75),
					new ExpectedCellValue(sheetName, 31, 5, 831.5),
					new ExpectedCellValue(sheetName, 32, 5, 415.75),
					new ExpectedCellValue(sheetName, 33, 5, 831.5),
					new ExpectedCellValue(sheetName, 34, 5, 415.75),
					new ExpectedCellValue(sheetName, 35, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J25:P31"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, null),
					new ExpectedCellValue(sheetName, 28, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 29, 10, "February"),
					new ExpectedCellValue(sheetName, 30, 10, 20100070),
					new ExpectedCellValue(sheetName, 31, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 25, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 26, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 27, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 28, 11, "Tent"),
					new ExpectedCellValue(sheetName, 29, 11, 199d),
					new ExpectedCellValue(sheetName, 30, 11, 199d),
					new ExpectedCellValue(sheetName, 31, 11, 199d),
					new ExpectedCellValue(sheetName, 25, 12, null),
					new ExpectedCellValue(sheetName, 26, 12, null),
					new ExpectedCellValue(sheetName, 27, 12, "Sum of Total"),
					new ExpectedCellValue(sheetName, 28, 12, "Tent"),
					new ExpectedCellValue(sheetName, 29, 12, 1194d),
					new ExpectedCellValue(sheetName, 30, 12, 1194d),
					new ExpectedCellValue(sheetName, 31, 12, 1194d),
					new ExpectedCellValue(sheetName, 25, 13, null),
					new ExpectedCellValue(sheetName, 26, 13, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 13, null),
					new ExpectedCellValue(sheetName, 28, 13, null),
					new ExpectedCellValue(sheetName, 29, 13, 199d),
					new ExpectedCellValue(sheetName, 30, 13, 199d),
					new ExpectedCellValue(sheetName, 31, 13, 199d),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 14, null),
					new ExpectedCellValue(sheetName, 28, 14, null),
					new ExpectedCellValue(sheetName, 29, 14, 1194d),
					new ExpectedCellValue(sheetName, 30, 14, 1194d),
					new ExpectedCellValue(sheetName, 31, 14, 1194d),
					new ExpectedCellValue(sheetName, 25, 15, null),
					new ExpectedCellValue(sheetName, 26, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 15, null),
					new ExpectedCellValue(sheetName, 28, 15, null),
					new ExpectedCellValue(sheetName, 29, 15, 199d),
					new ExpectedCellValue(sheetName, 30, 15, 199d),
					new ExpectedCellValue(sheetName, 31, 15, 199d),
					new ExpectedCellValue(sheetName, 25, 16, null),
					new ExpectedCellValue(sheetName, 26, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, null),
					new ExpectedCellValue(sheetName, 29, 16, 1194d),
					new ExpectedCellValue(sheetName, 30, 16, 1194d),
					new ExpectedCellValue(sheetName, 31, 16, 1194d)

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B40:E45"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 2, null),
					new ExpectedCellValue(sheetName, 42, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 43, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 44, 2, 20100090),
					new ExpectedCellValue(sheetName, 45, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 41, 3, "January"),
					new ExpectedCellValue(sheetName, 42, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 43, 3, 831.5),
					new ExpectedCellValue(sheetName, 44, 3, 831.5),
					new ExpectedCellValue(sheetName, 45, 3, 831.5),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 41, 4, "January Total"),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, 831.5),
					new ExpectedCellValue(sheetName, 44, 4, 831.5),
					new ExpectedCellValue(sheetName, 45, 4, 831.5),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 41, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, 831.5),
					new ExpectedCellValue(sheetName, 44, 5, 831.5),
					new ExpectedCellValue(sheetName, 45, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotContainsWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotContains";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J40:M45"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 10, null),
					new ExpectedCellValue(sheetName, 42, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 43, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 44, 10, 20100017),
					new ExpectedCellValue(sheetName, 45, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 41, 11, "March"),
					new ExpectedCellValue(sheetName, 42, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 43, 11, 831.5),
					new ExpectedCellValue(sheetName, 44, 11, 831.5),
					new ExpectedCellValue(sheetName, 45, 11, 831.5),
					new ExpectedCellValue(sheetName, 40, 12, null),
					new ExpectedCellValue(sheetName, 41, 12, "March Total"),
					new ExpectedCellValue(sheetName, 42, 12, null),
					new ExpectedCellValue(sheetName, 43, 12, 831.5),
					new ExpectedCellValue(sheetName, 44, 12, 831.5),
					new ExpectedCellValue(sheetName, 45, 12, 831.5),
					new ExpectedCellValue(sheetName, 40, 13, null),
					new ExpectedCellValue(sheetName, 41, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 42, 13, null),
					new ExpectedCellValue(sheetName, 43, 13, 831.5),
					new ExpectedCellValue(sheetName, 44, 13, 831.5),
					new ExpectedCellValue(sheetName, 45, 13, 831.5)
				});
			}
		}
		#endregion

		#region CaptionGreaterThan
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 99d),
					new ExpectedCellValue(sheetName, 7, 4, 1293d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 514.75),
					new ExpectedCellValue(sheetName, 7, 6, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:M8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 8, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "January"),
					new ExpectedCellValue(sheetName, 5, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 2078.75),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "March"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, null),
					new ExpectedCellValue(sheetName, 8, 12, 856.49),
					new ExpectedCellValue(sheetName, 3, 13, null),
					new ExpectedCellValue(sheetName, 4, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 13, 856.49),
					new ExpectedCellValue(sheetName, 6, 13, 1663),
					new ExpectedCellValue(sheetName, 7, 13, 415.75),
					new ExpectedCellValue(sheetName, 8, 13, 2935.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithRowAndColumnFilterTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:G19"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 16, 2, "January"),
					new ExpectedCellValue(sheetName, 17, 2, 20100076),
					new ExpectedCellValue(sheetName, 18, 2, 20100090),
					new ExpectedCellValue(sheetName, 19, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 15, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, null),
					new ExpectedCellValue(sheetName, 18, 3, 831.5),
					new ExpectedCellValue(sheetName, 19, 3, 831.5),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "Nashville Total"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, null),
					new ExpectedCellValue(sheetName, 18, 4, 831.5),
					new ExpectedCellValue(sheetName, 19, 4, 831.5),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 15, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 5, 415.75),
					new ExpectedCellValue(sheetName, 17, 5, 415.75),
					new ExpectedCellValue(sheetName, 18, 5, null),
					new ExpectedCellValue(sheetName, 19, 5, 415.75),
					new ExpectedCellValue(sheetName, 13, 6, null),
					new ExpectedCellValue(sheetName, 14, 6, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 15, 6, null),
					new ExpectedCellValue(sheetName, 16, 6, 415.75),
					new ExpectedCellValue(sheetName, 17, 6, 415.75),
					new ExpectedCellValue(sheetName, 18, 6, null),
					new ExpectedCellValue(sheetName, 19, 6, 415.75),
					new ExpectedCellValue(sheetName, 13, 7, null),
					new ExpectedCellValue(sheetName, 14, 7, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 7, null),
					new ExpectedCellValue(sheetName, 16, 7, 1247.25),
					new ExpectedCellValue(sheetName, 17, 7, 415.75),
					new ExpectedCellValue(sheetName, 18, 7, 831.5),
					new ExpectedCellValue(sheetName, 19, 7, 1247.25)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B24:E35"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 2, null),
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 27, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 28, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 29, 2, 20100085),
					new ExpectedCellValue(sheetName, 30, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 2, 20100085),
					new ExpectedCellValue(sheetName, 32, 2, "Sleeping Bag Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "Sleeping Bag Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 25, 3, "San Francisco"),
					new ExpectedCellValue(sheetName, 26, 3, "February"),
					new ExpectedCellValue(sheetName, 27, 3, null),
					new ExpectedCellValue(sheetName, 28, 3, null),
					new ExpectedCellValue(sheetName, 29, 3, 99d),
					new ExpectedCellValue(sheetName, 30, 3, null),
					new ExpectedCellValue(sheetName, 31, 3, 99d),
					new ExpectedCellValue(sheetName, 32, 3, 99d),
					new ExpectedCellValue(sheetName, 33, 3, 99d),
					new ExpectedCellValue(sheetName, 34, 3, 99d),
					new ExpectedCellValue(sheetName, 35, 3, 99d),
					new ExpectedCellValue(sheetName, 24, 4, null),
					new ExpectedCellValue(sheetName, 25, 4, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 26, 4, null),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, null),
					new ExpectedCellValue(sheetName, 29, 4, 99d),
					new ExpectedCellValue(sheetName, 30, 4, null),
					new ExpectedCellValue(sheetName, 31, 4, 99d),
					new ExpectedCellValue(sheetName, 32, 4, 99d),
					new ExpectedCellValue(sheetName, 33, 4, 99d),
					new ExpectedCellValue(sheetName, 34, 4, 99d),
					new ExpectedCellValue(sheetName, 35, 4, 99d),
					new ExpectedCellValue(sheetName, 24, 5, null),
					new ExpectedCellValue(sheetName, 25, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, null),
					new ExpectedCellValue(sheetName, 29, 5, 99d),
					new ExpectedCellValue(sheetName, 30, 5, null),
					new ExpectedCellValue(sheetName, 31, 5, 99d),
					new ExpectedCellValue(sheetName, 32, 5, 99d),
					new ExpectedCellValue(sheetName, 33, 5, 99d),
					new ExpectedCellValue(sheetName, 34, 5, 99d),
					new ExpectedCellValue(sheetName, 35, 5, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J24:P30"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 10, null),
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 28, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 29, 10, 20100085),
					new ExpectedCellValue(sheetName, 30, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 24, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 25, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 27, 11, "February"),
					new ExpectedCellValue(sheetName, 28, 11, 99d),
					new ExpectedCellValue(sheetName, 29, 11, 99d),
					new ExpectedCellValue(sheetName, 30, 11, 99d),
					new ExpectedCellValue(sheetName, 24, 12, null),
					new ExpectedCellValue(sheetName, 25, 12, null),
					new ExpectedCellValue(sheetName, 26, 12, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 27, 12, null),
					new ExpectedCellValue(sheetName, 28, 12, 99d),
					new ExpectedCellValue(sheetName, 29, 12, 99d),
					new ExpectedCellValue(sheetName, 30, 12, 99d),
					new ExpectedCellValue(sheetName, 24, 13, null),
					new ExpectedCellValue(sheetName, 25, 13, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 13, "San Francisco"),
					new ExpectedCellValue(sheetName, 27, 13, "February"),
					new ExpectedCellValue(sheetName, 28, 13, 99d),
					new ExpectedCellValue(sheetName, 29, 13, 99d),
					new ExpectedCellValue(sheetName, 30, 13, 99d),
					new ExpectedCellValue(sheetName, 24, 14, null),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 27, 14, null),
					new ExpectedCellValue(sheetName, 28, 14, 99d),
					new ExpectedCellValue(sheetName, 29, 14, 99d),
					new ExpectedCellValue(sheetName, 30, 14, 99d),
					new ExpectedCellValue(sheetName, 24, 15, null),
					new ExpectedCellValue(sheetName, 25, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 15, null),
					new ExpectedCellValue(sheetName, 27, 15, null),
					new ExpectedCellValue(sheetName, 28, 15, 99d),
					new ExpectedCellValue(sheetName, 29, 15, 99d),
					new ExpectedCellValue(sheetName, 30, 15, 99d),
					new ExpectedCellValue(sheetName, 24, 16, null),
					new ExpectedCellValue(sheetName, 25, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 16, null),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, 99d),
					new ExpectedCellValue(sheetName, 29, 16, 99d),
					new ExpectedCellValue(sheetName, 30, 16, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B41:E46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 41, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 2, null),
					new ExpectedCellValue(sheetName, 43, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 45, 2, 20100085),
					new ExpectedCellValue(sheetName, 46, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 42, 3, "February"),
					new ExpectedCellValue(sheetName, 43, 3, "San Francisco"),
					new ExpectedCellValue(sheetName, 44, 3, 99d),
					new ExpectedCellValue(sheetName, 45, 3, 99d),
					new ExpectedCellValue(sheetName, 46, 3, 99d),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 42, 4, "February Total"),
					new ExpectedCellValue(sheetName, 43, 4, null),
					new ExpectedCellValue(sheetName, 44, 4, 99d),
					new ExpectedCellValue(sheetName, 45, 4, 99d),
					new ExpectedCellValue(sheetName, 46, 4, 99d),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 42, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 43, 5, null),
					new ExpectedCellValue(sheetName, 44, 5, 99d),
					new ExpectedCellValue(sheetName, 45, 5, 99d),
					new ExpectedCellValue(sheetName, 46, 5, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2s()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J41:M46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 41, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 10, null),
					new ExpectedCellValue(sheetName, 43, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 45, 10, 20100083),
					new ExpectedCellValue(sheetName, 46, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 42, 11, "March"),
					new ExpectedCellValue(sheetName, 43, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 44, 11, 24.99),
					new ExpectedCellValue(sheetName, 45, 11, 24.99),
					new ExpectedCellValue(sheetName, 46, 11, 24.99),
					new ExpectedCellValue(sheetName, 41, 12, null),
					new ExpectedCellValue(sheetName, 42, 12, "March Total"),
					new ExpectedCellValue(sheetName, 43, 12, null),
					new ExpectedCellValue(sheetName, 44, 12, 24.99),
					new ExpectedCellValue(sheetName, 45, 12, 24.99),
					new ExpectedCellValue(sheetName, 46, 12, 24.99),
					new ExpectedCellValue(sheetName, 41, 13, null),
					new ExpectedCellValue(sheetName, 42, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 43, 13, null),
					new ExpectedCellValue(sheetName, 44, 13, 24.99),
					new ExpectedCellValue(sheetName, 45, 13, 24.99),
					new ExpectedCellValue(sheetName, 46, 13, 24.99)
				});
			}
		}
		#endregion

		#region CaptionGreaterThanOrEqual Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 99d),
					new ExpectedCellValue(sheetName, 7, 4, 1293d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 514.75),
					new ExpectedCellValue(sheetName, 7, 6, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:M8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 8, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "January"),
					new ExpectedCellValue(sheetName, 5, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 2078.75),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "March"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, null),
					new ExpectedCellValue(sheetName, 8, 12, 856.49),
					new ExpectedCellValue(sheetName, 3, 13, null),
					new ExpectedCellValue(sheetName, 4, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 13, 856.49),
					new ExpectedCellValue(sheetName, 6, 13, 1663d),
					new ExpectedCellValue(sheetName, 7, 13, 415.75),
					new ExpectedCellValue(sheetName, 8, 13, 2935.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithRowAndColumnFiltersTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 16, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 18, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "January"),
					new ExpectedCellValue(sheetName, 15, 3, 20100090),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 18, 3, 831.5),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "January Total"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 831.5),
					new ExpectedCellValue(sheetName, 18, 4, 831.5),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5),
					new ExpectedCellValue(sheetName, 18, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B23:E33"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 23, 2, null),
					new ExpectedCellValue(sheetName, 24, 2, null),
					new ExpectedCellValue(sheetName, 25, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 26, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 27, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 28, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 29, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 2, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 2, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 24, 3, "January"),
					new ExpectedCellValue(sheetName, 25, 3, 20100090),
					new ExpectedCellValue(sheetName, 26, 3, null),
					new ExpectedCellValue(sheetName, 27, 3, null),
					new ExpectedCellValue(sheetName, 28, 3, 415.75),
					new ExpectedCellValue(sheetName, 29, 3, 831.5),
					new ExpectedCellValue(sheetName, 30, 3, 415.75),
					new ExpectedCellValue(sheetName, 31, 3, 831.5),
					new ExpectedCellValue(sheetName, 32, 3, 415.75),
					new ExpectedCellValue(sheetName, 33, 3, 831.5),
					new ExpectedCellValue(sheetName, 23, 4, null),
					new ExpectedCellValue(sheetName, 24, 4, "January Total"),
					new ExpectedCellValue(sheetName, 25, 4, null),
					new ExpectedCellValue(sheetName, 26, 4, null),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, 415.75),
					new ExpectedCellValue(sheetName, 29, 4, 831.5),
					new ExpectedCellValue(sheetName, 30, 4, 415.75),
					new ExpectedCellValue(sheetName, 31, 4, 831.5),
					new ExpectedCellValue(sheetName, 32, 4, 415.75),
					new ExpectedCellValue(sheetName, 33, 4, 831.5),
					new ExpectedCellValue(sheetName, 23, 5, null),
					new ExpectedCellValue(sheetName, 24, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, 415.75),
					new ExpectedCellValue(sheetName, 29, 5, 831.5),
					new ExpectedCellValue(sheetName, 30, 5, 415.75),
					new ExpectedCellValue(sheetName, 31, 5, 831.5),
					new ExpectedCellValue(sheetName, 32, 5, 415.75),
					new ExpectedCellValue(sheetName, 33, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J23:P29"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 23, 10, null),
					new ExpectedCellValue(sheetName, 24, 10, null),
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 27, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 28, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 29, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 23, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 24, 11, "January"),
					new ExpectedCellValue(sheetName, 25, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 11, 20100090),
					new ExpectedCellValue(sheetName, 27, 11, 415.75),
					new ExpectedCellValue(sheetName, 28, 11, 415.75),
					new ExpectedCellValue(sheetName, 29, 11, 415.75),
					new ExpectedCellValue(sheetName, 23, 12, null),
					new ExpectedCellValue(sheetName, 24, 12, null),
					new ExpectedCellValue(sheetName, 25, 12, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 12, 20100090),
					new ExpectedCellValue(sheetName, 27, 12, 831.5),
					new ExpectedCellValue(sheetName, 28, 12, 831.5),
					new ExpectedCellValue(sheetName, 29, 12, 831.5),
					new ExpectedCellValue(sheetName, 23, 13, null),
					new ExpectedCellValue(sheetName, 24, 13, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 13, null),
					new ExpectedCellValue(sheetName, 26, 13, null),
					new ExpectedCellValue(sheetName, 27, 13, 415.75),
					new ExpectedCellValue(sheetName, 28, 13, 415.75),
					new ExpectedCellValue(sheetName, 29, 13, 415.75),
					new ExpectedCellValue(sheetName, 23, 14, null),
					new ExpectedCellValue(sheetName, 24, 14, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, null),
					new ExpectedCellValue(sheetName, 27, 14, 831.5),
					new ExpectedCellValue(sheetName, 28, 14, 831.5),
					new ExpectedCellValue(sheetName, 29, 14, 831.5),
					new ExpectedCellValue(sheetName, 23, 15, null),
					new ExpectedCellValue(sheetName, 24, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 15, null),
					new ExpectedCellValue(sheetName, 26, 15, null),
					new ExpectedCellValue(sheetName, 27, 15, 415.75),
					new ExpectedCellValue(sheetName, 28, 15, 415.75),
					new ExpectedCellValue(sheetName, 29, 15, 415.75),
					new ExpectedCellValue(sheetName, 23, 16, null),
					new ExpectedCellValue(sheetName, 24, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 16, null),
					new ExpectedCellValue(sheetName, 26, 16, null),
					new ExpectedCellValue(sheetName, 27, 16, 831.5),
					new ExpectedCellValue(sheetName, 28, 16, 831.5),
					new ExpectedCellValue(sheetName, 29, 16, 831.5),

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B38:M50"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 38, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 39, 2, null),
					new ExpectedCellValue(sheetName, 40, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 41, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 42, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 43, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 44, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 45, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 46, 2, "Tent"),
					new ExpectedCellValue(sheetName, 47, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 49, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 50, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 38, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 39, 3, "January"),
					new ExpectedCellValue(sheetName, 40, 3, 20100007),
					new ExpectedCellValue(sheetName, 41, 3, 831.5),
					new ExpectedCellValue(sheetName, 42, 3, 831.5),
					new ExpectedCellValue(sheetName, 43, 3, null),
					new ExpectedCellValue(sheetName, 44, 3, null),
					new ExpectedCellValue(sheetName, 45, 3, null),
					new ExpectedCellValue(sheetName, 46, 3, null),
					new ExpectedCellValue(sheetName, 47, 3, null),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 50, 3, 831.5),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 39, 4, null),
					new ExpectedCellValue(sheetName, 40, 4, 20100076),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, null),
					new ExpectedCellValue(sheetName, 44, 4, null),
					new ExpectedCellValue(sheetName, 45, 4, null),
					new ExpectedCellValue(sheetName, 46, 4, null),
					new ExpectedCellValue(sheetName, 47, 4, 415.75),
					new ExpectedCellValue(sheetName, 48, 4, 415.75),
					new ExpectedCellValue(sheetName, 49, 4, null),
					new ExpectedCellValue(sheetName, 50, 4, 415.75),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 39, 5, null),
					new ExpectedCellValue(sheetName, 40, 5, 20100090),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, null),
					new ExpectedCellValue(sheetName, 44, 5, 831.5),
					new ExpectedCellValue(sheetName, 45, 5, 831.5),
					new ExpectedCellValue(sheetName, 46, 5, null),
					new ExpectedCellValue(sheetName, 47, 5, null),
					new ExpectedCellValue(sheetName, 48, 5, null),
					new ExpectedCellValue(sheetName, 49, 5, null),
					new ExpectedCellValue(sheetName, 50, 5, 831.5),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 39, 6, "January Total"),
					new ExpectedCellValue(sheetName, 40, 6, null),
					new ExpectedCellValue(sheetName, 41, 6, 831.5),
					new ExpectedCellValue(sheetName, 42, 6, 831.5),
					new ExpectedCellValue(sheetName, 43, 6, null),
					new ExpectedCellValue(sheetName, 44, 6, 831.5),
					new ExpectedCellValue(sheetName, 45, 6, 831.5),
					new ExpectedCellValue(sheetName, 46, 6, null),
					new ExpectedCellValue(sheetName, 47, 6, 415.75),
					new ExpectedCellValue(sheetName, 48, 6, 415.75),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 6, 2078.75),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 39, 7, "February"),
					new ExpectedCellValue(sheetName, 40, 7, 20100070),
					new ExpectedCellValue(sheetName, 41, 7, null),
					new ExpectedCellValue(sheetName, 42, 7, null),
					new ExpectedCellValue(sheetName, 43, 7, null),
					new ExpectedCellValue(sheetName, 44, 7, 1194d),
					new ExpectedCellValue(sheetName, 45, 7, null),
					new ExpectedCellValue(sheetName, 46, 7, 1194d),
					new ExpectedCellValue(sheetName, 47, 7, null),
					new ExpectedCellValue(sheetName, 48, 7, null),
					new ExpectedCellValue(sheetName, 49, 7, null),
					new ExpectedCellValue(sheetName, 50, 7, 1194d),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 40, 8, 20100085),
					new ExpectedCellValue(sheetName, 41, 8, null),
					new ExpectedCellValue(sheetName, 42, 8, null),
					new ExpectedCellValue(sheetName, 43, 8, null),
					new ExpectedCellValue(sheetName, 44, 8, null),
					new ExpectedCellValue(sheetName, 45, 8, null),
					new ExpectedCellValue(sheetName, 46, 8, null),
					new ExpectedCellValue(sheetName, 47, 8, 99d),
					new ExpectedCellValue(sheetName, 48, 8, null),
					new ExpectedCellValue(sheetName, 49, 8, 99d),
					new ExpectedCellValue(sheetName, 50, 8, 99d),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 39, 9, "February Total"),
					new ExpectedCellValue(sheetName, 40, 9, null),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 42, 9, null),
					new ExpectedCellValue(sheetName, 43, 9, null),
					new ExpectedCellValue(sheetName, 44, 9, 1194d),
					new ExpectedCellValue(sheetName, 45, 9, null),
					new ExpectedCellValue(sheetName, 46, 9, 1194d),
					new ExpectedCellValue(sheetName, 47, 9, 99d),
					new ExpectedCellValue(sheetName, 48, 9, null),
					new ExpectedCellValue(sheetName, 49, 9, 99d),
					new ExpectedCellValue(sheetName, 50, 9, 1293d),
					new ExpectedCellValue(sheetName, 38, 10, null),
					new ExpectedCellValue(sheetName, 39, 10, "March"),
					new ExpectedCellValue(sheetName, 40, 10, 20100017),
					new ExpectedCellValue(sheetName, 41, 10, null),
					new ExpectedCellValue(sheetName, 42, 10, null),
					new ExpectedCellValue(sheetName, 43, 10, null),
					new ExpectedCellValue(sheetName, 44, 10, 831.5),
					new ExpectedCellValue(sheetName, 45, 10, 831.5),
					new ExpectedCellValue(sheetName, 46, 10, null),
					new ExpectedCellValue(sheetName, 47, 10, null),
					new ExpectedCellValue(sheetName, 48, 10, null),
					new ExpectedCellValue(sheetName, 49, 10, null),
					new ExpectedCellValue(sheetName, 50, 10, 831.5),
					new ExpectedCellValue(sheetName, 38, 11, null),
					new ExpectedCellValue(sheetName, 39, 11, null),
					new ExpectedCellValue(sheetName, 40, 11, 20100083),
					new ExpectedCellValue(sheetName, 41, 11, 24.99),
					new ExpectedCellValue(sheetName, 42, 11, null),
					new ExpectedCellValue(sheetName, 43, 11, 24.99),
					new ExpectedCellValue(sheetName, 44, 11, null),
					new ExpectedCellValue(sheetName, 45, 11, null),
					new ExpectedCellValue(sheetName, 46, 11, null),
					new ExpectedCellValue(sheetName, 47, 11, null),
					new ExpectedCellValue(sheetName, 48, 11, null),
					new ExpectedCellValue(sheetName, 49, 11, null),
					new ExpectedCellValue(sheetName, 50, 11, 24.99),
					new ExpectedCellValue(sheetName, 38, 12, null),
					new ExpectedCellValue(sheetName, 39, 12, "March Total"),
					new ExpectedCellValue(sheetName, 40, 12, null),
					new ExpectedCellValue(sheetName, 41, 12, 24.99),
					new ExpectedCellValue(sheetName, 42, 12, null),
					new ExpectedCellValue(sheetName, 43, 12, 24.99),
					new ExpectedCellValue(sheetName, 44, 12, 831.5),
					new ExpectedCellValue(sheetName, 45, 12, 831.5),
					new ExpectedCellValue(sheetName, 46, 12, null),
					new ExpectedCellValue(sheetName, 47, 12, null),
					new ExpectedCellValue(sheetName, 48, 12, null),
					new ExpectedCellValue(sheetName, 49, 12, null),
					new ExpectedCellValue(sheetName, 50, 12, 856.49),
					new ExpectedCellValue(sheetName, 38, 13, null),
					new ExpectedCellValue(sheetName, 39, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 13, null),
					new ExpectedCellValue(sheetName, 41, 13, 856.49),
					new ExpectedCellValue(sheetName, 42, 13, 831.5),
					new ExpectedCellValue(sheetName, 43, 13, 24.99),
					new ExpectedCellValue(sheetName, 44, 13, 2857d),
					new ExpectedCellValue(sheetName, 45, 13, 1663d),
					new ExpectedCellValue(sheetName, 46, 13, 1194d),
					new ExpectedCellValue(sheetName, 47, 13, 514.75),
					new ExpectedCellValue(sheetName, 48, 13, 415.75),
					new ExpectedCellValue(sheetName, 49, 13, 99d),
					new ExpectedCellValue(sheetName, 50, 13, 4228.24),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionGreaterThanOrEqualWithRegularExpressionRowAndColumnFiltersEnabledOneRowFieldOneColumnField2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionGreaterThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B55:F60"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 55, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 56, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 57, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 58, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 59, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 60, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 55, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 56, 3, "January"),
					new ExpectedCellValue(sheetName, 57, 3, 831.5),
					new ExpectedCellValue(sheetName, 58, 3, 831.5),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 60, 3, 2078.75),
					new ExpectedCellValue(sheetName, 55, 4, null),
					new ExpectedCellValue(sheetName, 56, 4, "February"),
					new ExpectedCellValue(sheetName, 57, 4, null),
					new ExpectedCellValue(sheetName, 58, 4, 1194d),
					new ExpectedCellValue(sheetName, 59, 4, 99d),
					new ExpectedCellValue(sheetName, 60, 4, 1293d),
					new ExpectedCellValue(sheetName, 55, 5, null),
					new ExpectedCellValue(sheetName, 56, 5, "March"),
					new ExpectedCellValue(sheetName, 57, 5, 24.99),
					new ExpectedCellValue(sheetName, 58, 5, 831.5),
					new ExpectedCellValue(sheetName, 59, 5, null),
					new ExpectedCellValue(sheetName, 60, 5, 856.49),
					new ExpectedCellValue(sheetName, 55, 6, null),
					new ExpectedCellValue(sheetName, 56, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 57, 6, 856.49),
					new ExpectedCellValue(sheetName, 58, 6, 2857d),
					new ExpectedCellValue(sheetName, 59, 6, 514.75),
					new ExpectedCellValue(sheetName, 60, 6, 4228.24),
				});
			}
		}
		#endregion

		#region CaptionLessThan Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:E6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "March"),
					new ExpectedCellValue(sheetName, 5, 4, 24.99),
					new ExpectedCellValue(sheetName, 6, 4, 24.99),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 856.49),
					new ExpectedCellValue(sheetName, 6, 5, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "February"),
					new ExpectedCellValue(sheetName, 5, 11, 1194d),
					new ExpectedCellValue(sheetName, 6, 11, 99d),
					new ExpectedCellValue(sheetName, 7, 11, 1293d),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 1194d),
					new ExpectedCellValue(sheetName, 6, 12, 99d),
					new ExpectedCellValue(sheetName, 7, 12, 1293d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithRowAndColumnFiltersTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 16, 2, "January"),
					new ExpectedCellValue(sheetName, 17, 2, 20100007),
					new ExpectedCellValue(sheetName, 18, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 18, 3, 831.5),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 831.5),
					new ExpectedCellValue(sheetName, 18, 4, 831.5),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5),
					new ExpectedCellValue(sheetName, 18, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B24:E35"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 2, null),
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 27, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 28, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 29, 2, 20100007),
					new ExpectedCellValue(sheetName, 30, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 2, 20100007),
					new ExpectedCellValue(sheetName, 32, 2, "Car Rack Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 25, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 26, 3, "January"),
					new ExpectedCellValue(sheetName, 27, 3, null),
					new ExpectedCellValue(sheetName, 28, 3, null),
					new ExpectedCellValue(sheetName, 29, 3, 415.75),
					new ExpectedCellValue(sheetName, 30, 3, null),
					new ExpectedCellValue(sheetName, 31, 3, 831.5),
					new ExpectedCellValue(sheetName, 32, 3, 415.75),
					new ExpectedCellValue(sheetName, 33, 3, 831.5),
					new ExpectedCellValue(sheetName, 34, 3, 415.75),
					new ExpectedCellValue(sheetName, 35, 3, 831.5),
					new ExpectedCellValue(sheetName, 24, 4, null),
					new ExpectedCellValue(sheetName, 25, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 26, 4, null),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, null),
					new ExpectedCellValue(sheetName, 29, 4, 415.75),
					new ExpectedCellValue(sheetName, 30, 4, null),
					new ExpectedCellValue(sheetName, 31, 4, 831.5),
					new ExpectedCellValue(sheetName, 32, 4, 415.75),
					new ExpectedCellValue(sheetName, 33, 4, 831.5),
					new ExpectedCellValue(sheetName, 34, 4, 415.75),
					new ExpectedCellValue(sheetName, 35, 4, 831.5),
					new ExpectedCellValue(sheetName, 24, 5, null),
					new ExpectedCellValue(sheetName, 25, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, null),
					new ExpectedCellValue(sheetName, 29, 5, 415.75),
					new ExpectedCellValue(sheetName, 30, 5, null),
					new ExpectedCellValue(sheetName, 31, 5, 831.5),
					new ExpectedCellValue(sheetName, 32, 5, 415.75),
					new ExpectedCellValue(sheetName, 33, 5, 831.5),
					new ExpectedCellValue(sheetName, 34, 5, 415.75),
					new ExpectedCellValue(sheetName, 35, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J24:P30"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 10, null),
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 28, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 29, 10, 20100007),
					new ExpectedCellValue(sheetName, 30, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 24, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 25, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 27, 11, "January"),
					new ExpectedCellValue(sheetName, 28, 11, 415.75),
					new ExpectedCellValue(sheetName, 29, 11, 415.75),
					new ExpectedCellValue(sheetName, 30, 11, 415.75),
					new ExpectedCellValue(sheetName, 24, 12, null),
					new ExpectedCellValue(sheetName, 25, 12, null),
					new ExpectedCellValue(sheetName, 26, 12, "Chicago Total"),
					new ExpectedCellValue(sheetName, 27, 12, null),
					new ExpectedCellValue(sheetName, 28, 12, 415.75),
					new ExpectedCellValue(sheetName, 29, 12, 415.75),
					new ExpectedCellValue(sheetName, 30, 12, 415.75),
					new ExpectedCellValue(sheetName, 24, 13, null),
					new ExpectedCellValue(sheetName, 25, 13, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 13, "Chicago"),
					new ExpectedCellValue(sheetName, 27, 13, "January"),
					new ExpectedCellValue(sheetName, 28, 13, 831.5),
					new ExpectedCellValue(sheetName, 29, 13, 831.5),
					new ExpectedCellValue(sheetName, 30, 13, 831.5),
					new ExpectedCellValue(sheetName, 24, 14, null),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, "Chicago Total"),
					new ExpectedCellValue(sheetName, 27, 14, null),
					new ExpectedCellValue(sheetName, 28, 14, 831.5),
					new ExpectedCellValue(sheetName, 29, 14, 831.5),
					new ExpectedCellValue(sheetName, 30, 14, 831.5),
					new ExpectedCellValue(sheetName, 24, 15, null),
					new ExpectedCellValue(sheetName, 25, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 15, null),
					new ExpectedCellValue(sheetName, 27, 15, null),
					new ExpectedCellValue(sheetName, 28, 15, 415.75),
					new ExpectedCellValue(sheetName, 29, 15, 415.75),
					new ExpectedCellValue(sheetName, 30, 15, 415.75),
					new ExpectedCellValue(sheetName, 24, 16, null),
					new ExpectedCellValue(sheetName, 25, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 16, null),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, 831.5),
					new ExpectedCellValue(sheetName, 29, 16, 831.5),
					new ExpectedCellValue(sheetName, 30, 16, 831.5),

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B41:C45"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 41, 2, null),
					new ExpectedCellValue(sheetName, 42, 2, null),
					new ExpectedCellValue(sheetName, 43, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 45, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 42, 3, "Grand Total"),
					new ExpectedCellValue(sheetName, 43, 3, null),
					new ExpectedCellValue(sheetName, 44, 3, null),
					new ExpectedCellValue(sheetName, 45, 3, null)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThan";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J41:K45"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 41, 10, null),
					new ExpectedCellValue(sheetName, 42, 10, null),
					new ExpectedCellValue(sheetName, 43, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 45, 10, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 42, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 43, 11, null),
					new ExpectedCellValue(sheetName, 44, 11, null),
					new ExpectedCellValue(sheetName, 45, 11, null)
				});
			}
		}
		#endregion

		#region CaptionLessThanOrEqual Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:E6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "March"),
					new ExpectedCellValue(sheetName, 5, 4, 24.99),
					new ExpectedCellValue(sheetName, 6, 4, 24.99),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 856.49),
					new ExpectedCellValue(sheetName, 6, 5, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "February"),
					new ExpectedCellValue(sheetName, 5, 11, 1194d),
					new ExpectedCellValue(sheetName, 6, 11, 99d),
					new ExpectedCellValue(sheetName, 7, 11, 1293d),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 1194d),
					new ExpectedCellValue(sheetName, 6, 12, 99d),
					new ExpectedCellValue(sheetName, 7, 12, 1293d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithRowAndColumnFiltersTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B13:E18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 16, 2, "January"),
					new ExpectedCellValue(sheetName, 17, 2, 20100007),
					new ExpectedCellValue(sheetName, 18, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 14, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 18, 3, 831.5),
					new ExpectedCellValue(sheetName, 13, 4, null),
					new ExpectedCellValue(sheetName, 14, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 15, 4, null),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 831.5),
					new ExpectedCellValue(sheetName, 18, 4, 831.5),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 14, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5),
					new ExpectedCellValue(sheetName, 18, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B24:E35"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 2, null),
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 27, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 28, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 29, 2, 20100007),
					new ExpectedCellValue(sheetName, 30, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 2, 20100007),
					new ExpectedCellValue(sheetName, 32, 2, "Car Rack Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 25, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 26, 3, "January"),
					new ExpectedCellValue(sheetName, 27, 3, null),
					new ExpectedCellValue(sheetName, 28, 3, null),
					new ExpectedCellValue(sheetName, 29, 3, 415.75),
					new ExpectedCellValue(sheetName, 30, 3, null),
					new ExpectedCellValue(sheetName, 31, 3, 831.5),
					new ExpectedCellValue(sheetName, 32, 3, 415.75),
					new ExpectedCellValue(sheetName, 33, 3, 831.5),
					new ExpectedCellValue(sheetName, 34, 3, 415.75),
					new ExpectedCellValue(sheetName, 35, 3, 831.5),
					new ExpectedCellValue(sheetName, 24, 4, null),
					new ExpectedCellValue(sheetName, 25, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 26, 4, null),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, null),
					new ExpectedCellValue(sheetName, 29, 4, 415.75),
					new ExpectedCellValue(sheetName, 30, 4, null),
					new ExpectedCellValue(sheetName, 31, 4, 831.5),
					new ExpectedCellValue(sheetName, 32, 4, 415.75),
					new ExpectedCellValue(sheetName, 33, 4, 831.5),
					new ExpectedCellValue(sheetName, 34, 4, 415.75),
					new ExpectedCellValue(sheetName, 35, 4, 831.5),
					new ExpectedCellValue(sheetName, 24, 5, null),
					new ExpectedCellValue(sheetName, 25, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, null),
					new ExpectedCellValue(sheetName, 29, 5, 415.75),
					new ExpectedCellValue(sheetName, 30, 5, null),
					new ExpectedCellValue(sheetName, 31, 5, 831.5),
					new ExpectedCellValue(sheetName, 32, 5, 415.75),
					new ExpectedCellValue(sheetName, 33, 5, 831.5),
					new ExpectedCellValue(sheetName, 34, 5, 415.75),
					new ExpectedCellValue(sheetName, 35, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J24:R32"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 10, null),
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 28, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 29, 10, 20100007),
					new ExpectedCellValue(sheetName, 30, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 31, 10, 20100083),
					new ExpectedCellValue(sheetName, 32, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 24, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 25, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 27, 11, "January"),
					new ExpectedCellValue(sheetName, 28, 11, 415.75),
					new ExpectedCellValue(sheetName, 29, 11, 415.75),
					new ExpectedCellValue(sheetName, 30, 11, null),
					new ExpectedCellValue(sheetName, 31, 11, null),
					new ExpectedCellValue(sheetName, 32, 11, 415.75),
					new ExpectedCellValue(sheetName, 24, 12, null),
					new ExpectedCellValue(sheetName, 25, 12, null),
					new ExpectedCellValue(sheetName, 26, 12, null),
					new ExpectedCellValue(sheetName, 27, 12, "March"),
					new ExpectedCellValue(sheetName, 28, 12, null),
					new ExpectedCellValue(sheetName, 29, 12, null),
					new ExpectedCellValue(sheetName, 30, 12, 24.99),
					new ExpectedCellValue(sheetName, 31, 12, 24.99),
					new ExpectedCellValue(sheetName, 32, 12, 24.99),
					new ExpectedCellValue(sheetName, 24, 13, null),
					new ExpectedCellValue(sheetName, 25, 13, null),
					new ExpectedCellValue(sheetName, 26, 13, "Chicago Total"),
					new ExpectedCellValue(sheetName, 27, 13, null),
					new ExpectedCellValue(sheetName, 28, 13, 415.75),
					new ExpectedCellValue(sheetName, 29, 13, 415.75),
					new ExpectedCellValue(sheetName, 30, 13, 24.99),
					new ExpectedCellValue(sheetName, 31, 13, 24.99),
					new ExpectedCellValue(sheetName, 32, 13, 440.74),
					new ExpectedCellValue(sheetName, 24, 14, null),
					new ExpectedCellValue(sheetName, 25, 14, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 14, "Chicago"),
					new ExpectedCellValue(sheetName, 27, 14, "January"),
					new ExpectedCellValue(sheetName, 28, 14, 831.5),
					new ExpectedCellValue(sheetName, 29, 14, 831.5),
					new ExpectedCellValue(sheetName, 30, 14, null),
					new ExpectedCellValue(sheetName, 31, 14, null),
					new ExpectedCellValue(sheetName, 32, 14, 831.5),
					new ExpectedCellValue(sheetName, 24, 15, null),
					new ExpectedCellValue(sheetName, 25, 15, null),
					new ExpectedCellValue(sheetName, 26, 15, null),
					new ExpectedCellValue(sheetName, 27, 15, "March"),
					new ExpectedCellValue(sheetName, 28, 15, null),
					new ExpectedCellValue(sheetName, 29, 15, null),
					new ExpectedCellValue(sheetName, 30, 15, 24.99),
					new ExpectedCellValue(sheetName, 31, 15, 24.99),
					new ExpectedCellValue(sheetName, 32, 15, 24.99),
					new ExpectedCellValue(sheetName, 24, 16, null),
					new ExpectedCellValue(sheetName, 25, 16, null),
					new ExpectedCellValue(sheetName, 26, 16, "Chicago Total"),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, 831.5),
					new ExpectedCellValue(sheetName, 29, 16, 831.5),
					new ExpectedCellValue(sheetName, 30, 16, 24.99),
					new ExpectedCellValue(sheetName, 31, 16, 24.99),
					new ExpectedCellValue(sheetName, 32, 16, 856.49),
					new ExpectedCellValue(sheetName, 24, 17, null),
					new ExpectedCellValue(sheetName, 25, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 17, null),
					new ExpectedCellValue(sheetName, 27, 17, null),
					new ExpectedCellValue(sheetName, 28, 17, 415.75),
					new ExpectedCellValue(sheetName, 29, 17, 415.75),
					new ExpectedCellValue(sheetName, 30, 17, 24.99),
					new ExpectedCellValue(sheetName, 31, 17, 24.99),
					new ExpectedCellValue(sheetName, 32, 17, 440.74),
					new ExpectedCellValue(sheetName, 24, 18, null),
					new ExpectedCellValue(sheetName, 25, 18, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 18, null),
					new ExpectedCellValue(sheetName, 27, 18, null),
					new ExpectedCellValue(sheetName, 28, 18, 831.5),
					new ExpectedCellValue(sheetName, 29, 18, 831.5),
					new ExpectedCellValue(sheetName, 30, 18, 24.99),
					new ExpectedCellValue(sheetName, 31, 18, 24.99),
					new ExpectedCellValue(sheetName, 32, 18, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B41:E46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 41, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 2, null),
					new ExpectedCellValue(sheetName, 43, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 45, 2, 20100076),
					new ExpectedCellValue(sheetName, 46, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 42, 3, "San Francisco"),
					new ExpectedCellValue(sheetName, 43, 3, "January"),
					new ExpectedCellValue(sheetName, 44, 3, 415.75),
					new ExpectedCellValue(sheetName, 45, 3, 415.75 ),
					new ExpectedCellValue(sheetName, 46, 3, 415.75),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 42, 4, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 43, 4, null),
					new ExpectedCellValue(sheetName, 44, 4, 415.75),
					new ExpectedCellValue(sheetName, 45, 4, 415.75 ),
					new ExpectedCellValue(sheetName, 46, 4, 415.75),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 42, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 43, 5, null),
					new ExpectedCellValue(sheetName, 44, 5, 415.75),
					new ExpectedCellValue(sheetName, 45, 5, 415.75 ),
					new ExpectedCellValue(sheetName, 46, 5, 415.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionLessThanOrEqualWithRegularExpressionRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields2()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionLessThanOrEqual";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J41:M46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 41, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 10, null),
					new ExpectedCellValue(sheetName, 43, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 44, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 45, 10, 20100017),
					new ExpectedCellValue(sheetName, 46, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 42, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 43, 11, "March"),
					new ExpectedCellValue(sheetName, 44, 11, 831.5),
					new ExpectedCellValue(sheetName, 45, 11, 831.5),
					new ExpectedCellValue(sheetName, 46, 11, 831.5),
					new ExpectedCellValue(sheetName, 41, 12, null),
					new ExpectedCellValue(sheetName, 42, 12, "Nashville Total"),
					new ExpectedCellValue(sheetName, 43, 12, null),
					new ExpectedCellValue(sheetName, 44, 12, 831.5),
					new ExpectedCellValue(sheetName, 45, 12, 831.5),
					new ExpectedCellValue(sheetName, 46, 12, 831.5),
					new ExpectedCellValue(sheetName, 41, 13, null),
					new ExpectedCellValue(sheetName, 42, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 43, 13, null),
					new ExpectedCellValue(sheetName, 44, 13, 831.5),
					new ExpectedCellValue(sheetName, 45, 13, 831.5),
					new ExpectedCellValue(sheetName, 46, 13, 831.5)
				});
			}
		}
		#endregion

		#region CaptionBetween
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBetweenWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:F7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "February"),
					new ExpectedCellValue(sheetName, 5, 4, 1194d),
					new ExpectedCellValue(sheetName, 6, 4, 99d),
					new ExpectedCellValue(sheetName, 7, 4, 1293d),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "March"),
					new ExpectedCellValue(sheetName, 5, 5, 831.5),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 6, 2857d),
					new ExpectedCellValue(sheetName, 6, 6, 514.75),
					new ExpectedCellValue(sheetName, 7, 6, 3371.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBetweenWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:M8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 8, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "January"),
					new ExpectedCellValue(sheetName, 5, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 831.5),
					new ExpectedCellValue(sheetName, 7, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 2078.75),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "March"),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 831.5),
					new ExpectedCellValue(sheetName, 7, 12, null),
					new ExpectedCellValue(sheetName, 8, 12, 856.49),
					new ExpectedCellValue(sheetName, 3, 13, null),
					new ExpectedCellValue(sheetName, 4, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 13, 856.49),
					new ExpectedCellValue(sheetName, 6, 13, 1663d),
					new ExpectedCellValue(sheetName, 7, 13, 415.75),
					new ExpectedCellValue(sheetName, 8, 13, 2935.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBetweenWithRowAndColumnFiltersTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:E19"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 12, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 2, null),
					new ExpectedCellValue(sheetName, 14, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 15, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 16, 2, 20100090),
					new ExpectedCellValue(sheetName, 17, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 18, 2, 20100076),
					new ExpectedCellValue(sheetName, 19, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 13, 3, "January"),
					new ExpectedCellValue(sheetName, 14, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 3, 831.5),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 415.75),
					new ExpectedCellValue(sheetName, 18, 3, 415.75),
					new ExpectedCellValue(sheetName, 19, 3, 1247.25),
					new ExpectedCellValue(sheetName, 12, 4, null),
					new ExpectedCellValue(sheetName, 13, 4, "January Total"),
					new ExpectedCellValue(sheetName, 14, 4, null),
					new ExpectedCellValue(sheetName, 15, 4, 831.5),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 415.75),
					new ExpectedCellValue(sheetName, 18, 4, 415.75),
					new ExpectedCellValue(sheetName, 19, 4, 1247.25),
					new ExpectedCellValue(sheetName, 12, 5, null),
					new ExpectedCellValue(sheetName, 13, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 14, 5, null),
					new ExpectedCellValue(sheetName, 15, 5, 831.5),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 415.75),
					new ExpectedCellValue(sheetName, 18, 5, 415.75),
					new ExpectedCellValue(sheetName, 19, 5, 1247.25)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBetweenWithMultipleRowDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B25:E43"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, null),
					new ExpectedCellValue(sheetName, 27, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 28, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 29, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 2, 20100090),
					new ExpectedCellValue(sheetName, 31, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 2, 20100090),
					new ExpectedCellValue(sheetName, 33, 2, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 34, 2, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 35, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 36, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 37, 2, 20100076),
					new ExpectedCellValue(sheetName, 38, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 39, 2, 20100076),
					new ExpectedCellValue(sheetName, 40, 2, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 41, 2, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 43, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 26, 3, "January"),
					new ExpectedCellValue(sheetName, 27, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 28, 3, null),
					new ExpectedCellValue(sheetName, 29, 3, null),
					new ExpectedCellValue(sheetName, 30, 3, 415.75),
					new ExpectedCellValue(sheetName, 31, 3, null),
					new ExpectedCellValue(sheetName, 32, 3, 831.5),
					new ExpectedCellValue(sheetName, 33, 3, 415.75),
					new ExpectedCellValue(sheetName, 34, 3, 831.5),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 37, 3, 415.75),
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 39, 3, 415.75),
					new ExpectedCellValue(sheetName, 40, 3, 415.75),
					new ExpectedCellValue(sheetName, 41, 3, 415.75),
					new ExpectedCellValue(sheetName, 42, 3, 831.5),
					new ExpectedCellValue(sheetName, 43, 3, 1247.25),
					new ExpectedCellValue(sheetName, 25, 4, null),
					new ExpectedCellValue(sheetName, 26, 4, "January Total"),
					new ExpectedCellValue(sheetName, 27, 4, null),
					new ExpectedCellValue(sheetName, 28, 4, null),
					new ExpectedCellValue(sheetName, 29, 4, null),
					new ExpectedCellValue(sheetName, 30, 4, 415.75),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 32, 4, 831.5),
					new ExpectedCellValue(sheetName, 33, 4, 415.75),
					new ExpectedCellValue(sheetName, 34, 4, 831.5),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 37, 4, 415.75),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 39, 4, 415.75),
					new ExpectedCellValue(sheetName, 40, 4, 415.75),
					new ExpectedCellValue(sheetName, 41, 4, 415.75),
					new ExpectedCellValue(sheetName, 42, 4, 831.5),
					new ExpectedCellValue(sheetName, 43, 4, 1247.25),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 28, 5, null),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, 415.75),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 32, 5, 831.5),
					new ExpectedCellValue(sheetName, 33, 5, 415.75),
					new ExpectedCellValue(sheetName, 34, 5, 831.5),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 37, 5, 415.75),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 39, 5, 415.75),
					new ExpectedCellValue(sheetName, 40, 5, 415.75),
					new ExpectedCellValue(sheetName, 41, 5, 415.75),
					new ExpectedCellValue(sheetName, 42, 5, 831.5),
					new ExpectedCellValue(sheetName, 43, 5, 1247.25)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionBetweenWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J25:P33"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, null),
					new ExpectedCellValue(sheetName, 28, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 29, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 30, 10, 20100090),
					new ExpectedCellValue(sheetName, 31, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 10, 20100076),
					new ExpectedCellValue(sheetName, 33, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 25, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 26, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 11, "January"),
					new ExpectedCellValue(sheetName, 28, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 29, 11, 415.75),
					new ExpectedCellValue(sheetName, 30, 11, 415.75),
					new ExpectedCellValue(sheetName, 31, 11, 415.75),
					new ExpectedCellValue(sheetName, 32, 11, 415.75),
					new ExpectedCellValue(sheetName, 33, 11, 831.5),
					new ExpectedCellValue(sheetName, 25, 12, null),
					new ExpectedCellValue(sheetName, 26, 12, null),
					new ExpectedCellValue(sheetName, 27, 12, "January Total"),
					new ExpectedCellValue(sheetName, 28, 12, null),
					new ExpectedCellValue(sheetName, 29, 12, 415.75),
					new ExpectedCellValue(sheetName, 30, 12, 415.75),
					new ExpectedCellValue(sheetName, 31, 12, 415.75),
					new ExpectedCellValue(sheetName, 32, 12, 415.75),
					new ExpectedCellValue(sheetName, 33, 12, 831.5),
					new ExpectedCellValue(sheetName, 25, 13, null),
					new ExpectedCellValue(sheetName, 26, 13, "Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 13, "January"),
					new ExpectedCellValue(sheetName, 28, 13, "Car Rack"),
					new ExpectedCellValue(sheetName, 29, 13, 831.5),
					new ExpectedCellValue(sheetName, 30, 13, 831.5),
					new ExpectedCellValue(sheetName, 31, 13, 415.75),
					new ExpectedCellValue(sheetName, 32, 13, 415.75),
					new ExpectedCellValue(sheetName, 33, 13, 1247.25),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, null),
					new ExpectedCellValue(sheetName, 27, 14, "January Total"),
					new ExpectedCellValue(sheetName, 28, 14, null),
					new ExpectedCellValue(sheetName, 29, 14, 831.5),
					new ExpectedCellValue(sheetName, 30, 14, 831.5),
					new ExpectedCellValue(sheetName, 31, 14, 415.75),
					new ExpectedCellValue(sheetName, 32, 14, 415.75),
					new ExpectedCellValue(sheetName, 33, 14, 1247.25),
					new ExpectedCellValue(sheetName, 25, 15, null),
					new ExpectedCellValue(sheetName, 26, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 15, null),
					new ExpectedCellValue(sheetName, 28, 15, null),
					new ExpectedCellValue(sheetName, 29, 15, 415.75),
					new ExpectedCellValue(sheetName, 30, 15, 415.75),
					new ExpectedCellValue(sheetName, 31, 15, 415.75),
					new ExpectedCellValue(sheetName, 32, 15, 415.75),
					new ExpectedCellValue(sheetName, 33, 15, 831.5),
					new ExpectedCellValue(sheetName, 25, 16, null),
					new ExpectedCellValue(sheetName, 26, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 16, null),
					new ExpectedCellValue(sheetName, 28, 16, null),
					new ExpectedCellValue(sheetName, 29, 16, 831.5),
					new ExpectedCellValue(sheetName, 30, 16, 831.5),
					new ExpectedCellValue(sheetName, 31, 16, 415.75),
					new ExpectedCellValue(sheetName, 32, 16, 415.75),
					new ExpectedCellValue(sheetName, 33, 16, 1247.25)
				});
			}
		}
		#endregion

		#region CaptionNotBetween
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBetweenWithRowFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B3:E6"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "January"),
					new ExpectedCellValue(sheetName, 5, 3, 831.5),
					new ExpectedCellValue(sheetName, 6, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, "March"),
					new ExpectedCellValue(sheetName, 5, 4, 24.99),
					new ExpectedCellValue(sheetName, 6, 4, 24.99),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 4, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 856.49),
					new ExpectedCellValue(sheetName, 6, 5, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBetweenWithColumnFilterOnlyOneRowFieldOneColumnField()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J3:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 10, "Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 11, "February"),
					new ExpectedCellValue(sheetName, 5, 11, 1194d),
					new ExpectedCellValue(sheetName, 6, 11, 99d),
					new ExpectedCellValue(sheetName, 7, 11, 1293d),
					new ExpectedCellValue(sheetName, 3, 12, null),
					new ExpectedCellValue(sheetName, 4, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 1194d),
					new ExpectedCellValue(sheetName, 6, 12, 99d),
					new ExpectedCellValue(sheetName, 7, 12, 1293d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBetweenWithRowAndColumnFiltersTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:E17"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 12, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 2, null),
					new ExpectedCellValue(sheetName, 14, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 15, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 16, 2, 20100017),
					new ExpectedCellValue(sheetName, 17, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 13, 3, "March"),
					new ExpectedCellValue(sheetName, 14, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 15, 3, 831.5),
					new ExpectedCellValue(sheetName, 16, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 12, 4, null),
					new ExpectedCellValue(sheetName, 13, 4, "March Total"),
					new ExpectedCellValue(sheetName, 14, 4, null),
					new ExpectedCellValue(sheetName, 15, 4, 831.5),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 17, 4, 831.5),
					new ExpectedCellValue(sheetName, 12, 5, null),
					new ExpectedCellValue(sheetName, 13, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 14, 5, null),
					new ExpectedCellValue(sheetName, 15, 5, 831.5),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBetweenWithMultipleRowDataFieldsRowAndColumnFilterEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B23:E33"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 23, 2, null),
					new ExpectedCellValue(sheetName, 24, 2, null),
					new ExpectedCellValue(sheetName, 25, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 26, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 27, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 28, 2, 20100085),
					new ExpectedCellValue(sheetName, 29, 2, "Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 31, 2, 20100085),
					new ExpectedCellValue(sheetName, 32, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 2, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 24, 3, "February"),
					new ExpectedCellValue(sheetName, 25, 3, "San Francisco"),
					new ExpectedCellValue(sheetName, 26, 3, null),
					new ExpectedCellValue(sheetName, 27, 3, 99d),
					new ExpectedCellValue(sheetName, 28, 3, 99d),
					new ExpectedCellValue(sheetName, 29, 3, null),
					new ExpectedCellValue(sheetName, 30, 3, 99d),
					new ExpectedCellValue(sheetName, 31, 3, 99d),
					new ExpectedCellValue(sheetName, 32, 3, 99d),
					new ExpectedCellValue(sheetName, 33, 3, 99d),
					new ExpectedCellValue(sheetName, 23, 4, null),
					new ExpectedCellValue(sheetName, 24, 4, "February Total"),
					new ExpectedCellValue(sheetName, 25, 4, null),
					new ExpectedCellValue(sheetName, 26, 4, null),
					new ExpectedCellValue(sheetName, 27, 4, 99d),
					new ExpectedCellValue(sheetName, 28, 4, 99d),
					new ExpectedCellValue(sheetName, 29, 4, null),
					new ExpectedCellValue(sheetName, 30, 4, 99d),
					new ExpectedCellValue(sheetName, 31, 4, 99d),
					new ExpectedCellValue(sheetName, 32, 4, 99d),
					new ExpectedCellValue(sheetName, 33, 4, 99d),
					new ExpectedCellValue(sheetName, 23, 5, null),
					new ExpectedCellValue(sheetName, 24, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, 99d),
					new ExpectedCellValue(sheetName, 28, 5, 99d),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, 99d),
					new ExpectedCellValue(sheetName, 31, 5, 99d),
					new ExpectedCellValue(sheetName, 32, 5, 99d),
					new ExpectedCellValue(sheetName, 33, 5, 99d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableLabelFilters.xlsx")]
		public void PivotTableLabelFiltersCaptionNotBetweenWithMultipleColumnDataFieldsRowAndColumnFiltersEnabledTwoRowFieldsTwoColumnFields()
		{
			var file = new FileInfo("PivotTableLabelFilters.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "CaptionNotBetween";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J23:P29"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 23, 10, null),
					new ExpectedCellValue(sheetName, 24, 10, null),
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 27, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 28, 10, 20100085),
					new ExpectedCellValue(sheetName, 29, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 23, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 24, 11, "February"),
					new ExpectedCellValue(sheetName, 25, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 27, 11, 99d),
					new ExpectedCellValue(sheetName, 28, 11, 99d),
					new ExpectedCellValue(sheetName, 29, 11, 99d),
					new ExpectedCellValue(sheetName, 23, 12, null),
					new ExpectedCellValue(sheetName, 24, 12, null),
					new ExpectedCellValue(sheetName, 25, 12, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 27, 12, 99d),
					new ExpectedCellValue(sheetName, 28, 12, 99d),
					new ExpectedCellValue(sheetName, 29, 12, 99d),
					new ExpectedCellValue(sheetName, 23, 13, null),
					new ExpectedCellValue(sheetName, 24, 13, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 13, null),
					new ExpectedCellValue(sheetName, 26, 13, null),
					new ExpectedCellValue(sheetName, 27, 13, 99d),
					new ExpectedCellValue(sheetName, 28, 13, 99d),
					new ExpectedCellValue(sheetName, 29, 13, 99d),
					new ExpectedCellValue(sheetName, 23, 14, null),
					new ExpectedCellValue(sheetName, 24, 14, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, null),
					new ExpectedCellValue(sheetName, 27, 14, 99d),
					new ExpectedCellValue(sheetName, 28, 14, 99d),
					new ExpectedCellValue(sheetName, 29, 14, 99d),
					new ExpectedCellValue(sheetName, 23, 15, null),
					new ExpectedCellValue(sheetName, 24, 15, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 15, null),
					new ExpectedCellValue(sheetName, 26, 15, null),
					new ExpectedCellValue(sheetName, 27, 15, 99d),
					new ExpectedCellValue(sheetName, 28, 15, 99d),
					new ExpectedCellValue(sheetName, 29, 15, 99d),
					new ExpectedCellValue(sheetName, 23, 16, null),
					new ExpectedCellValue(sheetName, 24, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 16, null),
					new ExpectedCellValue(sheetName, 26, 16, null),
					new ExpectedCellValue(sheetName, 27, 16, 99d),
					new ExpectedCellValue(sheetName, 28, 16, 99d),
					new ExpectedCellValue(sheetName, 29, 16, 99d),
				});
			}
		}
		#endregion
		#endregion

		#region Helper Methods
		private void CheckPivotTableAddress(ExcelAddress expectedAddress, ExcelAddress pivotTableAddress)
		{
			Assert.AreEqual(expectedAddress.Start.Row, pivotTableAddress.Start.Row);
			Assert.AreEqual(expectedAddress.Start.Column, pivotTableAddress.Start.Column);
			Assert.AreEqual(expectedAddress.End.Row, pivotTableAddress.End.Row);
			Assert.AreEqual(expectedAddress.End.Column, pivotTableAddress.End.Column);
		}
		#endregion
	}
}
