/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau, Evan Schallerer, and others as noted in the source history.
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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotTableTest
	{
		#region Integration Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTablesWorksheetSources.xlsx")]
		public void PivotTableXmlLoadsCorrectly()
		{
			var testFile = new FileInfo(@"PivotTablesWorksheetSources.xlsx");
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			testFile.CopyTo(tempFile.FullName);
			try
			{
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(2, package.Workbook.PivotCacheDefinitions.Count());

					var cacheRecords1 = package.Workbook.PivotCacheDefinitions[0].CacheRecords;
					var cacheRecords2 = package.Workbook.PivotCacheDefinitions[1].CacheRecords;

					Assert.AreNotEqual(cacheRecords1, cacheRecords2);
					Assert.AreEqual(22, cacheRecords1.Count);
					Assert.AreEqual(36, cacheRecords2.Count);
					Assert.AreEqual(cacheRecords1.Count, cacheRecords1.Count);
					Assert.AreEqual(cacheRecords2.Count, cacheRecords2.Count);

					var worksheet1 = package.Workbook.Worksheets["sheet1"];
					var worksheet2 = package.Workbook.Worksheets["sheet2"];
					var worksheet3 = package.Workbook.Worksheets["sheet3"];

					Assert.AreEqual(0, worksheet1.PivotTables.Count());
					Assert.AreEqual(2, worksheet2.PivotTables.Count());
					Assert.AreEqual(1, worksheet3.PivotTables.Count());

					Assert.AreEqual(worksheet2.PivotTables[0].CacheDefinition, worksheet2.PivotTables[1].CacheDefinition);
					Assert.AreNotEqual(worksheet2.PivotTables[0].CacheDefinition, worksheet3.PivotTables[0].CacheDefinition);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}
		#endregion

		#region Refresh Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void PivotTableRefreshFromCacheWithChangedData()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				worksheet.Cells[4, 5].Value = "Blue";
				worksheet.Cells[5, 5].Value = "Green";
				worksheet.Cells[6, 5].Value = "Purple";
				cacheDefinition.UpdateData();
				Assert.AreEqual(4, pivotTable.Fields.Count);
				var pivotField1 = pivotTable.Fields[0];
				Assert.AreEqual(0, pivotField1.Items.Count);
				var pivotField2 = pivotTable.Fields[1];
				Assert.AreEqual(4, pivotField2.Items.Count);
				var pivotField3 = pivotTable.Fields[2];
				Assert.AreEqual(6, pivotField3.Items.Count);
				this.CheckFieldItems(pivotField3);
				var pivotField4 = pivotTable.Fields[3];
				Assert.AreEqual(0, pivotField4.Items.Count);
				Assert.AreEqual(7, pivotTable.RowItems.Count);
				Assert.AreEqual("Blue", worksheet.Cells[11, 9].Value);
				//Assert.AreEqual(100, worksheet.Cells[11, 10].Value);
				Assert.AreEqual("Bike", worksheet.Cells[12, 9].Value);
				//Assert.AreEqual(100, worksheet.Cells[12, 10].Value);
				Assert.AreEqual("Green", worksheet.Cells[13, 9].Value);
				//Assert.AreEqual(90000, worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Car", worksheet.Cells[14, 9].Value);
				//Assert.AreEqual(90000, worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Purple", worksheet.Cells[15, 9].Value);
				//Assert.AreEqual(10, worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Skateboard", worksheet.Cells[16, 9].Value);
				//Assert.AreEqual(10, worksheet.Cells[16, 10].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[17, 9].Value);
				//Assert.AreEqual(90110, worksheet.Cells[17, 10].Value);
				package.SaveAs(new FileInfo("C:\\users\\mcl\\Downloads\\Test.xlsx"));
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void PivotTableRefreshFromCacheWithAddedData()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				worksheet.Cells[7, 3].Value = 4;
				worksheet.Cells[7, 4].Value = "Scooter";
				worksheet.Cells[7, 5].Value = "Purple";
				worksheet.Cells[7, 6].Value = 28;
				cacheDefinition.SourceRange = worksheet.Cells["C3:F7"];
				cacheDefinition.UpdateData();
				Assert.AreEqual(4, pivotTable.Fields.Count);
				var pivotField1 = pivotTable.Fields[0];
				Assert.AreEqual(0, pivotField1.Items.Count);
				var pivotField2 = pivotTable.Fields[1];
				Assert.AreEqual(5, pivotField2.Items.Count);
				var pivotField3 = pivotTable.Fields[2];
				Assert.AreEqual(4, pivotField3.Items.Count);
				this.CheckFieldItems(pivotField3);
				var pivotField4 = pivotTable.Fields[3];
				Assert.AreEqual(0, pivotField4.Items.Count);
				Assert.AreEqual(8, pivotTable.RowItems.Count);
				Assert.AreEqual("Black", worksheet.Cells[11, 9].Value);
				//Assert.AreEqual(110, worksheet.Cells[11, 10].Value);
				Assert.AreEqual("Bike", worksheet.Cells[12, 9].Value);
				//Assert.AreEqual(100, worksheet.Cells[12, 10].Value);
				Assert.AreEqual("Skateboard", worksheet.Cells[13, 9].Value);
				//Assert.AreEqual(10, worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Red", worksheet.Cells[14, 9].Value);
				//Assert.AreEqual(90000, worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Car", worksheet.Cells[15, 9].Value);
				//Assert.AreEqual(90000, worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Purple", worksheet.Cells[16, 9].Value);
				//Assert.AreEqual(28, worksheet.Cells[16, 10].Value);
				Assert.AreEqual("Scooter", worksheet.Cells[17, 9].Value);
				//Assert.AreEqual(28, worksheet.Cells[17, 10].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[18, 9].Value);
				//Assert.AreEqual(90138, worksheet.Cells[18, 10].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void PivotTableRefreshFromCacheRemoveRow()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				cacheDefinition.SourceRange = worksheet.Cells["C3:F5"];
				cacheDefinition.UpdateData();
				Assert.AreEqual(4, pivotTable.Fields.Count);
				var pivotField1 = pivotTable.Fields[0];
				Assert.AreEqual(0, pivotField1.Items.Count);
				var pivotField2 = pivotTable.Fields[1];
				Assert.AreEqual(4, pivotField2.Items.Count);
				var pivotField3 = pivotTable.Fields[2];
				Assert.AreEqual(3, pivotField3.Items.Count);
				this.CheckFieldItems(pivotField3);
				var pivotField4 = pivotTable.Fields[3];
				Assert.AreEqual(0, pivotField4.Items.Count);
				Assert.AreEqual(5, pivotTable.RowItems.Count);
				Assert.AreEqual("Black", worksheet.Cells[11, 9].Value);
				//Assert.AreEqual(100, worksheet.Cells[11, 10].Value);
				Assert.AreEqual("Bike", worksheet.Cells[12, 9].Value);
				//Assert.AreEqual(100, worksheet.Cells[12, 10].Value);
				Assert.AreEqual("Red", worksheet.Cells[13, 9].Value);
				//Assert.AreEqual(90000, worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Car", worksheet.Cells[14, 9].Value);
				//Assert.AreEqual(90000, worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[15, 9].Value);
				//Assert.AreEqual(90100, worksheet.Cells[15, 10].Value);
				Assert.IsNull(worksheet.Cells[16, 9].Value);
			}
		}
		#endregion

		#region Helper Methods
		private void CheckFieldItems(ExcelPivotTableField field)
		{
			int i = 0;
			for (; i < field.Items.Count - 1; i++)
			{
				Assert.AreEqual(i, field.Items[i].X);
			}

			var lastItem = field.Items[field.Items.Count - 1];
			if (string.IsNullOrEmpty(lastItem.T))
				Assert.AreEqual(i, lastItem.X);
		}
		#endregion
	}
}