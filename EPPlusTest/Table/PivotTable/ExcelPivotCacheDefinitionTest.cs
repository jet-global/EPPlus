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
using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotCacheDefinitionTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ConstructExistingExcelPivotCacheDefinition()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var uri = new Uri("xl/pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(uri);
				var cacheDefinition = new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), package, possiblePart, uri);
				Assert.IsNotNull(cacheDefinition);
				Assert.AreEqual(4, cacheDefinition.CacheFields.Count);
				Assert.AreEqual("C3:F6", cacheDefinition.GetSourceRangeAddress().Address);
			}
		}

		[TestMethod]
		public void ConstructEmptyExcelPivotCacheDefinition()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				sheet1.Cells[1, 1].Value = 1;
				sheet1.Cells[2, 1].Value = 2;
				sheet1.Cells[3, 1].Value = 3;
				sheet1.Cells[4, 1].Value = 4;
				sheet1.Cells[1, 2].Value = "a";
				sheet1.Cells[2, 2].Value = "b";
				sheet1.Cells[3, 2].Value = "c";
				sheet1.Cells[4, 2].Value = "d";
				sheet1.Cells[1, 3].Value = true;
				sheet1.Cells[2, 3].Value = true;
				sheet1.Cells[3, 3].Value = true;
				sheet1.Cells[4, 3].Value = false;
				var sourceAddress = sheet1.Cells["A1:C4"];
				var pivotTable = new ExcelPivotTable(sheet1, sheet1.Cells[10, 10], sourceAddress, "pivotTable1", 1);
				var cacheDefinition = new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), pivotTable, sourceAddress, 10);
				Assert.IsNotNull(cacheDefinition);
				Assert.AreEqual(0, cacheDefinition.CacheFields.Count);
				Assert.AreEqual(0, cacheDefinition.CacheRecords.Count);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ExcelPivotCacheDefinitionNullNamespaceManager()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var uri = new Uri("xl/pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(uri);
				new ExcelPivotCacheDefinition(null, package, possiblePart, uri);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ExcelPivotCacheDefinitionNullPackage()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var uri = new Uri("xl/pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(uri);
				new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), null, possiblePart, uri);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ExcelPivotCacheDefinitionNullXmlDocument()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var uri = new Uri("xl/pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
				new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), package, null, uri);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ExcelPivotCacheDefinitionNullCacheUri()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var uri = new Uri("xl/pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(uri);
				new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), package, possiblePart, null);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelPivotCacheDefinitionNullPivotTable()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), null, sheet1.Cells["A1:C4"], 10);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelPivotCacheDefinitionNullSourceAddress()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				sheet1.Cells[1, 1].Value = 1;
				sheet1.Cells[2, 1].Value = 2;
				sheet1.Cells[3, 1].Value = 3;
				sheet1.Cells[4, 1].Value = 4;
				sheet1.Cells[1, 2].Value = "a";
				sheet1.Cells[2, 2].Value = "b";
				sheet1.Cells[3, 2].Value = "c";
				sheet1.Cells[4, 2].Value = "d";
				sheet1.Cells[1, 3].Value = true;
				sheet1.Cells[2, 3].Value = true;
				sheet1.Cells[3, 3].Value = true;
				sheet1.Cells[4, 3].Value = false;
				var pivotTable = new ExcelPivotTable(sheet1, sheet1.Cells[10, 10], sheet1.Cells["A1:C4"], "pivotTable1", 1);
				var cacheDefinition = new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), pivotTable, null, 10);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void ExcelPivotCacheDefinitionInvalidTableId()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				sheet1.Cells[1, 1].Value = 1;
				sheet1.Cells[2, 1].Value = 2;
				sheet1.Cells[3, 1].Value = 3;
				sheet1.Cells[4, 1].Value = 4;
				sheet1.Cells[1, 2].Value = "a";
				sheet1.Cells[2, 2].Value = "b";
				sheet1.Cells[3, 2].Value = "c";
				sheet1.Cells[4, 2].Value = "d";
				sheet1.Cells[1, 3].Value = true;
				sheet1.Cells[2, 3].Value = true;
				sheet1.Cells[3, 3].Value = true;
				sheet1.Cells[4, 3].Value = false;
				var sourceAddress = sheet1.Cells["A1:C4"];
				var pivotTable = new ExcelPivotTable(sheet1, sheet1.Cells[10, 10], sourceAddress, "pivotTable1", 1);
				var cacheDefinition = new ExcelPivotCacheDefinition(TestUtility.CreateDefaultNSM(), pivotTable, sourceAddress, -1);
			}
		}
		#endregion

		#region UpdateData Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateDataTest()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				worksheet.Cells[3, 3].Value = "Item No.";
				worksheet.Cells[5, 4].Value = "Scooter";
				worksheet.Cells[5, 5].Value = "Yellow";
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				cacheDefinition.UpdateData();
				Assert.AreEqual(4, cacheDefinition.CacheFields.Count);
				Assert.AreEqual("Item No.", cacheDefinition.CacheFields.First().Name);
				var cacheRecords = cacheDefinition.CacheRecords;
				this.AssertCacheRecord(cacheRecords, 0, 0, PivotCacheRecordType.n, "1");
				this.AssertCacheRecord(cacheRecords, 1, 0, PivotCacheRecordType.n, "2");
				this.AssertCacheRecord(cacheRecords, 2, 0, PivotCacheRecordType.n, "3");
				this.AssertCacheRecord(cacheRecords, 0, 1, PivotCacheRecordType.x, "0");
				this.AssertCacheRecord(cacheRecords, 1, 1, PivotCacheRecordType.x, "3");
				this.AssertCacheRecord(cacheRecords, 2, 1, PivotCacheRecordType.x, "2");
				this.AssertCacheRecord(cacheRecords, 0, 2, PivotCacheRecordType.x, "0");
				this.AssertCacheRecord(cacheRecords, 1, 2, PivotCacheRecordType.x, "2");
				this.AssertCacheRecord(cacheRecords, 2, 2, PivotCacheRecordType.x, "0");
				this.AssertCacheRecord(cacheRecords, 0, 3, PivotCacheRecordType.n, "100");
				this.AssertCacheRecord(cacheRecords, 1, 3, PivotCacheRecordType.n, "90000");
				this.AssertCacheRecord(cacheRecords, 2, 3, PivotCacheRecordType.n, "10");
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableBackedByExcelTable.xlsx")]
		public void UpdateDataOfPivotTableBackedByExcelTableTest()
		{
			var file = new FileInfo("PivotTableBackedByExcelTable.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				cacheDefinition.UpdateData();
				Assert.AreEqual("Transaction", cacheDefinition.CacheFields.First().Name);
				Assert.AreEqual(3, cacheDefinition.CacheFields[2].SharedItems.Count);
				Assert.AreEqual("January", worksheet.Cells[17, 2].Value);
				Assert.AreEqual("February", worksheet.Cells[18, 2].Value);
				Assert.AreEqual("March", worksheet.Cells[19, 2].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[20, 2].Value);
				Assert.AreEqual(2078.75, worksheet.Cells[17, 3].Value);
				Assert.AreEqual(1293d, worksheet.Cells[18, 3].Value);
				Assert.AreEqual(856.49, worksheet.Cells[19, 3].Value);
				Assert.AreEqual(4228.24, worksheet.Cells[20, 3].Value);
			}
		}
		#endregion

		#region GetRelatedPivotTables Tests
		[TestMethod]
		public void GetRelatedPivotTables()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("sheet1");
				sheet1.Cells[1, 1].Value = 1;
				sheet1.Cells[2, 1].Value = 2;
				sheet1.Cells[3, 1].Value = 3;
				sheet1.Cells[4, 1].Value = 4;
				sheet1.Cells[1, 2].Value = "a";
				sheet1.Cells[2, 2].Value = "b";
				sheet1.Cells[3, 2].Value = "c";
				sheet1.Cells[4, 2].Value = "d";
				sheet1.Cells[1, 3].Value = true;
				sheet1.Cells[2, 3].Value = true;
				sheet1.Cells[3, 3].Value = true;
				sheet1.Cells[4, 3].Value = false;
				var pivotTable1 = new ExcelPivotTable(sheet1, sheet1.Cells[10, 10], sheet1.Cells["A1:D3"], "pivotTable1", 1);
				var pivotTable2 = new ExcelPivotTable(sheet1, sheet1.Cells[50, 10], sheet1.Cells["A1:D3"], "pivotTable2", 2);
				var pivotTables = package.Workbook.PivotCacheDefinitions.Single().GetRelatedPivotTables();
				Assert.AreEqual(2, pivotTables.Count);
				Assert.IsTrue(pivotTables.Any(p => p.Name == pivotTable1.Name));
				Assert.IsTrue(pivotTables.Any(p => p.Name == pivotTable2.Name));
			}
		}
		#endregion

		#region Helper Methods
		private void AssertCacheRecord(ExcelPivotCacheRecords records, int row, int col, PivotCacheRecordType type, string value)
		{
			Assert.AreEqual(value, records[row].Items[col].Value);
			Assert.AreEqual(type, records[row].Items[col].Type);
		}
		#endregion
	}
}