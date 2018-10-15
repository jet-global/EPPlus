using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotCacheRecordsTest
	{
		#region Constructor Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ConstructExcelPivotCacheRecords()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var ns = TestUtility.CreateDefaultNSM();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				var records = new ExcelPivotCacheRecords(ns, package, possiblePart, partUri, cacheDefinition);
				Assert.AreEqual(3, records.Count);
				Assert.AreEqual(4, records[0].Items.Count);
				Assert.AreEqual(4, records[1].Items.Count);
				Assert.AreEqual(4, records[2].Items.Count);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void ConstructEmptyExcelPivotCacheRecords()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var ns = TestUtility.CreateDefaultNSM();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				int tableId = 2;
				var records = new ExcelPivotCacheRecords(ns, package.Package, ref tableId, cacheDefinition);
				Assert.IsNotNull(records);
				Assert.AreEqual(0, records.Count);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void EmptyExcelPivotCacheRecordsNullNamespaceManager()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				new ExcelPivotCacheRecords(null, package, possiblePart, partUri, cacheDefinition);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void EmptyExcelPivotCacheRecordsNullcacheRecordsXml()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				new ExcelPivotCacheRecords(TestUtility.CreateDefaultNSM(), package, null, partUri, cacheDefinition);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void EmptyExcelPivotCacheRecordsNullTargetUri()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				new ExcelPivotCacheRecords(TestUtility.CreateDefaultNSM(), package, possiblePart, null, cacheDefinition);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void EmptyExcelPivotCacheRecordsNullCacheDefinition()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				new ExcelPivotCacheRecords(TestUtility.CreateDefaultNSM(), package, possiblePart, partUri, null);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void EmptyExcelPivotCacheRecordsNullPackage()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				int tableId = 2;
				new ExcelPivotCacheRecords(TestUtility.CreateDefaultNSM(), null, ref tableId, cacheDefinition);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void EmptyExcelPivotCacheRecordsInvalidTableId()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var partUri = new Uri("xl/pivotCache/pivotCacheRecords1.xml", UriKind.Relative);
				var possiblePart = package.GetXmlFromUri(partUri);
				int tableId = 0;
				new ExcelPivotCacheRecords(TestUtility.CreateDefaultNSM(), package.Package, ref tableId, cacheDefinition);
			}
		}
		#endregion

		#region UpdateRecords Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateRecordsExistingData()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var cacheRecords = cacheDefinition.CacheRecords;
				var worksheet = package.Workbook.Worksheets.First();
				worksheet.Cells[5, 3].Value = 5;
				worksheet.Cells[5, 4].Value = "Scooter";
				worksheet.Cells[5, 5].Value = "Orange";
				worksheet.Cells[5, 6].Value = 98;
				cacheRecords.UpdateRecords(worksheet.Cells["C4:F6"]);
				var record1 = cacheRecords[0];
				var record2 = cacheRecords[1];
				var record3 = cacheRecords[2];
				Assert.AreEqual(3, cacheRecords.Count);
				// record 1
				this.AssertCacheItem(record1.Items[0], "1", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record1.Items[1], 1, cacheDefinition), "Bike", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record1.Items[2], 2, cacheDefinition), "Black", PivotCacheRecordType.s);
				this.AssertCacheItem(record1.Items[3], "100", PivotCacheRecordType.n);
				//record 2
				this.AssertCacheItem(record2.Items[0], "5", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record2.Items[1], 1, cacheDefinition), "Scooter", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record2.Items[2], 2, cacheDefinition), "Orange", PivotCacheRecordType.s);
				this.AssertCacheItem(record2.Items[3], "98", PivotCacheRecordType.n);
				// record 3
				this.AssertCacheItem(record3.Items[0], "3", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record3.Items[1], 1, cacheDefinition), "Skateboard", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record3.Items[2], 2, cacheDefinition), "Black", PivotCacheRecordType.s);
				this.AssertCacheItem(record3.Items[3], "10", PivotCacheRecordType.n);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateRecordsAddNewRecord()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var cacheRecords = cacheDefinition.CacheRecords;
				var worksheet = package.Workbook.Worksheets.First();
				worksheet.Cells[7, 3].Value = 5;
				worksheet.Cells[7, 4].Value = "Scooter";
				worksheet.Cells[7, 5].Value = "Orange";
				worksheet.Cells[7, 6].Value = 98;
				cacheRecords.UpdateRecords(worksheet.Cells["C4:F7"]);
				var record1 = cacheRecords[0];
				var record2 = cacheRecords[1];
				var record3 = cacheRecords[2];
				var record4 = cacheRecords[3];
				Assert.AreEqual(4, cacheRecords.Count);
				// record 1
				this.AssertCacheItem(record1.Items[0], "1", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record1.Items[1], 1, cacheDefinition), "Bike", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record1.Items[2], 2, cacheDefinition), "Black", PivotCacheRecordType.s);
				this.AssertCacheItem(record1.Items[3], "100", PivotCacheRecordType.n);
				//record 2
				this.AssertCacheItem(record2.Items[0], "2", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record2.Items[1], 1, cacheDefinition), "Car", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record2.Items[2], 2, cacheDefinition), "Red", PivotCacheRecordType.s);
				this.AssertCacheItem(record2.Items[3], "90000", PivotCacheRecordType.n);
				// record 3
				this.AssertCacheItem(record3.Items[0], "3", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record3.Items[1], 1, cacheDefinition), "Skateboard", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record3.Items[2], 2, cacheDefinition), "Black", PivotCacheRecordType.s);
				this.AssertCacheItem(record3.Items[3], "10", PivotCacheRecordType.n);
				// record 4
				this.AssertCacheItem(record4.Items[0], "5", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record4.Items[1], 1, cacheDefinition), "Scooter", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record4.Items[2], 2, cacheDefinition), "Orange", PivotCacheRecordType.s);
				this.AssertCacheItem(record4.Items[3], "98", PivotCacheRecordType.n);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateRecordsRemoveRecord()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var cacheRecords = cacheDefinition.CacheRecords;
				var worksheet = package.Workbook.Worksheets.First();
				cacheRecords.UpdateRecords(worksheet.Cells["C4:F5"]);
				var record1 = cacheRecords[0];
				var record2 = cacheRecords[1];
				Assert.AreEqual(2, cacheRecords.Count);
				// record 1
				this.AssertCacheItem(record1.Items[0], "1", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record1.Items[1], 1, cacheDefinition), "Bike", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record1.Items[2], 2, cacheDefinition), "Black", PivotCacheRecordType.s);
				this.AssertCacheItem(record1.Items[3], "100", PivotCacheRecordType.n);
				//record 2
				this.AssertCacheItem(record2.Items[0], "2", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(record2.Items[1], 1, cacheDefinition), "Car", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(record2.Items[2], 2, cacheDefinition), "Red", PivotCacheRecordType.s);
				this.AssertCacheItem(record2.Items[3], "90000", PivotCacheRecordType.n);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateRecordsRemoveMultipleRecords()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var cacheRecords = cacheDefinition.CacheRecords;
				var worksheet = package.Workbook.Worksheets.First();
				cacheRecords.UpdateRecords(worksheet.Cells["C4:F4"]);
				Assert.AreEqual(1, cacheRecords.Count);
				this.AssertCacheItem(cacheRecords[0].Items[0], "1", PivotCacheRecordType.n);
				this.AssertCacheItem(this.ResolveXCacheItem(cacheRecords[0].Items[1], 1, cacheDefinition), "Bike", PivotCacheRecordType.s);
				this.AssertCacheItem(this.ResolveXCacheItem(cacheRecords[0].Items[2], 2, cacheDefinition), "Black", PivotCacheRecordType.s);
				this.AssertCacheItem(cacheRecords[0].Items[3], "100", PivotCacheRecordType.n);
			}
		}
		#endregion

		#region Helper Methods
		private CacheItem ResolveXCacheItem(CacheItem item, int fieldIndex, ExcelPivotCacheDefinition cacheDefinition)
		{
			if (item.Type != PivotCacheRecordType.x)
				throw new InvalidOperationException("The cache item was not a reference item.");
			int sharedItemIndex = int.Parse(item.Value);
			return cacheDefinition.CacheFields[fieldIndex].SharedItems[sharedItemIndex];
		}

		private void AssertCacheItem(CacheItem item, string value, PivotCacheRecordType type)
		{
			Assert.AreEqual(value, item.Value);
			Assert.AreEqual(type, item.Type);
		}
		#endregion
	}
}