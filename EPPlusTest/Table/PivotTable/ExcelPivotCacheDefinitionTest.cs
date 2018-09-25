using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotCacheDefinitionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ExcelPivotCacheDefinitionConstructorTest()
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
				var pivotTable = new ExcelPivotTable(sheet1, sheet1.Cells[10, 10], sheet1.Cells["A1:D3"], "pivotTable1", 1);
				Assert.AreEqual(1, package.Workbook.PivotCacheDefinitions.Count);
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				Assert.AreEqual(0, cacheDefinition.CacheRecords.Records.Count);
			}
		}
		#endregion

		#region UpdateRecords Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateRecordsTest()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				worksheet.Cells[3, 3].Value = "Item No.";
				worksheet.Cells[5, 5].Value = "Blue";
				worksheet.Cells[6, 6].Value = 50;
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				cacheDefinition.UpdateRecords();
				Assert.AreEqual(4, cacheDefinition.CacheFields.Count);
				Assert.AreEqual("Item No.", cacheDefinition.CacheFields.First().Name);
				var cacheRecords = cacheDefinition.CacheRecords;
				this.AssertCacheRecord(cacheRecords, 0, 0, PivotCacheRecordType.n, "1");
				this.AssertCacheRecord(cacheRecords, 1, 0, PivotCacheRecordType.n, "2");
				this.AssertCacheRecord(cacheRecords, 2, 0, PivotCacheRecordType.n, "3");
				this.AssertCacheRecord(cacheRecords, 0, 1, PivotCacheRecordType.x, "0");
				this.AssertCacheRecord(cacheRecords, 1, 1, PivotCacheRecordType.x, "1");
				this.AssertCacheRecord(cacheRecords, 2, 1, PivotCacheRecordType.x, "2");
				this.AssertCacheRecord(cacheRecords, 0, 2, PivotCacheRecordType.x, "0");
				this.AssertCacheRecord(cacheRecords, 1, 2, PivotCacheRecordType.x, "2");
				this.AssertCacheRecord(cacheRecords, 2, 2, PivotCacheRecordType.x, "0");
				this.AssertCacheRecord(cacheRecords, 0, 3, PivotCacheRecordType.n, "100");
				this.AssertCacheRecord(cacheRecords, 1, 3, PivotCacheRecordType.n, "90000");
				this.AssertCacheRecord(cacheRecords, 2, 3, PivotCacheRecordType.n, "50");
			}
		}
		#endregion

		#region Helper Methods
		private void AssertCacheRecord(ExcelPivotCacheRecords records, int row, int col, PivotCacheRecordType type, string value)
		{
			Assert.AreEqual(value, records.Records[row].Items[col].Value);
			Assert.AreEqual(type, records.Records[row].Items[col].Type);
		}
		#endregion
	}
}