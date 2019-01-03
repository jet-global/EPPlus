using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotTableFieldTest
	{
		#region DisableDefaultSubtotal Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableBackedByExcelTable.xlsx")]
		public void DisableDefaultSubtotalRemovesDefaultItem()
		{
			var file = new FileInfo("PivotTableBackedByExcelTable.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					Assert.IsTrue(pivotTable.Fields.All(f => f.DefaultSubtotal));
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					var field = pivotTable.Fields[2];
					Assert.AreEqual(4, field.Items.Count);
					Assert.AreEqual("default", field.Items[3].T);
					field = pivotTable.Fields[3];
					Assert.AreEqual(5, field.Items.Count);
					Assert.AreEqual("default", field.Items[4].T);
					foreach (var ptField in pivotTable.Fields)
					{
						ptField.DisableDefaultSubtotal();
					}
					Assert.IsFalse(pivotTable.Fields.All(f => f.DefaultSubtotal));
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					field = pivotTable.Fields[2];
					Assert.AreEqual(3, field.Items.Count);
					Assert.IsTrue(string.IsNullOrEmpty(field.Items[2].T));
					field = pivotTable.Fields[3];
					Assert.AreEqual(4, field.Items.Count);
					Assert.IsTrue(string.IsNullOrEmpty(field.Items[3].T));
				}
			}
		}
		#endregion
	}
}
