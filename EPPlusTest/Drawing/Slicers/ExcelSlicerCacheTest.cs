using System.Collections.Generic;
using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class ExcelSlicerCacheTest
	{
		#region Refresh Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\Slicers.xlsx")]
		public void RefreshSlicersWithStaticData()
		{
			var file = new FileInfo("Slicers.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					package.Workbook.PivotCacheDefinitions.First().UpdateData();
					package.SaveAs(newFile.File);
				}

				var postingDateSlicerValues = new List<SlicerCacheValue>
				{
					new SlicerCacheValue("Jan"),
					new SlicerCacheValue("Feb"),
					new SlicerCacheValue("Mar"),
					new SlicerCacheValue("Apr"),
					new SlicerCacheValue("May"),
					new SlicerCacheValue("Jun"),
					new SlicerCacheValue("Jul"),
					new SlicerCacheValue("Aug"),
					new SlicerCacheValue("Sep"),
					new SlicerCacheValue("Oct"),
					new SlicerCacheValue("Nov"),
					new SlicerCacheValue("Dec"),
					new SlicerCacheValue("<12/31/2012"),
					new SlicerCacheValue(">1/2/2013"),
				};

				var sourceCodeSlicerValues = new List<SlicerCacheValue>
				{
					new SlicerCacheValue("END"),
					new SlicerCacheValue("MIDDLE"),
					new SlicerCacheValue("PURCHASES"),
					new SlicerCacheValue("SALES"),
					new SlicerCacheValue("START")
				};

				var descriptionSlicerValues = new List<SlicerCacheValue>
				{
					new SlicerCacheValue("Entries, January 2013"),
					new SlicerCacheValue("Opening Entry"),
					new SlicerCacheValue("Order 106015"),
					new SlicerCacheValue("Order 106018")
				};

				var amountSlicerValues = new List<SlicerCacheValue>
				{
					new SlicerCacheValue("-752562.89"),
					new SlicerCacheValue("-558283.32"),
					new SlicerCacheValue("-244909.87"),
					new SlicerCacheValue("-53800.14"),
					new SlicerCacheValue("82.28"),
					new SlicerCacheValue("122.09"),
					new SlicerCacheValue("183.13"),
					new SlicerCacheValue("232.02"),
					new SlicerCacheValue("286.31"),
					new SlicerCacheValue("305.22"),
					new SlicerCacheValue("828.97"),
					new SlicerCacheValue("1243.45"),
					new SlicerCacheValue("2072.43"),
					new SlicerCacheValue("3734.73"),
					new SlicerCacheValue("5602.1"),
					new SlicerCacheValue("5662.08"),
					new SlicerCacheValue("8277.85"),
					new SlicerCacheValue("9336.83"),
					new SlicerCacheValue("11836.03"),
					new SlicerCacheValue("53800.14"),
				};

				using (var package = new ExcelPackage(newFile.File))
				{
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					var postingDateSlicerCache = package.Workbook.SlicerCaches.First(c => c.Name == "Slicer_Posting_Date");
					var descriptionSlicerCache = package.Workbook.SlicerCaches.First(c => c.Name == "Slicer_Description");
					var amountSlicerCache = package.Workbook.SlicerCaches.First(c => c.Name == "Slicer_Amount");
					var sourceCodeSlicerCache = package.Workbook.SlicerCaches.First(c => c.Name == "Slicer_Source_Code");
					this.ValidateSlicer(postingDateSlicerValues, postingDateSlicerCache, cacheDefinition);
					this.ValidateSlicer(descriptionSlicerValues, descriptionSlicerCache, cacheDefinition);
					this.ValidateSlicer(sourceCodeSlicerValues, sourceCodeSlicerCache, cacheDefinition);
					this.ValidateSlicer(amountSlicerValues, amountSlicerCache, cacheDefinition);
				}
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\Slicers.xlsx")]
		public void RefreshSlicersWithStaticDataNoCustomListSorting()
		{
			var file = new FileInfo("Slicers.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					package.Workbook.PivotCacheDefinitions.First().UpdateData();
					package.SaveAs(newFile.File);
				}

				var postingDateSlicerValues = new List<SlicerCacheValue>
				{
					new SlicerCacheValue("<12/31/2012"),
					new SlicerCacheValue(">1/2/2013"),
					new SlicerCacheValue("Apr"),
					new SlicerCacheValue("Aug"),
					new SlicerCacheValue("Dec"),
					new SlicerCacheValue("Feb"),
					new SlicerCacheValue("Jan"),
					new SlicerCacheValue("Jul"),
					new SlicerCacheValue("Jun"),
					new SlicerCacheValue("Mar"),
					new SlicerCacheValue("May"),
					new SlicerCacheValue("Nov"),
					new SlicerCacheValue("Oct"),
					new SlicerCacheValue("Sep"),
				};

				var amountSlicerValues = new List<SlicerCacheValue>
				{
					new SlicerCacheValue("53800.14"),
					new SlicerCacheValue("11836.03"),
					new SlicerCacheValue("9336.83"),
					new SlicerCacheValue("8277.85"),
					new SlicerCacheValue("5662.08"),
					new SlicerCacheValue("5602.1"),
					new SlicerCacheValue("3734.73"),
					new SlicerCacheValue("2072.43"),
					new SlicerCacheValue("1243.45"),
					new SlicerCacheValue("828.97"),
					new SlicerCacheValue("305.22"),
					new SlicerCacheValue("286.31"),
					new SlicerCacheValue("232.02"),
					new SlicerCacheValue("183.13"),
					new SlicerCacheValue("122.09"),
					new SlicerCacheValue("82.28"),
					new SlicerCacheValue("-53800.14"),
					new SlicerCacheValue("-244909.87"),
					new SlicerCacheValue("-558283.32"),
					new SlicerCacheValue("-752562.89")
				};

				using (var package = new ExcelPackage(newFile.File))
				{
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					var postingDateSlicerCache = package.Workbook.SlicerCaches.First(c => c.Name == "Slicer_Posting_Date1");
					var amountSlicerCache = package.Workbook.SlicerCaches.First(c => c.Name == "Slicer_Amount1");
					this.ValidateSlicer(postingDateSlicerValues, postingDateSlicerCache, cacheDefinition);
					this.ValidateSlicer(amountSlicerValues, amountSlicerCache, cacheDefinition);
				}
			}
		}
		#endregion

		#region Helper Methods
		private void ValidateSlicer(List<SlicerCacheValue> expectedValues, ExcelSlicerCache slicerCache, ExcelPivotCacheDefinition cacheDefinition)
		{
			Assert.AreEqual(expectedValues.Count, slicerCache.TabularDataNode.Items.Count);
			var cacheField = cacheDefinition.CacheFields.First(c => c.Name.IsEquivalentTo(slicerCache.SourceName));
			var items = cacheField.IsGroupField ? cacheField.FieldGroup.GroupItems : cacheField.SharedItems;

			for (int i = 0; i < expectedValues.Count; i++)
			{
				var expected = expectedValues[i];
				var cacheItem = slicerCache.TabularDataNode.Items[i];
				Assert.AreEqual(expected.IsSelected, cacheItem.IsSelected);
				Assert.AreEqual(expected.NoData, cacheItem.NoData);

				var actual = items[cacheItem.AtomIndex];
				this.CompareValue(expected.Value, actual);
			}
		}

		private void CompareValue(string expected, CacheItem actual)
		{
			if (actual.Type == PivotCacheRecordType.n)
			{
				var expectedNumeric = double.Parse(expected);
				var actualNumeric = double.Parse(actual.Value);
				Assert.AreEqual(expectedNumeric, actualNumeric, .0000001);
			}
			else
				Assert.AreEqual(expected, actual.Value);
		}
		#endregion

		#region Nested Classes
		private class SlicerCacheValue
		{
			#region Properties
			public string Value { get; set; }

			public bool IsSelected { get; set; }

			public bool NoData { get; set; }
			#endregion

			#region Constructors
			public SlicerCacheValue(string value, bool isSelected = true, bool nonDisplay = false)
			{
				this.Value = value;
				this.IsSelected = isSelected;
				this.NoData = nonDisplay;
			}
			#endregion
		}
		#endregion
	}
}
