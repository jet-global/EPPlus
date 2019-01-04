using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable.DataFieldFunctionTypes
{
	[TestClass]
	public class DataFieldFunctionProductTypeTest : DataFieldFunctionTypeTestBase
	{
		#region SingleDataField Override Methods
		public override void ConfigurePivotTableDataFieldFunction()
		{
			base.PivotTable.DataFields.First().Function = DataFieldFunctions.Product;
			base.PivotTable.DataFields.First().Name = "Product of Total";
		}

		public override void ValidatePivotTableRefreshDataField(FileInfo file, string sheetName)
		{
			TestHelperUtility.ValidateWorksheet(file, sheetName, new[]
			{
				new ExpectedCellValue(sheetName, 14, 1, "January"),
				new ExpectedCellValue(sheetName, 15, 1, "February"),
				new ExpectedCellValue(sheetName, 16, 1, "March"),
				new ExpectedCellValue(sheetName, 17, 1, "Grand Total"),

				new ExpectedCellValue(sheetName, 13, 2, "Car Rack"),
				new ExpectedCellValue(sheetName, 14, 2, 287446327.94),
				new ExpectedCellValue(sheetName, 16, 2, 831.5),
				new ExpectedCellValue(sheetName, 17, 2, 239011621680.03),

				new ExpectedCellValue(sheetName, 13, 3, "Headlamp"),
				new ExpectedCellValue(sheetName, 16, 3, 24.99),
				new ExpectedCellValue(sheetName, 17, 3, 24.99),

				new ExpectedCellValue(sheetName, 13, 4, "Sleeping Bag"),
				new ExpectedCellValue(sheetName, 15, 4, 99d),
				new ExpectedCellValue(sheetName, 17, 4, 99d),

				new ExpectedCellValue(sheetName, 13, 5, "Tent"),
				new ExpectedCellValue(sheetName, 15, 5, 1194d),
				new ExpectedCellValue(sheetName, 17, 5, 1194d),

				new ExpectedCellValue(sheetName, 13, 6, "Grand Total"),
				new ExpectedCellValue(sheetName, 14, 6, 287446327.94),
				new ExpectedCellValue(sheetName, 15, 6, 118206),
				new ExpectedCellValue(sheetName, 16, 6, 20779.18),
				new ExpectedCellValue(sheetName, 17, 6, 7.0603266773022118E+17)
			});
		}
		#endregion

		#region MultipleColumnDataFields Override Methods
		public override void ConfigurePivotTableMultipleColumnDataFieldsFunction()
		{
			base.PivotTable.DataFields[0].Function = DataFieldFunctions.Min;
			base.PivotTable.DataFields[0].Name = "Min of Total";
			base.PivotTable.DataFields[1].Function = DataFieldFunctions.Product;
			base.PivotTable.DataFields[1].Name = "Product of Units Sold";
		}

		public override void ValidatePivotTableRefreshMultipleColumnDataFields(FileInfo file, string sheetName)
		{
			TestHelperUtility.ValidateWorksheet(file, sheetName, new[]
			{
				new ExpectedCellValue(sheetName, 24, 1, "January"),
				new ExpectedCellValue(sheetName, 25, 1, "Car Rack"),
				new ExpectedCellValue(sheetName, 26, 1, "February"),
				new ExpectedCellValue(sheetName, 27, 1, "Sleeping Bag"),
				new ExpectedCellValue(sheetName, 28, 1, "Tent"),
				new ExpectedCellValue(sheetName, 29, 1, "March"),
				new ExpectedCellValue(sheetName, 30, 1, "Car Rack"),
				new ExpectedCellValue(sheetName, 31, 1, "Headlamp"),
				new ExpectedCellValue(sheetName, 32, 1, "Grand Total"),

				new ExpectedCellValue(sheetName, 22, 2, "Chicago"),
				new ExpectedCellValue(sheetName, 23, 2, "Min of Total"),
				new ExpectedCellValue(sheetName, 24, 2, 831.5),
				new ExpectedCellValue(sheetName, 25, 2, 831.5),
				new ExpectedCellValue(sheetName, 29, 2, 24.99),
				new ExpectedCellValue(sheetName, 31, 2, 24.99),
				new ExpectedCellValue(sheetName, 32, 2, 24.99),

				new ExpectedCellValue(sheetName, 23, 3, "Product of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 3, 2d),
				new ExpectedCellValue(sheetName, 25, 3, 2d),
				new ExpectedCellValue(sheetName, 29, 3, 1d),
				new ExpectedCellValue(sheetName, 31, 3, 1d),
				new ExpectedCellValue(sheetName, 32, 3, 2d),

				new ExpectedCellValue(sheetName, 22, 4, "Nashville"),
				new ExpectedCellValue(sheetName, 23, 4, "Min of Total"),
				new ExpectedCellValue(sheetName, 24, 4, 831.5),
				new ExpectedCellValue(sheetName, 25, 4, 831.5),
				new ExpectedCellValue(sheetName, 26, 4, 1194d),
				new ExpectedCellValue(sheetName, 28, 4, 1194d),
				new ExpectedCellValue(sheetName, 29, 4, 831.5),
				new ExpectedCellValue(sheetName, 30, 4, 831.5),
				new ExpectedCellValue(sheetName, 32, 4, 831.5),

				new ExpectedCellValue(sheetName, 23, 5, "Product of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 5, 2d),
				new ExpectedCellValue(sheetName, 25, 5, 2d),
				new ExpectedCellValue(sheetName, 26, 5, 6d),
				new ExpectedCellValue(sheetName, 28, 5, 6d),
				new ExpectedCellValue(sheetName, 29, 5, 2d),
				new ExpectedCellValue(sheetName, 30, 5, 2d),
				new ExpectedCellValue(sheetName, 32, 5, 24d),

				new ExpectedCellValue(sheetName, 22, 6, "San Francisco"),
				new ExpectedCellValue(sheetName, 23, 6, "Min of Total"),
				new ExpectedCellValue(sheetName, 24, 6, 415.75),
				new ExpectedCellValue(sheetName, 25, 6, 415.75),
				new ExpectedCellValue(sheetName, 26, 6, 99d),
				new ExpectedCellValue(sheetName, 26, 6, 99d),
				new ExpectedCellValue(sheetName, 32, 6, 99d),

				new ExpectedCellValue(sheetName, 23, 7, "Product of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 7, 1d),
				new ExpectedCellValue(sheetName, 25, 7, 1d),
				new ExpectedCellValue(sheetName, 26, 7, 1d),
				new ExpectedCellValue(sheetName, 27, 7, 1d),
				new ExpectedCellValue(sheetName, 32, 7, 1d),

				new ExpectedCellValue(sheetName, 22, 8, "Total Min of Total"),
				new ExpectedCellValue(sheetName, 24, 8, 415.75),
				new ExpectedCellValue(sheetName, 25, 8, 415.75),
				new ExpectedCellValue(sheetName, 26, 8, 99d),
				new ExpectedCellValue(sheetName, 27, 8, 99d),
				new ExpectedCellValue(sheetName, 28, 8, 1194d),
				new ExpectedCellValue(sheetName, 29, 8, 24.99),
				new ExpectedCellValue(sheetName, 30, 8, 831.5),
				new ExpectedCellValue(sheetName, 31, 8, 24.99),
				new ExpectedCellValue(sheetName, 32, 8, 24.99),

				new ExpectedCellValue(sheetName, 22, 9, "Total Product of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 9, 4d),
				new ExpectedCellValue(sheetName, 25, 9, 4d),
				new ExpectedCellValue(sheetName, 26, 9, 6d),
				new ExpectedCellValue(sheetName, 27, 9, 1d),
				new ExpectedCellValue(sheetName, 28, 9, 6d),
				new ExpectedCellValue(sheetName, 29, 9, 2d),
				new ExpectedCellValue(sheetName, 30, 9, 2d),
				new ExpectedCellValue(sheetName, 31, 9, 1d),
				new ExpectedCellValue(sheetName, 32, 9, 48d)
			});
		}
		#endregion

		#region MultipleRowDataFields Override Methods
		public override void ConfigurePivotTableMultipleRowDataFieldsFunction()
		{
			base.PivotTable.DataFields[0].Function = DataFieldFunctions.Max;
			base.PivotTable.DataFields[0].Name = "Max of Total";
			base.PivotTable.DataFields[1].Function = DataFieldFunctions.Product;
			base.PivotTable.DataFields[1].Name = "Product of Units Sold";
		}

		public override void ValidatePivotTableRefreshMultipleRowDataFields(FileInfo file, string sheetName)
		{
			TestHelperUtility.ValidateWorksheet(file, sheetName, new[]
			{
				new ExpectedCellValue(sheetName, 39, 1, "January"),
				new ExpectedCellValue(sheetName, 40, 1, "Max of Total"),
				new ExpectedCellValue(sheetName, 41, 1, "Product of Units Sold"),
				new ExpectedCellValue(sheetName, 42, 1, "February"),
				new ExpectedCellValue(sheetName, 43, 1, "Max of Total"),
				new ExpectedCellValue(sheetName, 44, 1, "Product of Units Sold"),
				new ExpectedCellValue(sheetName, 45, 1, "March"),
				new ExpectedCellValue(sheetName, 46, 1, "Max of Total"),
				new ExpectedCellValue(sheetName, 47, 1, "Product of Units Sold"),
				new ExpectedCellValue(sheetName, 48, 1, "Total Max of Total"),
				new ExpectedCellValue(sheetName, 49, 1, "Total Product of Units Sold"),

				new ExpectedCellValue(sheetName, 37, 2, "Chicago"),
				new ExpectedCellValue(sheetName, 38, 2, "Car Rack"),
				new ExpectedCellValue(sheetName, 40, 2, 831.5),
				new ExpectedCellValue(sheetName, 41, 2, 2d),
				new ExpectedCellValue(sheetName, 48, 2, 831.5),
				new ExpectedCellValue(sheetName, 49, 2, 2d),

				new ExpectedCellValue(sheetName, 38, 3, "Headlamp"),
				new ExpectedCellValue(sheetName, 46, 3, 24.99),
				new ExpectedCellValue(sheetName, 47, 3, 1d),
				new ExpectedCellValue(sheetName, 48, 3, 24.99),
				new ExpectedCellValue(sheetName, 49, 3, 1d),

				new ExpectedCellValue(sheetName, 37, 4, "Chicago Total"),
				new ExpectedCellValue(sheetName, 40, 4, 831.5),
				new ExpectedCellValue(sheetName, 41, 4, 2d),
				new ExpectedCellValue(sheetName, 46, 4, 24.99),
				new ExpectedCellValue(sheetName, 47, 4, 1d),
				new ExpectedCellValue(sheetName, 48, 4, 831.5),
				new ExpectedCellValue(sheetName, 49, 4, 2d),

				new ExpectedCellValue(sheetName, 37, 5, "Nashville"),
				new ExpectedCellValue(sheetName, 38, 5, "Car Rack"),
				new ExpectedCellValue(sheetName, 40, 5, 831.5),
				new ExpectedCellValue(sheetName, 41, 5, 2d),
				new ExpectedCellValue(sheetName, 46, 5, 831.5),
				new ExpectedCellValue(sheetName, 47, 5, 2d),
				new ExpectedCellValue(sheetName, 48, 5, 831.5),
				new ExpectedCellValue(sheetName, 49, 5, 4d),

				new ExpectedCellValue(sheetName, 38, 6, "Tent"),
				new ExpectedCellValue(sheetName, 43, 6, 1194d),
				new ExpectedCellValue(sheetName, 44, 6, 6d),
				new ExpectedCellValue(sheetName, 48, 6, 1194d),
				new ExpectedCellValue(sheetName, 49, 6, 6d),

				new ExpectedCellValue(sheetName, 37, 7, "Nashville Total"),
				new ExpectedCellValue(sheetName, 40, 7, 831.5),
				new ExpectedCellValue(sheetName, 41, 7, 2d),
				new ExpectedCellValue(sheetName, 43, 7, 1194d),
				new ExpectedCellValue(sheetName, 44, 7, 6d),
				new ExpectedCellValue(sheetName, 46, 7, 831.5),
				new ExpectedCellValue(sheetName, 47, 7, 2d),
				new ExpectedCellValue(sheetName, 48, 7, 1194d),
				new ExpectedCellValue(sheetName, 49, 7, 24d),

				new ExpectedCellValue(sheetName, 37, 8, "San Francisco"),
				new ExpectedCellValue(sheetName, 38, 8, "Car Rack"),
				new ExpectedCellValue(sheetName, 40, 8, 415.75),
				new ExpectedCellValue(sheetName, 41, 8, 1d),
				new ExpectedCellValue(sheetName, 48, 8, 415.75),
				new ExpectedCellValue(sheetName, 49, 8, 1d),

				new ExpectedCellValue(sheetName, 38, 9, "Sleeping Bag"),
				new ExpectedCellValue(sheetName, 43, 9, 99d),
				new ExpectedCellValue(sheetName, 44, 9, 1d),
				new ExpectedCellValue(sheetName, 48, 9, 99d),
				new ExpectedCellValue(sheetName, 49, 9, 1d),

				new ExpectedCellValue(sheetName, 37, 10, "San Francisco Total"),
				new ExpectedCellValue(sheetName, 40, 10, 415.75),
				new ExpectedCellValue(sheetName, 41, 10, 1d),
				new ExpectedCellValue(sheetName, 43, 10, 99d),
				new ExpectedCellValue(sheetName, 44, 10, 1d),
				new ExpectedCellValue(sheetName, 48, 10, 415.75),
				new ExpectedCellValue(sheetName, 49, 10, 1d),

				new ExpectedCellValue(sheetName, 37, 11, "Grand Total"),
				new ExpectedCellValue(sheetName, 40, 11, 831.5),
				new ExpectedCellValue(sheetName, 41, 11, 4d),
				new ExpectedCellValue(sheetName, 43, 11, 1194d),
				new ExpectedCellValue(sheetName, 44, 11, 6d),
				new ExpectedCellValue(sheetName, 46, 11, 831.5),
				new ExpectedCellValue(sheetName, 47, 11, 2d),
				new ExpectedCellValue(sheetName, 48, 11, 1194d),
				new ExpectedCellValue(sheetName, 49, 11, 48d)
			});
		}
		#endregion
	}
}