using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	internal static class ShowDataAsFactory
	{
		#region Public Static Methods
		public static ShowDataAsCalculatorBase GetShowDataAsCalculator(
			ShowDataAs showDataAs, 
			ExcelPivotTable pivotTable,
			int dataFieldCollectionIndex,
			PivotCellBackingData[,] backingDatas = null, 
			PivotCellBackingData[] grandGrandTotalValues = null,
			List<PivotCellBackingData> rowGrandTotalsValuesLists = null, 
			List<PivotCellBackingData> columnGrandTotalsValuesLists = null,
			 int dataRow = -1, int dataColumn = -1)
		{
			switch (showDataAs)
			{
				case ShowDataAs.NoCalculation:
					return new NoCalculationCalcutor(pivotTable,
						dataFieldCollectionIndex,
						backingDatas,
						grandGrandTotalValues,
						rowGrandTotalsValuesLists,
						columnGrandTotalsValuesLists,
						dataRow, dataColumn);
				case ShowDataAs.PercentOfTotal:
					return new PercentOfTotalCalculator(
						pivotTable,
						dataFieldCollectionIndex, 
						backingDatas, 
						grandGrandTotalValues, 
						rowGrandTotalsValuesLists, 
						columnGrandTotalsValuesLists,
						dataRow, dataColumn);
				case ShowDataAs.PercentOfRow:
					return new PercentOfRowCalculator(
						pivotTable,
						dataFieldCollectionIndex,
						backingDatas,
						grandGrandTotalValues,
						rowGrandTotalsValuesLists,
						columnGrandTotalsValuesLists,
						dataRow, dataColumn);
				case ShowDataAs.PercentOfCol:
					return new PercentOfColCalculator(
						pivotTable,
						dataFieldCollectionIndex,
						backingDatas,
						grandGrandTotalValues,
						rowGrandTotalsValuesLists,
						columnGrandTotalsValuesLists,
						dataRow, dataColumn);
				case ShowDataAs.Percent:
					return new PercentOfCalculator(
						pivotTable,
						dataFieldCollectionIndex,
						backingDatas,
						grandGrandTotalValues,
						rowGrandTotalsValuesLists,
						columnGrandTotalsValuesLists,
						dataRow, dataColumn);
				default:
					// TODO: Implement the rest of these settings. See user story 11453.
					throw new InvalidOperationException($"Unsupported dataField ShowDataAs setting '{showDataAs}'");
			}
		}
		#endregion
	}
}
