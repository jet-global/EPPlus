using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	internal class PercentOfTotalCalculator : ShowDataAsCalculatorBase
	{
		#region Constructors
		public PercentOfTotalCalculator(ExcelPivotTable pivotTable,
			int dataFieldCollectionIndex, 
			PivotCellBackingData[,] backingDatas, 
			PivotCellBackingData[] grandGrandTotalValues,
			List<PivotCellBackingData> rowGrandTotalsValuesLists, 
			List<PivotCellBackingData> columnGrandTotalsValuesLists,
			int dataRow, int dataColumn)
			: base(pivotTable, backingDatas, grandGrandTotalValues, rowGrandTotalsValuesLists,
					columnGrandTotalsValuesLists, dataFieldCollectionIndex, dataRow, dataColumn)
		{ }
		#endregion

		#region ShowDataAsCalculatorBase Overrides
		public override  object CalculateBodyValue()
		{
			var cellBackingData = base.GetBodyBackingData();
			if (cellBackingData == null)
				return null;
			else if (cellBackingData.Result == null)
				return 0;
			else
			{
				double denominator = (double)base.GrandGrandTotalValues[base.DataFieldCollectionIndex].Result;
				return (double)cellBackingData.Result / denominator;
			}
		}

		public override object CalculateGrandTotalValue(PivotCellBackingData grandTotalBackingData, PivotCellBackingData[] columnGrandGrandTotalValues, bool isRowTotal)
		{
			if (columnGrandGrandTotalValues.Length > grandTotalBackingData.DataFieldCollectionIndex)
			{
				if (grandTotalBackingData?.Result == null)
					return 0;
				else
				{
					double grandGrandTotalValue = (double)columnGrandGrandTotalValues[grandTotalBackingData.DataFieldCollectionIndex].Result;
					return (double)grandTotalBackingData.Result / grandGrandTotalValue;
				}
			}
			return 1;
		}

		public override object CalculateGrandGrandTotalValue(PivotCellBackingData backingData)
		{
			return 1;
		}
		#endregion
	}
}
