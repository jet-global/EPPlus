using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	internal class NoCalculationCalcutor : ShowDataAsCalculatorBase
	{
		#region Constructors
		public NoCalculationCalcutor(ExcelPivotTable pivotTable,
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
		public override object CalculateBodyValue()
		{
			// If no ShowDataAs value is selected, the "For empty cells show: [missingCaption]" setting can be applied.
			var cellBackingData = base.GetBodyBackingData();
			return this.GetCellNoCalculationValue(cellBackingData);
		}

		public override object CalculateGrandTotalValue(PivotCellBackingData grandTotalBackingData, PivotCellBackingData[] columnGrandGrandTotalValues, bool isRowTotal)
		{
			return this.GetCellNoCalculationValue(grandTotalBackingData);
		}

		public override object CalculateGrandGrandTotalValue(PivotCellBackingData backingData)
		{
			return this.GetCellNoCalculationValue(backingData);
		}
		#endregion

		#region Private Methods
		private object GetCellNoCalculationValue(PivotCellBackingData cellBackingData)
		{
			// Non-null backing data indicates that this cell is eligible for a value.
			if (cellBackingData != null && cellBackingData.Result == null)
				return base.PivotTable.ShowMissing ? base.PivotTable.MissingCaption : "0";
			return cellBackingData?.Result;
		}
		#endregion
	}
}
