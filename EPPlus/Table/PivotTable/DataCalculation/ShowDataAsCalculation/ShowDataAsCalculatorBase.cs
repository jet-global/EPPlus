using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	internal abstract class ShowDataAsCalculatorBase
	{
		#region Properties
		protected ExcelPivotTable PivotTable { get; }
		protected PivotCellBackingData[,] BackingDatas { get; }
		protected PivotCellBackingData[] GrandGrandTotalValues { get; }
		protected List<PivotCellBackingData> ColumnGrandTotalsValuesLists { get; }
		protected List<PivotCellBackingData> RowGrandTotalsValuesLists { get; }
		protected int DataFieldCollectionIndex { get; }
		protected int DataRow { get; }
		protected int DataColumn { get; }
		protected int SheetRow { get; }
		protected int SheetColumn { get; }
		#endregion

		#region Constructors
		public ShowDataAsCalculatorBase(ExcelPivotTable pivotTable,
			PivotCellBackingData[,] backingDatas, 
			PivotCellBackingData[] grandGrandTotalValues,
			List<PivotCellBackingData> rowGrandTotalsValuesLists, 
			List<PivotCellBackingData> columnGrandTotalsValuesLists,
			int dataFieldCollectionIndex, int dataRow, int dataColumn)
		{
			this.PivotTable = pivotTable;
			this.BackingDatas = backingDatas;
			this.GrandGrandTotalValues = grandGrandTotalValues;
			this.RowGrandTotalsValuesLists = rowGrandTotalsValuesLists;
			this.ColumnGrandTotalsValuesLists = columnGrandTotalsValuesLists;
			this.DataFieldCollectionIndex = dataFieldCollectionIndex;
			this.DataRow = dataRow;
			this.DataColumn = dataColumn;
			this.SheetRow = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow + this.DataRow;
			this.SheetColumn = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol + this.DataColumn;
		}
		#endregion

		#region Abstract Methods
		public abstract object CalculateBodyValue();

		public abstract object CalculateGrandTotalValue(PivotCellBackingData grandTotalBackingData, PivotCellBackingData[] columnGrandGrandTotalValues, bool isRowTotal);

		public abstract object CalculateGrandGrandTotalValue(PivotCellBackingData backingData);
		#endregion

		#region Protected Methods
		protected PivotCellBackingData GetBodyBackingData()
		{
			return this.BackingDatas[this.DataRow, this.DataColumn];
		}
		#endregion
	}
}
