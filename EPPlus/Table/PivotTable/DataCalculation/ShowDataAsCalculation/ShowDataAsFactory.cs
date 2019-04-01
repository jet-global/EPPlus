using System;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Factory class to get the appropriate class to calculate ShowDataAs values.
	/// </summary>
	internal static class ShowDataAsFactory
	{
		#region Public Static Methods
		/// <summary>
		/// Factory method to get the appropriate class to calculate ShowDataAs values.
		/// </summary>
		/// <param name="showDataAs">The <see cref="ShowDataAs"/> type to get a calculator class for.</param>
		/// <param name="pivotTable">The pivot table that the calculator is calculating against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		/// <param name="totalsCalculator">The <see cref="TotalsFunctionHelper"/> to calculate totals.</param>
		/// <returns>The appropriate calculator class for the ShowDataAs value.</returns>
		public static ShowDataAsCalculatorBase GetShowDataAsCalculator(ShowDataAs showDataAs, ExcelPivotTable pivotTable, int dataFieldCollectionIndex, TotalsFunctionHelper totalsCalculator)
		{
			switch (showDataAs)
			{
				case ShowDataAs.NoCalculation:
					return new NoCalculationCalcutor(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentOfTotal:
					return new PercentOfTotalCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentOfRow:
					return new PercentOfRowCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentOfCol:
					return new PercentOfColCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.Percent:
					return new PercentOfCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentOfParentRow:
					return new PercentOfParentRowCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentOfParentCol:
					return new PercentOfParentColumnCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentOfParent:
					return new PercentOfParentCalculator(pivotTable, dataFieldCollectionIndex, totalsCalculator);
				case ShowDataAs.PercentDiff:
				case ShowDataAs.PercentOfRunningTotal:
				case ShowDataAs.RankAscending:
				case ShowDataAs.RankDescending:
				case ShowDataAs.RunTotal:
				case ShowDataAs.Index:
				case ShowDataAs.Difference:
					// TODO: Implement the rest of these settings. See user stories 11882..11890
					throw new InvalidOperationException($"Unsupported dataField ShowDataAs setting '{showDataAs}'");
				default:
					throw new InvalidOperationException($"Unknown dataField ShowDataAs setting '{showDataAs}'");
			}
		}
		#endregion
	}
}
