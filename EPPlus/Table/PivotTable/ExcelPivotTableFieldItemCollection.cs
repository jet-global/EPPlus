namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection class for <see cref="ExcelPivotTableFieldItem"/>.
	/// </summary>
	public class ExcelPivotTableFieldItemCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>
	{
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="table">The existing pivot table.</param>
		public ExcelPivotTableFieldItemCollection(ExcelPivotTable table) : base(table)
		{

		}
	}
}