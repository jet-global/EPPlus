using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculates totals for pivot table data fields.
	/// </summary>
	internal class TotalsFunctionHelper : IDisposable
	{
		#region Properties
		private ExcelPivotTable PivotTable { get; }

		private ExcelWorksheet TempWorksheet { get; }

		private ExcelPackage Package { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		public TotalsFunctionHelper(ExcelPivotTable pivotTable)
		{
			this.PivotTable = pivotTable;

			// TODO: Consider creating a shadow sheet in this workbook instead.
			//		Be sure to evaluate risks of cross-contamination and the like.
			this.Package = new ExcelPackage();
			this.TempWorksheet = this.Package.Workbook.Worksheets.Add("Sheet1");
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Applies a function specified by the <paramref name="dataField"/>
		/// over the specified collection of <paramref name="values"/>.
		/// </summary>
		/// <param name="dataField">The dataField containing the function to be applied.</param>
		/// <param name="values">The values to apply the function to.</param>
		/// <returns>The result of the function.</returns>
		public object Calculate(ExcelPivotTableDataField dataField, List<object> values)
		{
			if (values.Count == 0)
				return new CompileResult(0, DataType.Decimal);
			// Write the values into a temp worksheet.
			int row = 1;
			for (int i = 0; i < values.Count; i++)
			{
				row = i + 1;
				this.TempWorksheet.Cells[row, 1].Value = values[i];
			}
			var resultCell = this.TempWorksheet.Cells[row + 1, 1];
			resultCell.Formula = $"={this.GetCorrespondingExcelFunction(dataField.Function)}(A1:A{row})";
			resultCell.Calculate();
			return resultCell.Value;
		}
		#endregion

		#region Private Methods
		private string GetCorrespondingExcelFunction(DataFieldFunctions dataFieldFunction)
		{
			switch (dataFieldFunction)
			{
				case DataFieldFunctions.Count:
					return "COUNTA";
				case DataFieldFunctions.CountNums:
					return "COUNT";
				case DataFieldFunctions.None: // No function specified, default to sum.
				case DataFieldFunctions.Sum:
					return "SUM";
				case DataFieldFunctions.Average:
					return "AVERAGE";
				case DataFieldFunctions.Max:
					return "MAX";
				case DataFieldFunctions.Min:
					return "MIN";
				case DataFieldFunctions.Product:
					return "PRODUCT";
				case DataFieldFunctions.StdDev:
					return "STDEV.S";
				case DataFieldFunctions.StdDevP:
					return "STDEV.P";
				case DataFieldFunctions.Var:
					return "VAR.S";
				case DataFieldFunctions.VarP:
					return "VAR.S";
				default:
					throw new InvalidOperationException($"Invalid data field function: {dataFieldFunction}.");
			}
		}
		#endregion

		#region IDisposable Overrides
		public void Dispose()
		{
			this.Package.Dispose();
		}
		#endregion
	}
}
