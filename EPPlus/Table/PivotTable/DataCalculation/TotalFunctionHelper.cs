using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculates totals for pivot table data fields.
	/// </summary>
	internal class TotalsFunctionHelper : IDisposable
	{
		#region Constants
		private const int DataColumn = 1;
		private const int NameValueColumn = 2;
		private const int CacheFieldFormulaColumn = 3;
		#endregion

		#region Properties
		private ExcelWorksheet TempWorksheet { get; }

		private ExcelPackage Package { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		public TotalsFunctionHelper()
		{
			this.Package = new ExcelPackage();
			this.TempWorksheet = this.Package.Workbook.Worksheets.Add("Sheet1");
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Applies a function specified by the <paramref name="dataFieldFunction"/>
		/// over the specified collection of <paramref name="values"/>.
		/// </summary>
		/// <param name="dataFieldFunction">The dataField function to be applied.</param>
		/// <param name="values">The values to apply the function to.</param>
		/// <returns>The result of the function.</returns>
		public object Calculate(DataFieldFunctions dataFieldFunction, List<object> values)
		{
			if (values == null || values.Count == 0)
				return null;
			// Write the values into a temp worksheet.
			int row = 1;
			for (int i = 0; i < values.Count; i++)
			{
				row = i + 1;
				this.TempWorksheet.Cells[row, TotalsFunctionHelper.DataColumn].Value = values[i];
			}
			var resultCell = this.TempWorksheet.Cells[row + 1, TotalsFunctionHelper.DataColumn];
			resultCell.Formula = $"={this.GetCorrespondingExcelFunction(dataFieldFunction)}(A1:A{row})";
			resultCell.Calculate();
			return resultCell.Value;
		}

		/// <summary>
		/// Tokenizes the specified specified <paramref name="formula"/> and returns
		/// the <see cref="TokenType.NameValue"/>s in the formula. <see cref="AddNames(IEnumerable{string})"/>
		/// must be called before this method in order to parse the name values correctly.
		/// </summary>
		/// <param name="formula">The formula to tokenize.</param>
		/// <returns>A collection of the <see cref="TokenType.NameValue"/>s in the formula.</returns>
		public IEnumerable<Token> GetTokenNameValues(string formula)
		{
			return this.Package.Workbook.FormulaParser.Lexer.Tokenize(formula, this.TempWorksheet.Name)
				.Where(t => t.TokenType == TokenType.NameValue);
		}

		/// <summary>
		/// Adds the collection of names to the temp worksheet to prepare for evaluating a calculated field.
		/// </summary>
		/// <param name="fieldNames">The field names to add.</param>
		public void AddNames(IEnumerable<string> fieldNames)
		{
			// Create named ranges for each one of the fields in the pivot table source data
			// so that the field references in formulas will tokenize as named ranges.
			var row = 1;
			foreach (var fieldName in fieldNames)
			{
				var cell = this.TempWorksheet.Cells[row, TotalsFunctionHelper.NameValueColumn];
				if (!this.TempWorksheet.Names.ContainsKey(fieldName))
					this.TempWorksheet.Names.Add(fieldName, cell.FullAddressAbsolute);
				row++;
			}
		}

		/// <summary>
		/// Evaluates the specified <paramref name="formula"/> with the values specified by the 
		/// <paramref name="namesToValues"/> dictionary.
		/// </summary>
		/// <param name="namesToValues">The dictionary of field names to the values they resolve to.</param>
		/// <param name="formula">The formula to evaluate.</param>
		/// <returns>The result of evaluating the formula.</returns>
		public object EvaluateCalculatedFieldFormula(Dictionary<string, List<object>> namesToValues, string formula)
		{
			if (string.IsNullOrEmpty(formula))
				return null;
			// Write in a row of values for each field name.
			int row = 1;
			foreach (var nameToValues in namesToValues)
			{
				string sumArgument = "0";
				if (nameToValues.Value.Any())
				{
					int column = TotalsFunctionHelper.NameValueColumn;
					foreach (var value in nameToValues.Value)
					{
						this.TempWorksheet.Cells[row, column].Value = value;
						column++;
					}
					sumArgument = new ExcelRange(this.TempWorksheet, row, TotalsFunctionHelper.NameValueColumn, row, column - 1).FullAddressAbsolute;
				}
				// Update the formula of the named range with the name of the field to be the sum of the 
				// field values.
				this.TempWorksheet.Names[nameToValues.Key].NameFormula = $"SUM({sumArgument})";
				row++;
			}
			// Evaluate the formula. The fields that are referenced in it
		  // are now named ranges in this workbook, so they will resolve to the appropriate values.
			var result = this.TempWorksheet.Calculate(formula);
			return result;
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
					return "VAR.P";
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
