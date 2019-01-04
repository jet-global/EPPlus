using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.DataCalculation;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// Implements the Excel GETPIVOTDATA function.
	/// </summary>
	public class GetPivotData : LookupFunction
	{
		#region LookupFunction Overrides
		public override List<int> LookupArgumentIndicies => new List<int> { 1 };
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the GETPIVOTDATA function with the specified arguments.
		/// </summary>
		/// <param name="arguments">The arguments of the function to evaluate.</param>
		/// <param name="context">The context that the function is to be evaluated in.</param>
		/// <returns>The result of function evaluation as a <see cref="CompileResult"/>.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			string fieldName = arguments.ElementAt(0).Value?.ToString();
			var pivotTableAddress = arguments.ElementAt(1).ValueAsRangeInfo?.Address;
			if (string.IsNullOrEmpty(fieldName) || pivotTableAddress == null)
				return new CompileResult(eErrorType.Ref);
			var pivotTable = context.ExcelDataProvider.GetPivotTable(pivotTableAddress);
			if (pivotTable == null)
				return new CompileResult(eErrorType.Ref);

			var fieldPairs = arguments.Skip(2).ToList();
			if (fieldPairs.Count % 2 == 1)
				return new CompileResult(eErrorType.Ref);
			// Convert arguments 2...n into field/value pairs.
			var fieldValueIndices = this.ResolveFieldValuePairs(fieldPairs, pivotTable.CacheDefinition);

			int fieldIndex = this.GetFieldIndex(fieldName, pivotTable);
			// The field/value pairs or field was not found in the pivot table.
			if (fieldValueIndices == null || fieldIndex == -1)
				return new CompileResult(eErrorType.Ref);

			using (var totalsCalculator = new TotalsFunctionHelper())
			{
				var pageFieldIndices = pivotTable.GetPageFieldIndices();
				var matchingValues = pivotTable.CacheDefinition.CacheRecords.FindMatchingValues(fieldValueIndices, null, pageFieldIndices, fieldIndex);
				var dataField = pivotTable.DataFields.FirstOrDefault(d => d.Index == fieldIndex);
				var subtotal = totalsCalculator.Calculate(dataField.Function, matchingValues);
				if (subtotal == null)
					return new CompileResult(eErrorType.Ref);
				return new CompileResult(subtotal, DataType.Decimal);
			}
		}
		#endregion

		#region Private Methods
		private List<Tuple<int, int>> ResolveFieldValuePairs(List<FunctionArgument> fieldValueArguments, ExcelPivotCacheDefinition cacheDefinition)
		{
			var indices = new List<Tuple<int, int>>();
			for (int i = 0; i + 1 < fieldValueArguments.Count; i += 2)
			{
				string fieldName = fieldValueArguments[i].Value.ToString();
				string value = fieldValueArguments[i + 1].Value.ToString();
				int fieldIndex = -1;
				CacheFieldNode cacheField = null;
				for (int j = 0; j < cacheDefinition.CacheFields.Count; j++)
				{
					var currentCacheField = cacheDefinition.CacheFields[j];
					if (currentCacheField.Name.IsEquivalentTo(fieldName))
					{
						fieldIndex = j;
						cacheField = currentCacheField;
						break;
					}
				}
				if (fieldIndex == -1)
					return null;

				int valueIndex = -1;
				for (int j = 0; j < cacheField.SharedItems.Count; j++)
				{
					if (cacheField.SharedItems[j].Value.IsEquivalentTo(value))
					{
						valueIndex = j;
						break;
					}
				}
				if (valueIndex == -1)
					return null;

				var indexPair = new Tuple<int, int>(fieldIndex, valueIndex);
				indices.Add(indexPair);
			}
			return indices;
		}

		private int GetFieldIndex(string fieldName, ExcelPivotTable pivotTable)
		{
			int foundIndex = -1;
			int i = 0;
			foreach (var cacheField in pivotTable.CacheDefinition.CacheFields)
			{
				if (cacheField.Name.IsEquivalentTo(fieldName))
				{
					foundIndex = i;
					break;
				}
				i++;
			}
			if (foundIndex == -1)
			{
				foreach (var dataField in pivotTable.DataFields)
				{
					if (dataField.Name.IsEquivalentTo(fieldName))
						return dataField.Index;
				}
			}
			return foundIndex;
		}
		#endregion
	}
}
