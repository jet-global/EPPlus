/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Michelle Lau and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Extensions;
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

		private Dictionary<string, string> FieldNamesToSanitizedFieldNames { get; } = new Dictionary<string, string>();

		private Dictionary<string, string> SanitizedFieldNamesToFieldNames { get; } = new Dictionary<string, string>();

		private Dictionary<string, string> SanitizedFormulas { get; } = new Dictionary<string, string>();
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
			if (values == null || values.Count == 0 || values.All(v => v == null))
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
		/// Tokenizes the specified specified <paramref name="formula"/>. 
		/// <see cref="AddNames(HashSet{string})"/> must be called before this 
		/// method in order to parse the name values correctly.
		/// </summary>
		/// <param name="formula">The formula to tokenize.</param>
		/// <returns>A collection of the tokens in the formula.</returns>
		public List<Token> Tokenize(string formula)
		{
			if (string.IsNullOrEmpty(formula))
				return null;
			formula = this.SanitizeFormula(formula);
			var tokens = this.Package.Workbook.FormulaParser.Lexer.Tokenize(formula, this.TempWorksheet.Name);
			var nameValues = tokens.Where(t => t.TokenType == TokenType.NameValue).Select(t => t.Value);
			foreach (var formulaFieldName in nameValues)
			{
				string fieldName = formulaFieldName;
				if (this.SanitizedFieldNamesToFieldNames.ContainsKey(formulaFieldName))
					fieldName = this.SanitizedFieldNamesToFieldNames[formulaFieldName];
			}

			// Reconstruct the token list with the un-sanitized field names.
			var dirtyTokens = new List<Token>();
			foreach (var token in tokens)
			{
				if (this.SanitizedFieldNamesToFieldNames.ContainsKey(token.Value))
					dirtyTokens.Add(new Token(this.SanitizedFieldNamesToFieldNames[token.Value], TokenType.NameValue));
				else
					dirtyTokens.Add(token);
			}
			return dirtyTokens;
		}

		/// <summary>
		/// Adds the set of names to the temp worksheet to prepare for evaluating a calculated field.
		/// Names with illegal characters are mapped to new names.
		/// </summary>
		/// <param name="fieldNames">The set of field names to add.</param>
		public void AddNames(HashSet<string> fieldNames)
		{
			// Create named ranges for each one of the fields in the pivot table source data
			// so that the field references in formulas will tokenize as named ranges.
			var row = 1;
			foreach (var fieldName in fieldNames)
			{
				// Map field names with characters that are not valid for a named range to a GUID.
				var sanitizedName = fieldName;
				if (sanitizedName.IndexOfAny(ExcelNamedRange.IllegalCharacters) != -1)
				{
					sanitizedName = Guid.NewGuid().ToString("N");
					this.FieldNamesToSanitizedFieldNames.Add(fieldName, sanitizedName);
					this.SanitizedFieldNamesToFieldNames.Add(sanitizedName, fieldName);
				}
				var cell = this.TempWorksheet.Cells[row, TotalsFunctionHelper.NameValueColumn];
				if (!this.TempWorksheet.Names.ContainsKey(sanitizedName))
					this.TempWorksheet.Names.Add(sanitizedName, cell.FullAddressAbsolute);
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
			// Create a named range that calculates each field referenced by the formula.
			foreach (var nameToValues in namesToValues)
			{
				var excelName = nameToValues.Key;
				if (this.FieldNamesToSanitizedFieldNames.ContainsKey(excelName))
					excelName = this.FieldNamesToSanitizedFieldNames[excelName];
				// Update the formula of the named range with the name of the field 
				// to be the sum of the field values.
				this.TempWorksheet.Names[excelName].NameFormula = $"SUM({string.Join(",", nameToValues.Value)})";
			}
			formula = this.SanitizeFormula(formula);
			// Evaluate the formula. The fields that are referenced in it
		  // are now named ranges in this workbook, so they will resolve to the appropriate values.
			var result = this.TempWorksheet.Calculate(formula);
			return result;
		}
		
		public object CalculateCellTotal(ExcelPivotTableDataField dataField, PivotCellBackingData backingData,
		string rowTotalType = null, string columnTotalType = null)
		{
			if (backingData == null)
				return null;
			if (!string.IsNullOrEmpty(rowTotalType) && !rowTotalType.IsEquivalentTo("default"))
			{
				// Only calculate a value if the row and column functions match up, or if there is no column function specified.
				if (string.IsNullOrEmpty(columnTotalType) || rowTotalType.IsEquivalentTo(columnTotalType))
				{
					// Calculate the value with rowTotalType as the function.
					var function = ExcelPivotTableField.SubtotalFunctionTypeToDataFieldFunctionEnum[rowTotalType];
					return this.Calculate(function, backingData.GetBackingValues());
				}
				// No value for this cell.
				return null;
			}
			else if (!string.IsNullOrEmpty(columnTotalType) && !columnTotalType.IsEquivalentTo("default"))
			{
				// We already know that the row subtotal function type is either empty or default because of the previous condition.
				// Calculate the value with columnTotalType as the function.
				var function = ExcelPivotTableField.SubtotalFunctionTypeToDataFieldFunctionEnum[columnTotalType];
				return this.Calculate(function, backingData.GetBackingValues());
			}
			else if (string.IsNullOrEmpty(backingData.Formula))
				return this.Calculate(dataField.Function, backingData.GetBackingValues());
			else
				return this.EvaluateCalculatedFieldFormula(backingData.GetCalculatedCellBackingValues(), backingData.Formula);
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

		private string SanitizeFormula(string formula)
		{
			string sanitizedFormula = formula;
			if (this.SanitizedFormulas.ContainsKey(formula))
				return this.SanitizedFormulas[formula];
			foreach (var fieldName in this.FieldNamesToSanitizedFieldNames.Keys)
			{
				var quotedValue = $"'{fieldName}'";
				if (sanitizedFormula.Contains(quotedValue))
					sanitizedFormula = sanitizedFormula.Replace(quotedValue, this.FieldNamesToSanitizedFieldNames[fieldName]);
				else if (sanitizedFormula.Contains(fieldName))
					sanitizedFormula = sanitizedFormula.Replace(fieldName, this.FieldNamesToSanitizedFieldNames[fieldName]);
			}
			this.SanitizedFormulas.Add(formula, sanitizedFormula);
			return sanitizedFormula;
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
