using System;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml
{
	/// <summary>
	/// Contains logic for parsing and updating references in formula strings.
	/// </summary>
	public interface IFormulaManager
	{
		/// <summary>
		/// Updates the Excel formula so that all the cellAddresses are incremented by the row and column increments
		/// if they fall after the afterRow and afterColumn.
		/// Supports inserting rows and columns into existing templates.
		/// </summary>
		/// <param name="formula">The Excel formula</param>
		/// <param name="rowIncrement">The amount to increment the cell reference by</param>
		/// <param name="colIncrement">The amount to increment the cell reference by</param>
		/// <param name="afterRow">Only change rows after this row</param>
		/// <param name="afterColumn">Only change columns after this column</param>
		/// <param name="currentSheet">The sheet that contains the formula currently being processed.</param>
		/// <param name="modifiedSheet">The sheet where cells are being inserted or deleted.</param>
		/// <param name="setFixed">Fixed address</param>
		/// <returns>The updated version of the <paramref name="formula"/>.</returns>
		string UpdateFormulaReferences(string formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, string currentSheet, string modifiedSheet, bool setFixed = false);

		/// <summary>
		/// Updates all the references to a renamed sheet in a formula.
		/// </summary>
		/// <param name="formula">The formula to updated.</param>
		/// <param name="oldSheetName">The old sheet name.</param>
		/// <param name="newSheetName">The new sheet name.</param>
		/// <returns>The formula with all cross-sheet references updated.</returns>
		string UpdateFormulaSheetReferences(string formula, string oldSheetName, string newSheetName);
	}

	internal class FormulaManager : IFormulaManager
	{
		/// <summary>
		/// Updates the Excel formula so that all the cellAddresses are incremented by the row and column increments
		/// if they fall after the afterRow and afterColumn.
		/// Supports inserting rows and columns into existing templates.
		/// </summary>
		/// <param name="originalFormula">The Excel formula</param>
		/// <param name="rowIncrement">The amount to increment the cell reference by</param>
		/// <param name="colIncrement">The amount to increment the cell reference by</param>
		/// <param name="afterRow">Only change rows after this row</param>
		/// <param name="afterColumn">Only change columns after this column</param>
		/// <param name="currentSheet">The sheet that contains the formula currently being processed.</param>
		/// <param name="modifiedSheet">The sheet where cells are being inserted or deleted.</param>
		/// <param name="setFixed">Fixed address</param>
		/// <returns>The updated version of the <paramref name="originalFormula"/>.</returns>
		public string UpdateFormulaReferences(string originalFormula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, string currentSheet, string modifiedSheet, bool setFixed = false)
		{
			try
			{
				var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
				var tokens = sct.Tokenize(originalFormula);
				String formula = "";
				foreach (var t in tokens)
				{
					if (t.TokenType == TokenType.ExcelAddress)
					{
						var address = new ExcelAddressBase(t.Value);
						var referencesModifiedWorksheet = (string.IsNullOrEmpty(address._ws) && currentSheet.Equals(modifiedSheet, StringComparison.CurrentCultureIgnoreCase)) || modifiedSheet.Equals(address._ws, StringComparison.CurrentCultureIgnoreCase);

						if (!setFixed && (!string.IsNullOrEmpty(address._wb) || !referencesModifiedWorksheet))
						{
							// This address is in a different worksheet or workbook; no update is required.
							formula += address.Address;
							continue;
						}
						// Persist fully-qualified worksheet references.
						if (!string.IsNullOrEmpty(address._ws))
						{
							formula += $"'{address._ws}'!";
						}
						if (rowIncrement > 0)
						{
							address = address?.AddRow(afterRow, rowIncrement, setFixed);
						}
						else if (rowIncrement < 0)
						{
							address = address?.DeleteRow(afterRow, -rowIncrement, setFixed);
						}
						if (colIncrement > 0)
						{
							address = address?.AddColumn(afterColumn, colIncrement, setFixed);
						}
						else if (colIncrement < 0)
						{
							address = address?.DeleteColumn(afterColumn, -colIncrement, setFixed);
						}
						if (address == null || !address.IsValidRowCol())
						{
							formula += "#REF!";
						}
						else
						{
							// If the address was not shifted, then a.Address will still have the sheet name.
							var splitAddress = address.Address.Split('!');
							if (splitAddress.Length > 1)
								formula += splitAddress[1];
							else
								formula += address.Address;
						}
					}
					else
					{
						if (t.TokenType == TokenType.StringContent)
							formula += t.Value.Replace("\"", "\"\"");
						else
							formula += t.Value;
					}
				}
				return formula;
			}
			catch //Invalid formula, skip updating addresses
			{
				return originalFormula;
			}
		}

		/// <summary>
		/// Updates all the references to a renamed sheet in a formula.
		/// </summary>
		/// <param name="originalFormula">The formula to updated.</param>
		/// <param name="oldSheetName">The old sheet name.</param>
		/// <param name="newSheetName">The new sheet name.</param>
		/// <returns>The formula with all cross-sheet references updated.</returns>
		public string UpdateFormulaSheetReferences(string originalFormula, string oldSheetName, string newSheetName)
		{
			if (string.IsNullOrEmpty(oldSheetName))
				throw new ArgumentNullException(nameof(oldSheetName));
			if (string.IsNullOrEmpty(newSheetName))
				throw new ArgumentNullException(nameof(newSheetName));
			try
			{
				var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
				var tokens = sct.Tokenize(originalFormula);
				String formula = "";
				foreach (var t in tokens)
				{
					if (t.TokenType == TokenType.ExcelAddress)
					{
						var address = new ExcelAddressBase(t.Value);
						if (address == null || !address.IsValidRowCol())
						{
							formula += "#REF!";
						}
						else
						{
							address.ChangeWorksheet(oldSheetName, newSheetName);
							formula += address.Address;
						}
					}
					else
					{
						formula += t.Value;
					}
				}
				return formula;
			}
			catch //Invalid formula, skip updating addresses
			{
				return originalFormula;
			}
		}
	}
}
