using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml
{
    internal class FormulaManager : IFormulaManager
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
        public string UpdateFormulaReferences(string formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, string currentSheet, string modifiedSheet, bool setFixed = false)
        {
            var d = new Dictionary<string, object>();
            try
            {
                var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
                var tokens = sct.Tokenize(formula);
                String f = "";
                foreach (var t in tokens)
                {
                    if (t.TokenType == TokenType.ExcelAddress)
                    {
                        var a = new ExcelAddressBase(t.Value);
                        var referencesModifiedWorksheet = (string.IsNullOrEmpty(a._ws) && currentSheet.Equals(modifiedSheet, StringComparison.CurrentCultureIgnoreCase)) || modifiedSheet.Equals(a._ws, StringComparison.CurrentCultureIgnoreCase);

                        if (!setFixed && (!string.IsNullOrEmpty(a._wb) || !referencesModifiedWorksheet))
                        {
                            // This address is in a different worksheet or workbook; no update is required.
                            f += a.Address;
                            continue;
                        }
                        // Persist fully-qualified worksheet references.
                        if (!string.IsNullOrEmpty(a._ws))
                        {
                            f += $"'{a._ws}'!";
                        }
                        if (rowIncrement > 0)
                        {
                            a = a.AddRow(afterRow, rowIncrement, setFixed);
                        }
                        else if (rowIncrement < 0)
                        {
                            a = a.DeleteRow(afterRow, -rowIncrement, setFixed);
                        }
                        if (colIncrement > 0)
                        {
                            a = a.AddColumn(afterColumn, colIncrement, setFixed);
                        }
                        else if (colIncrement < 0)
                        {
                            a = a.DeleteColumn(afterColumn, -colIncrement, setFixed);
                        }
                        if (a == null || !a.IsValidRowCol())
                        {
                            f += "#REF!";
                        }
                        else
                        {
                            // If the address was not shifted, then a.Address will still have the sheet name.
                            var address = a.Address.Split('!');
                            if (address.Length > 1)
                                f += address[1];
                            else
                                f += a.Address;
                        }
                    }
                    else
                    {
                        if (t.TokenType == TokenType.StringContent)
                            f += t.Value.Replace("\"", "\"\"");
                        else
                            f += t.Value;
                    }
                }
                return f;
            }
            catch //Invalid formula, skip updating addresses
            {
                return formula;
            }
        }

        /// <summary>
        /// Updates all the references to a renamed sheet in a formula.
        /// </summary>
        /// <param name="formula">The formula to updated.</param>
        /// <param name="oldSheetName">The old sheet name.</param>
        /// <param name="newSheetName">The new sheet name.</param>
        /// <returns>The formula with all cross-sheet references updated.</returns>
        public string UpdateFormulaSheetReferences(string formula, string oldSheetName, string newSheetName)
        {
            if (string.IsNullOrEmpty(oldSheetName))
                throw new ArgumentNullException(nameof(oldSheetName));
            if (string.IsNullOrEmpty(newSheetName))
                throw new ArgumentNullException(nameof(newSheetName));
            var d = new Dictionary<string, object>();
            try
            {
                var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
                var tokens = sct.Tokenize(formula);
                String f = "";
                foreach (var t in tokens)
                {
                    if (t.TokenType == TokenType.ExcelAddress)
                    {
                        var a = new ExcelAddressBase(t.Value);
                        if (a == null || !a.IsValidRowCol())
                        {
                            f += "#REF!";
                        }
                        else
                        {
                            a.ChangeWorksheet(oldSheetName, newSheetName);
                            f += a.Address;
                        }
                    }
                    else
                    {
                        f += t.Value;
                    }
                }
                return f;
            }
            catch //Invalid formula, skip updating addresses
            {
                return formula;
            }
        }
    }
    
    public interface IFormulaManager
    {
        string UpdateFormulaReferences(string formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, string currentSheet, string modifiedSheet, bool setFixed = false);

        string UpdateFormulaSheetReferences(string formula, string oldSheetName, string newSheetName);
    }
}
