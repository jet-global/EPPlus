using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.Utils
{
	public static class AddressUtility
	{
		#region Public Static Methods
		public static string ParseEntireColumnSelections(string address)
		{
			string parsedAddress = address;
			var matches = Regex.Matches(address, "[A-Z]+:[A-Z]+");
			foreach (System.Text.RegularExpressions.Match match in matches)
			{
				AddRowNumbersToEntireColumnRange(ref parsedAddress, match.Value);
			}
			return parsedAddress;
		}

		/// <summary>
		/// Attempts to parse the <paramref name="formula"/> as an address, evaluating reference functions and
		/// nested named ranges as necessary.
		/// </summary>
		/// <param name="workbook">The <see cref="ExcelWorkbook"/>.</param>
		/// <param name="localSheet">The <see cref="ExcelWorksheet"/> the formular is on.</param>
		/// <param name="formula">The formula that is trying to be parsed.</param>
		/// <returns>The formula as an <see cref="ExcelRangeBase"/> if it is an address, null otherwise.</returns>
		public static ExcelRangeBase GetFormulaAsCellRange(ExcelWorkbook workbook, ExcelWorksheet localSheet, string formula)
		{
			var stringBuilder = new StringBuilder();
			var tokens = workbook.FormulaParser.Lexer.Tokenize(formula).ToList();
			for (int i = 0; i < tokens.Count; ++i)
			{
				var token = tokens[i];
				if (token.TokenType == TokenType.ExcelAddress ||
					token.TokenType == TokenType.InvalidReference ||
					token.TokenType == TokenType.Comma)
				{
					stringBuilder.Append(token.Value);
				}
				else if (token.TokenType == TokenType.OpeningParenthesis || token.TokenType == TokenType.ClosingParenthesis)
					continue;
				else if (token.TokenType == TokenType.Function)
				{
					if (AddressUtility.TryCalculateReferenceFunction(tokens, localSheet, i, out string address, out i))
						stringBuilder.Append(address);
					else
						return null;
				}
				else if (token.TokenType == TokenType.NameValue)
				{
					var address = AddressUtility.HandleNameValue(workbook, localSheet, token.Value);
					if (address == null)
						return null;
					stringBuilder.Append(address);
				}
				else
					return null;
			}
			var addressString = stringBuilder.ToString();
			ExcelRangeBase.SplitAddress(addressString, out _, out string worksheetName, out _);
			if (string.IsNullOrEmpty(worksheetName))
				throw new InvalidOperationException("References in named ranges must be fully-qualified with sheet names.");
			var worksheet = workbook.Worksheets[worksheetName];
			if (worksheet == null)
				throw new InvalidOperationException($"The worksheet '{worksheetName}' in the named range formula {formula} does not exist.");
			return new ExcelRangeBase(worksheet, addressString);
		}
		#endregion

		#region Private Static Methods
		private static string HandleNameValue(ExcelWorkbook workbook, ExcelWorksheet localSheet, string tokenValue)
		{
			if (localSheet != null && localSheet.Names.ContainsKey(tokenValue))
			{
				var address = localSheet.Names[tokenValue].GetFormulaAsCellRange();
				if (address == null)
					return null;
				return address;
			}
			else if (workbook.Names.ContainsKey(tokenValue))
			{
				var address = workbook.Names[tokenValue].GetFormulaAsCellRange();
				if (address == null)
					return null;
				return address;
			}
			else
			{
				foreach (var sheet in workbook.Worksheets)
				{
					if (sheet.Tables.TableNames.ContainsKey(tokenValue))
					{
						var table = sheet.Tables[tokenValue];
						// If the table has a total in the last row, update the address so that it does not include it.
						var address = table.ShowTotal ? new ExcelAddress(table.Address.Start.Row, table.Address.Start.Column, 
							table.Address.End.Row - 1, table.Address.End.Column) : table.Address;
						address.ChangeWorksheet(address.WorkSheet, sheet.Name);
						return address.ToString();
					}
				}
				return null;
			}
		}

		private static bool TryCalculateReferenceFunction(List<Token> tokens, ExcelWorksheet localSheet, int index, out string address, out int i)
		{
			address = null;
			i = index;
			int parenCount = 0;
			var formula = string.Empty;
			var token = tokens[index];

			if (token.Value.StartsWith(Offset.Name, StringComparison.InvariantCultureIgnoreCase))
				formula += OffsetAddress.Name;
			else if (token.Value.StartsWith(Indirect.Name, StringComparison.InvariantCultureIgnoreCase))
				formula += IndirectAddress.Name;
			else
				return false;

			for (i = index + 1; i < tokens.Count; ++i)
			{
				token = tokens[i];
				formula += token.Value;
				if (token.TokenType == TokenType.OpeningParenthesis)
					parenCount++;
				else if (token.TokenType == TokenType.ClosingParenthesis)
				{
					parenCount--;
					if (parenCount == 0)
						break;
				}
			}
			address = localSheet.Calculate(formula) as string;
			if (!string.IsNullOrEmpty(address) && ExcelAddressUtil.IsValidAddress(address))
				return true;
			else
			{
				address = null;
				return false;
			}
		}

		private static void AddRowNumbersToEntireColumnRange(ref string address, string range)
		{
			var parsedRange = string.Format("{0}{1}", range, ExcelPackage.MaxRows);
			var splitArr = parsedRange.Split(new char[] { ':' });
			address = address.Replace(range, string.Format("{0}1:{1}", splitArr[0], splitArr[1]));
		}
		#endregion
	}
}
