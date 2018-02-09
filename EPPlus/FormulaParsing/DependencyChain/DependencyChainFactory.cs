/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * Code change notes:
 *
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman                      Added                       2012-03-04
 *******************************************************************************/

using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
	/// <summary>
	/// A factory for generating dependency chains for use when executing a calculation.
	/// </summary>
	internal static class DependencyChainFactory
	{
		#region Internal Static Methods
		internal static DependencyChain Create(ExcelWorkbook wb, ExcelCalculationOption options)
		{
			var depChain = new DependencyChain();
			foreach (var ws in wb.Worksheets)
			{
				if (!(ws is ExcelChartsheet))
					GetChain(depChain, wb.FormulaParser.Lexer, ws.Cells, options);
			}
			return depChain;
		}

		internal static DependencyChain Create(ExcelWorksheet ws, ExcelCalculationOption options)
		{
			ws.CheckSheetType();
			var depChain = new DependencyChain();
			GetChain(depChain, ws.Workbook.FormulaParser.Lexer, ws.Cells, options);
			return depChain;
		}

		internal static DependencyChain Create(ExcelWorksheet ws, string Formula, ExcelCalculationOption options)
		{
			ws.CheckSheetType();
			var depChain = new DependencyChain();

			GetChain(depChain, ws.Workbook.FormulaParser.Lexer, ws, Formula, options);

			return depChain;
		}

		internal static DependencyChain Create(ExcelRangeBase range, ExcelCalculationOption options)
		{
			var depChain = new DependencyChain();

			GetChain(depChain, range.Worksheet.Workbook.FormulaParser.Lexer, range, options);

			return depChain;
		}

		#endregion

		#region Private Static Methods
		private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelWorksheet ws, string formula, ExcelCalculationOption options)
		{
			var f = new FormulaCell() { SheetID = ws.SheetID, Row = -1, Column = -1 };
			f.Formula = formula;
			if (!string.IsNullOrEmpty(f.Formula))
			{
				f.Tokens = lexer.Tokenize(f.Formula, ws.Name).ToList();
				depChain.Add(f);
				FollowChain(depChain, lexer, ws.Workbook, ws, f, options);
			}
		}

		private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelRangeBase Range, ExcelCalculationOption options)
		{
			var ws = Range.Worksheet;
			var fs = ws._formulas.GetEnumerator(Range.Start.Row, Range.Start.Column, Range.End.Row, Range.End.Column);
			while (fs.MoveNext())
			{
				if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
				var id = ExcelCellBase.GetCellID(ws.SheetID, fs.Row, fs.Column);
				if (!depChain.Index.ContainsKey(id))
				{
					var f = new FormulaCell() { SheetID = ws.SheetID, Row = fs.Row, Column = fs.Column };
					if (fs.Value is int)
					{
						f.Formula = ws._sharedFormulas[(int)fs.Value].GetFormula(fs.Row, fs.Column, ws.Name);
					}
					else
					{
						f.Formula = fs.Value.ToString();
					}
					if (!string.IsNullOrEmpty(f.Formula))
					{
						f.Tokens = lexer.Tokenize(f.Formula, Range.Worksheet.Name).ToList();
						ws._formulaTokens.SetValue(fs.Row, fs.Column, f.Tokens);
						depChain.Add(f);
						FollowChain(depChain, lexer, ws.Workbook, ws, f, options);
					}
				}
			}
		}

		/// <summary>
		/// This method follows the calculation chain to get the order of the calculation
		/// Goto (!) is used internally to prevent stackoverflow on extremly large dependency trees (that is, many recursive formulas).
		/// </summary>
		/// <param name="depChain">The dependency chain object</param>
		/// <param name="lexer">The formula tokenizer</param>
		/// <param name="wb">The workbook where the formula comes from</param>
		/// <param name="ws">The worksheet where the formula comes from</param>
		/// <param name="f">The cell function object</param>
		/// <param name="options">Calcultaiton options</param>
		private static void FollowChain(DependencyChain depChain, ILexer lexer, ExcelWorkbook wb, ExcelWorksheet ws, FormulaCell f, ExcelCalculationOption options)
		{
			Stack<FormulaCell> stack = new Stack<FormulaCell>();
			iterateToken:
			while (f.tokenIx < f.Tokens.Count)
			{
				var t = f.Tokens[f.tokenIx];
				if (t.TokenType == TokenType.ExcelAddress)
				{
					var adr = new ExcelFormulaAddress(t.Value);
					if (adr.IsTableAddress)
					{
						adr.SetRCFromTable(ws.Package, new ExcelAddress(f.Row, f.Column, f.Row, f.Column));
					}

					if (adr.WorkSheet == null && adr.Collide(new ExcelAddress(f.Row, f.Column, f.Row, f.Column)) != ExcelAddress.eAddressCollition.No && !options.AllowCircularReferences)
					{
						throw (new CircularReferenceException(string.Format("Circular Reference in cell {0}", ExcelAddress.GetAddress(f.Row, f.Column))));
					}

					if (adr._fromRow > 0 && adr._fromCol > 0)
					{
						if (string.IsNullOrEmpty(adr.WorkSheet))
						{
							if (f.ws == null)
							{
								f.ws = ws;
							}
							else if (f.ws.SheetID != f.SheetID)
							{
								f.ws = wb.Worksheets.GetBySheetID(f.SheetID);
							}
						}
						else
						{
							f.ws = wb.Worksheets[adr.WorkSheet];
						}

						if (f.ws != null)
						{
							f.iterator = f.ws._formulas.GetEnumerator(adr.Start.Row, adr.Start.Column, adr.End.Row, adr.End.Column);
							goto iterateCells;
						}
					}
				}
				else if (t.TokenType == TokenType.NameValue)
				{
					ExcelNamedRange name = null;
					var worksheet = f.ws ?? ws;
					// Worksheet-scoped named ranges take precedence over workbook-scoped named ranges.
					if (worksheet?.Names?.ContainsKey(t.Value) == true)
						name = worksheet.Names[t.Value];
					else if (wb.Names.ContainsKey(t.Value))
						name = wb.Names[t.Value];
					if (name != null)
					{
						var nameFormulaTokens = name.GetRelativeNameFormula(f.Row, f.Column)?.ToList();
						if (nameFormulaTokens.Count == 0 && !string.IsNullOrEmpty(name.NameFormula))
							nameFormulaTokens = name.Workbook.FormulaParser.Lexer.Tokenize(name.NameFormula)?.ToList();
						// Remove the current named range token and replace it with the named range's formula.
						f.Tokens.RemoveAt(f.tokenIx);
						f.Tokens.InsertRange(f.tokenIx, nameFormulaTokens);
						goto iterateToken;
					}
				}
				else if (t.TokenType == TokenType.Function && t.Value.IsEquivalentTo(Offset.Name))
				{
					var stringBuilder = new StringBuilder($"{OffsetAddress.Name}(");
					int offsetStartIndex = f.tokenIx;
					int parenCount = 1;
					for (f.tokenIx += 2; parenCount > 0 && f.tokenIx < f.Tokens.Count; f.tokenIx++)
					{
						var token = f.Tokens[f.tokenIx];
						stringBuilder.Append(token.Value);
						if (token.TokenType == TokenType.OpeningParenthesis)
							parenCount++;
						else if (token.TokenType == TokenType.ClosingParenthesis)
							parenCount--;
					}
					ExcelRange cell = ws.Cells[f.Row, f.Column];
					string originalFormula = cell.Formula;
					string addressOffsetFormula = stringBuilder.ToString();
					stringBuilder.Clear();
					for (int i = 0; i < f.Tokens.Count; i++)
					{
						if (i == offsetStartIndex)
							stringBuilder.Append(0);
						else if (i < offsetStartIndex || i >= f.tokenIx)
							stringBuilder.Append(f.Tokens[i].Value);
					}
					cell.Formula = stringBuilder.ToString();
					var offsetResult = ws.Calculate(addressOffsetFormula, f.Row, f.Column);
					cell.Formula = originalFormula;
					if (offsetResult is string resultString)
					{
						ExcelAddress adr = new ExcelAddress(resultString);
						var worksheet = string.IsNullOrEmpty(adr.WorkSheet) ? ws : wb.Worksheets[adr.WorkSheet];
						// Only complete the OFFSET's dependency chain if a valid existing address was successfully parsed.
						if (worksheet != null)
						{
							f.Tokens.RemoveRange(offsetStartIndex, f.tokenIx - offsetStartIndex);
							var offsetResultTokens = wb.FormulaParser.Lexer.Tokenize(adr.FullAddress);
							f.Tokens.InsertRange(offsetStartIndex, offsetResultTokens);
							f.ws = worksheet;
							f.iterator = f.ws._formulas.GetEnumerator(adr.Start.Row, adr.Start.Column, adr.End.Row, adr.End.Column);
							goto iterateCells;
						}
					}
				}
				f.tokenIx++;
			}
			depChain.CalcOrder.Add(f.Index);
			if (stack.Count > 0)
			{
				f = stack.Pop();
				goto iterateCells;
			}
			return;
			iterateCells:

			while (f.iterator != null && f.iterator.MoveNext())
			{
				var v = f.iterator.Value;
				if (v == null || v.ToString().Trim() == "") continue;
				var id = ExcelAddress.GetCellID(f.ws.SheetID, f.iterator.Row, f.iterator.Column);
				if (!depChain.Index.ContainsKey(id))
				{
					var rf = new FormulaCell() { SheetID = f.ws.SheetID, Row = f.iterator.Row, Column = f.iterator.Column };
					if (f.iterator.Value is int)
					{
						rf.Formula = f.ws._sharedFormulas[(int)v].GetFormula(f.iterator.Row, f.iterator.Column, ws.Name);
					}
					else
					{
						rf.Formula = v.ToString();
					}
					rf.ws = f.ws;
					rf.Tokens = lexer.Tokenize(rf.Formula, f.ws.Name).ToList();
					ws._formulaTokens.SetValue(rf.Row, rf.Column, rf.Tokens);
					depChain.Add(rf);
					stack.Push(f);
					f = rf;
					goto iterateToken;
				}
				else
				{
					if (stack.Count > 0)
					{
						//Check for circular references
						foreach (var par in stack)
						{
							if (ExcelAddress.GetCellID(par.ws.SheetID, par.iterator.Row, par.iterator.Column) == id 
								|| ExcelAddress.GetCellID(par.ws.SheetID, par.Row, par.Column) == id)
							{
								if (options.AllowCircularReferences == false)
								{
									throw (new CircularReferenceException(string.Format("Circular Reference in cell {0}!{1}", par.ws.Name, ExcelAddress.GetAddress(f.Row, f.Column))));
								}
								else
								{
									f = stack.Pop();
									goto iterateCells;
								}
							}
						}
					}
				}
			}
			f.tokenIx++;
			goto iterateToken;
		}
		#endregion
	}
}
