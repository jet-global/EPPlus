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

using System;
using System.Collections.Generic;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml
{
	public static class CalculationExtension
	{
		#region Public Static Methods
		/// <summary>
		/// Recalculate this <see cref="ExcelWorkbook"/>.
		/// </summary>
		/// <param name="workbook">The workbook to be calculated.</param>
		public static void Calculate(this ExcelWorkbook workbook)
		{
			Calculate(workbook, new ExcelCalculationOption() { AllowCircularReferences = false });
		}

		/// <summary>
		/// Recalculate this <paramref name="workbook"/> with the specified <paramref name="options"/>.
		/// </summary>
		/// <param name="workbook">The workbook to calculate.</param>
		/// <param name="options">The calculation options (whether or not circular references are allowed). </param>
		public static void Calculate(this ExcelWorkbook workbook, ExcelCalculationOption options)
		{
			Init(workbook);

			var dc = DependencyChainFactory.Create(workbook, options);
			workbook.FormulaParser.InitNewCalc();
			if (workbook.FormulaParser.Logger != null)
			{
				var msg = string.Format("Starting... number of cells to parse: {0}", dc.List.Count);
				workbook.FormulaParser.Logger.Log(msg);
			}
			CalcChain(workbook, workbook.FormulaParser, dc);
		}

		/// <summary>
		/// Recalculate this <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="worksheet">The worksheet to recalculate.</param>
		public static void Calculate(this ExcelWorksheet worksheet)
		{
			Calculate(worksheet, new ExcelCalculationOption());
		}

		/// <summary>
		/// Recalculate this <paramref name="worksheet"/> with the specified <paramref name="options"/>.
		/// </summary>
		/// <param name="worksheet">The worksheet to calculate.</param>
		/// <param name="options">The calculation options (whether or not circular references are allowed). </param>
		public static void Calculate(this ExcelWorksheet worksheet, ExcelCalculationOption options)
		{
			Init(worksheet.Workbook);
			var dc = DependencyChainFactory.Create(worksheet, options);
			var parser = worksheet.Workbook.FormulaParser;
			parser.InitNewCalc();
			if (parser.Logger != null)
			{
				var msg = string.Format("Starting... number of cells to parse: {0}", dc.List.Count);
				parser.Logger.Log(msg);
			}
			CalcChain(worksheet.Workbook, parser, dc);
		}

		/// <summary>
		/// Recalculate this <paramref name="range"/>.
		/// </summary>
		/// <param name="range">The range to be calculated.</param>
		public static void Calculate(this ExcelRangeBase range)
		{
			Calculate(range, new ExcelCalculationOption());
		}

		/// <summary>
		/// Recalculate this <paramref name="range"/> with the specified <paramref name="options"/>.
		/// </summary>
		/// <param name="range">The range to be calculated.</param>
		/// <param name="options">Settings for this calculation.</param>
		public static void Calculate(this ExcelRangeBase range, ExcelCalculationOption options)
		{
			Init(range.myWorkbook);
			var parser = range.myWorkbook.FormulaParser;
			parser.InitNewCalc();
			var dc = DependencyChainFactory.Create(range, options);
			CalcChain(range.myWorkbook, parser, dc);
		}

		/// <summary>
		/// Calculate a specific <paramref name="formula"/> in the context of the specified <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="worksheet">The worksheet whose context should be used during the calculation for references that do not specify a sheet.</param>
		/// <param name="formula">The formula to be calculated.</param>
		/// <returns>The result of the calculation.</returns>
		public static object Calculate(this ExcelWorksheet worksheet, string formula)
		{
			return Calculate(worksheet, formula, new ExcelCalculationOption());
		}

		/// <summary>
		/// Calculate a specific <paramref name="formula"/> in the context of the specified <paramref name="worksheet"/> with the specified <paramref name="options"/>.
		/// </summary>
		/// <param name="worksheet">The worksheet whose context should be used during the calculation for references that do not specify a sheet.</param>
		/// <param name="formula">The formula to be calculated.</param>
		/// <param name="options">The options for this calculation. At the moment, this does nothing.</param>
		/// <returns>The result of the calculation.</returns>
		public static object Calculate(this ExcelWorksheet worksheet, string formula, ExcelCalculationOption options)
		{
			try
			{
				worksheet.CheckSheetType();
				if (string.IsNullOrEmpty(formula.Trim())) return null;
				Init(worksheet.Workbook);
				var parser = worksheet.Workbook.FormulaParser;
				parser.InitNewCalc();
				if (formula[0] == '=') formula = formula.Substring(1); //Remove any starting equal sign
				var dc = DependencyChainFactory.Create(worksheet, formula, options);
				var f = dc.List[0];
				dc.CalcOrder.RemoveAt(dc.CalcOrder.Count - 1);

				CalcChain(worksheet.Workbook, parser, dc);

				return parser.ParseCell(f.Tokens, worksheet.Name, -1, -1);
			}
			catch (Exception ex)
			{
				return new ExcelErrorValueException(ex.Message, ExcelErrorValue.Create(eErrorType.Value));
			}
		}
		#endregion

		#region Private Static Methods
		private static void CalcChain(ExcelWorkbook wb, FormulaParser parser, DependencyChain dc)
		{
			var debug = parser.Logger != null;
			foreach (var ix in dc.CalcOrder)
			{
				var item = dc.List[ix];
				try
				{
					var ws = wb.Worksheets.GetBySheetID(item.SheetID);
					var v = parser.ParseCell(item.Tokens, ws == null ? "" : ws.Name, item.Row, item.Column);
					SetValue(wb, item, v);
					if (debug)
					{
						parser.Logger.LogCellCounted();
					}
					Thread.Sleep(0);
				}
				catch (Exception ex) when ((ex is OperationCanceledException) == false)
				{
					parser.Logger?.Log(ex);
					var error = ExcelErrorValue.Parse(ExcelErrorValue.Values.Value);
					SetValue(wb, item, error);
				}
			}
		}

		private static void Init(ExcelWorkbook workbook)
		{
			workbook.FormulaTokens = new CellStore<List<Token>>(); ;
			foreach (var ws in workbook.Worksheets)
			{
				if (!(ws is ExcelChartsheet))
				{
					if (ws._formulaTokens != null)
					{
						ws._formulaTokens.Dispose();
					}
					ws._formulaTokens = new CellStore<List<Token>>();
				}
			}
		}

		private static void SetValue(ExcelWorkbook workbook, FormulaCell item, object v)
		{
			if (item.Column == 0)
			{
				if (item.SheetID == -1)
				{
					workbook.Names[item.Row].NameValue = v;
				}
				else
				{
					var sh = workbook.Worksheets.GetBySheetID(item.SheetID);
					sh.Names[item.Row].NameValue = v;
				}
			}
			else
			{
				var sheet = workbook.Worksheets.GetBySheetID(item.SheetID);
				sheet.SetValueInner(item.Row, item.Column, v);
			}
		}
		#endregion
	}
}
