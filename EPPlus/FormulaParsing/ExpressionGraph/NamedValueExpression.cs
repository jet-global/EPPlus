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
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	public class NamedValueExpression : AtomicExpression
	{
		#region Class Variables
		private readonly ParsingContext _parsingContext;
		#endregion

		#region Constructors
		/// <summary>
		/// Constructs a new <see cref="NamedValueExpression"/> instance.
		/// </summary>
		/// <param name="expression">The named range expression.</param>
		/// <param name="parsingContext">The parsing context of the named range expression.</param>
		public NamedValueExpression(string expression, ParsingContext parsingContext) : base(expression)
		{
			_parsingContext = parsingContext;
		}
		#endregion

		#region AtomicExpression Overrides
		/// <summary>
		/// Compiles the expression.
		/// </summary>
		/// <returns>A <see cref="CompileResult"/> result of the expression compilation.</returns>
		public override CompileResult Compile()
		{
			var c = this._parsingContext.Scopes.Current;
			var name = _parsingContext.ExcelDataProvider.GetName(c.Address.Worksheet, ExpressionString);
			if (name == null)
				return new CompileResult(eErrorType.Name);
			if (name.Value == null)
				return null;
			if (name.Value is ExcelDataProvider.IRangeInfo range)
			{
				if (range.IsMulti)
					return new CompileResult(range, DataType.Enumerable);
				else
				{
					if (range.IsEmpty && range.Address._fromRowFixed == true && range.Address._fromColFixed == true)
						return null;
					var address = range.Address;
					var column = address.Start.Column;
					var row = address.Start.Row;
					if (address._fromColFixed == false)
					{
						column = (column - 1 + c.Address.FromCol) % ExcelPackage.MaxColumns;
						if (column == 0)
							column = ExcelPackage.MaxColumns;
					}
					if (address._fromRowFixed == false)
					{
						row = (row - 1 + c.Address.FromRow) % ExcelPackage.MaxRows;
						if (row == 0)
							row = ExcelPackage.MaxRows;
					}
					var factory = new CompileResultFactory();
					return factory.Create(range.Worksheet.Cells[row, column].Value);
				}
			}
			else
			{
				var factory = new CompileResultFactory();
				return factory.Create(name.Value);
			}
		}
		#endregion
	}
}
