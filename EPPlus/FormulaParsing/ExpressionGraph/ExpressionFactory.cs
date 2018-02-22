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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Generates expressions based off of tokens.
	/// </summary>
	public class ExpressionFactory : IExpressionFactory
	{
		#region Class Variables
		private readonly ExcelDataProvider _excelDataProvider;
		private readonly ParsingContext _parsingContext;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExpressionFactory"/>.
		/// </summary>
		/// <param name="excelDataProvider">An <see cref="ExcelDataProvider"/> to retrieve data from a workbook.</param>
		/// <param name="context">The <see cref="ParsingContext"/> for the factory.</param>
		public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
		{
			_excelDataProvider = excelDataProvider;
			_parsingContext = context;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Creates expressions from the given <paramref name="token"/>.
		/// </summary>
		/// <param name="token">The <see cref="Token"/> to generate an expression from.</param>
		/// <returns>The <see cref="Expression"/>.</returns>
		public Expression Create(Token token)
		{
			switch (token.TokenType)
			{
				case TokenType.Integer:
					return new IntegerExpression(token.Value, token.IsNegated);
				case TokenType.String:
					return new StringExpression(token.Value);
				case TokenType.Decimal:
					return new DecimalExpression(token.Value, token.IsNegated);
				case TokenType.Boolean:
					return new BooleanExpression(token.Value);
				case TokenType.ExcelAddress:
					return new ExcelAddressExpression(token.Value, _excelDataProvider, _parsingContext, token.IsNegated);
				case TokenType.StructuredReference:
					return new StructuredReferenceExpression(token.Value, _excelDataProvider, _parsingContext, token.IsNegated);
				case TokenType.InvalidReference:
					return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Ref));
				case TokenType.NumericError:
					return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Num));
				case TokenType.ValueDataTypeError:
					return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Value));
				case TokenType.Null:
					return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Null));
				case TokenType.NameValue:
					return new NamedValueExpression(token.Value, _parsingContext);
				default:
					return new StringExpression(token.Value);
			}
		}
		#endregion
	}
}
