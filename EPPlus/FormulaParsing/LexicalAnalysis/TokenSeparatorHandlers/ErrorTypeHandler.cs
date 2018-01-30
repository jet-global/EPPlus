/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Jan Källman, Evan Schallerer, and others as noted in the source history.
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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
	/// <summary>
	/// Handles error type tokens.
	/// </summary>
	public class ErrorTypeHandler : SeparatorHandler
	{
		#region SeparatorHandler Overrides
		/// <summary>
		/// Creates an error type token if the specified character is a '#'.
		/// </summary>
		/// <param name="c">The character to handle.</param>
		/// <param name="tokenSeparator">The separator to handle.</param>
		/// <param name="context">The tokenization context.</param>
		/// <param name="tokenIndexProvider">The <see cref="ITokenIndexProvider"/>.</param>
		/// <returns>True if an error type token was created, false otherwise.</returns>
		public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
		{
			if (tokenSeparator.TokenType == TokenType.Hashtag && context.CurrentTokenHasValue == false)
			{
				do
				{
					context.AppendToCurrentToken(c);
					tokenIndexProvider.MoveIndexPointerForward();
					if (tokenIndexProvider.Index > context.FormulaChars.Length - 1)
						break;
					c = context.FormulaChars[tokenIndexProvider.Index];
				}
				while (!ExcelErrorValue.Values.StringIsErrorValue(context.CurrentToken));
				if (c != ':')
				{
					var errorType = ExcelErrorValue.Values.ToErrorType(context.CurrentToken);
					switch (errorType)
					{
						case eErrorType.Div0:
							context.AddToken(new Token(context.CurrentToken, TokenType.DivideByZeroError));
							break;
						case eErrorType.NA:
							context.AddToken(new Token(context.CurrentToken, TokenType.NotApplicableError));
							break;
						case eErrorType.Name:
							context.AddToken(new Token(context.CurrentToken, TokenType.NameError));
							break;
						case eErrorType.Null:
							context.AddToken(new Token(context.CurrentToken, TokenType.Null));
							break;
						case eErrorType.Num:
							context.AddToken(new Token(context.CurrentToken, TokenType.NumericError));
							break;
						case eErrorType.Ref:
							context.AddToken(new Token(context.CurrentToken, TokenType.InvalidReference));
							break;
						case eErrorType.Value:
							context.AddToken(new Token(context.CurrentToken, TokenType.ValueDataTypeError));
							break;
					}
					context.NewToken();
					return true;
				}
			}
			return false;
		}
		#endregion
	}
}
