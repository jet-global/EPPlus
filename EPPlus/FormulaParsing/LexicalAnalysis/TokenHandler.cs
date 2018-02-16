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
 * Mats Alm   		                Added       		        2015-12-28
 *******************************************************************************/
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
	/// <summary>
	/// A class to create and enumerate tokens. 
	/// </summary>
	public class TokenHandler : ITokenIndexProvider
	{
		#region Class Variables
		private readonly TokenizerContext _context;
		private readonly ITokenSeparatorProvider _tokenProvider;
		private readonly ITokenFactory _tokenFactory;
		private int _tokenIndex = -1;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the worksheet name used for tokenization.
		/// </summary>
		public string Worksheet { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="TokenHandler"/> object.
		/// </summary>
		/// <param name="context">The context within which to create and enumerate tokens.</param>
		/// <param name="tokenFactory">The token factory to use.</param>
		/// <param name="tokenProvider">The token provider to use.</param>
		public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
		{
			_context = context;
			_tokenFactory = tokenFactory;
			_tokenProvider = tokenProvider;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Determines if there are more tokens to handle.
		/// </summary>
		/// <returns>True if there are more tokens to handle, false otherwise.</returns>
		public bool HasMore()
		{
			return _tokenIndex < (_context.FormulaChars.Length - 1);
		}

		/// <summary>
		/// Handles the next token.
		/// </summary>
		public void Next()
		{
			_tokenIndex++;
			this.Handle();
		}
		#endregion

		#region ITokenIndexProvider Implentation
		/// <summary>
		/// Gets the current token index.
		/// </summary>
		int ITokenIndexProvider.Index
		{
			get { return _tokenIndex; }
		}

		/// <summary>
		/// Increments the current token index.
		/// </summary>
		void ITokenIndexProvider.MoveIndexPointerForward()
		{
			_tokenIndex++;
		}
		#endregion

		#region Private Methods
		private void Handle()
		{
			var c = _context.FormulaChars[_tokenIndex];
			if (this.CharIsTokenSeparator(c, out var tokenSeparator))
			{
				if (TokenSeparatorHandler.Handle(c, tokenSeparator, _context, this))
					return;

				if (_context.CurrentTokenHasValue)
				{
					if (Regex.IsMatch(_context.CurrentToken, "^\"*$"))
						_context.AddToken(_tokenFactory.Create(_context.CurrentToken, TokenType.StringContent));
					else
						_context.AddToken(CreateToken(_context, this.Worksheet));

					//If the a next token is an opening parantheses and the previous token is interpeted as an address or name, then the currenct token is a function
					if (tokenSeparator.TokenType == TokenType.OpeningParenthesis && (_context.LastToken.TokenType == TokenType.ExcelAddress || _context.LastToken.TokenType == TokenType.NameValue))
						_context.LastToken.TokenType = TokenType.Function;
				}
				if (tokenSeparator.Value == "-")
				{
					if (TokenIsNegator(_context))
					{
						_context.AddToken(new Token("-", TokenType.Negator));
						return;
					}
				}
				_context.AddToken(tokenSeparator);
				_context.NewToken();
				return;
			}
			else if (c == '#' && !_context.CurrentTokenHasValue && !_context.IsInSheetName && !_context.IsInString)
			{
				do
				{
					_context.AppendToCurrentToken(c);
					_tokenIndex++;
					if (_tokenIndex > _context.FormulaChars.Length - 1)
						break;
					c = _context.FormulaChars[_tokenIndex];
				}
				while (!ExcelErrorValue.Values.TryGetErrorType(_context.CurrentToken, out _));
				_tokenIndex--;
				if (this.CharIsTokenSeparator(c, out _) || _tokenIndex == _context.FormulaChars.Length - 1)
				{
					ExcelErrorValue.Values.TryGetErrorType(_context.CurrentToken, out eErrorType errorType);
					switch (errorType)
					{
						case eErrorType.Div0:
							_context.AddToken(new Token(_context.CurrentToken, TokenType.DivideByZeroError));
							break;
						case eErrorType.NA:
							_context.AddToken(new Token(_context.CurrentToken, TokenType.NotApplicableError));
							break;
						case eErrorType.Name:
							_context.AddToken(new Token(_context.CurrentToken, TokenType.NameError));
							break;
						case eErrorType.Null:
							_context.AddToken(new Token(_context.CurrentToken, TokenType.Null));
							break;
						case eErrorType.Num:
							_context.AddToken(new Token(_context.CurrentToken, TokenType.NumericError));
							break;
						case eErrorType.Ref:
							// Let #REF! errors be handled in the TokenFactory.
							return;
						case eErrorType.Value:
							_context.AddToken(new Token(_context.CurrentToken, TokenType.ValueDataTypeError));
							break;
					}
					_context.NewToken();
				}
				return;
			}
			_context.AppendToCurrentToken(c);
		}

		private bool CharIsTokenSeparator(char c, out Token token)
		{
			var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
			token = result ? token = _tokenProvider.Tokens[c.ToString()] : null;
			return result;
		}

		private static bool TokenIsNegator(TokenizerContext context)
		{
			return TokenIsNegator(context.LastToken);
		}
		private static bool TokenIsNegator(Token t)
		{
			return t == null
							||
							t.TokenType == TokenType.Operator
							||
							t.TokenType == TokenType.OpeningParenthesis
							||
							t.TokenType == TokenType.Comma
							||
							t.TokenType == TokenType.SemiColon
							||
							t.TokenType == TokenType.OpeningEnumerable;
		}

		private Token CreateToken(TokenizerContext context, string worksheet)
		{
			if (context.CurrentToken == "-")
			{
				if (context.LastToken == null && context.LastToken.TokenType == TokenType.Operator)
					return new Token("-", TokenType.Negator);
			}
			return _tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
		}
		#endregion
	}
}
