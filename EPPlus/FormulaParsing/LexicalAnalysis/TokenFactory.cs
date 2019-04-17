﻿/*******************************************************************************
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
 * Jan Källman                      Replaced Adress validate    2013-03-01
 * *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
	/// <summary>
	/// A factory class used to create tokens.
	/// </summary>
	public class TokenFactory : ITokenFactory
	{
		#region Class Variables
		private readonly ITokenSeparatorProvider _tokenSeparatorProvider;
		private readonly IFunctionNameProvider _functionNameProvider;
		private readonly INameValueProvider _nameValueProvider;
		#endregion

		#region Constructors
		/// <summary>
		/// Constructs a new <see cref="TokenFactory"/> object.
		/// </summary>
		/// <param name="functionNameProvider">The function name provider to use.</param>
		/// <param name="nameValueProvider">The name value provider to use.</param>
		public TokenFactory(IFunctionNameProvider functionNameProvider, INameValueProvider nameValueProvider)
			 : this(new TokenSeparatorProvider(), nameValueProvider, functionNameProvider)
		{
		}

		/// <summary>
		/// Constructs a new <see cref="TokenFactory"/> object.
		/// </summary>
		/// <param name="tokenSeparatorProvider">The token separator provider to use.</param>
		/// <param name="functionNameProvider">The function name provider to use.</param>
		/// <param name="nameValueProvider">The name value provider to use.</param>
		public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider, INameValueProvider nameValueProvider, IFunctionNameProvider functionNameProvider)
		{
			_tokenSeparatorProvider = tokenSeparatorProvider;
			_functionNameProvider = functionNameProvider;
			_nameValueProvider = nameValueProvider;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Create a new token.
		/// </summary>
		/// <param name="tokens">Existing tokens.</param>
		/// <param name="token">The token to create.</param>
		/// <returns>The created token.</returns>
		public Token Create(IEnumerable<Token> tokens, string token)
		{
			return Create(tokens, token, null);
		}

		/// <summary>
		/// Create a new token.
		/// </summary>
		/// <param name="tokens">Existing tokens.</param>
		/// <param name="token">The token to create.</param>
		/// <param name="worksheet">The worksheet name to use to create the token.</param>
		/// <returns>The created token.</returns>
		public Token Create(IEnumerable<Token> tokens, string token, string worksheet)
		{
			if (_tokenSeparatorProvider.Tokens.TryGetValue(token, out var tokenSeparator))
				return tokenSeparator;
			var tokenList = (IList<Token>)tokens;
			if (tokens.Any() && tokens.Last().TokenType == TokenType.String)
				return new Token(token, TokenType.StringContent);
			if (!string.IsNullOrEmpty(token))
				token = token.Trim();
			if (Regex.IsMatch(token, RegexConstants.Decimal))
				return new Token(token, TokenType.Decimal);
			if (Regex.IsMatch(token, RegexConstants.Integer))
				return new Token(token, TokenType.Integer);
			if (Regex.IsMatch(token, RegexConstants.Boolean, RegexOptions.IgnoreCase))
				return new Token(token, TokenType.Boolean);
			if (Regex.IsMatch(token, RegexConstants.StructuredReference, RegexOptions.IgnoreCase | RegexOptions.Multiline))
				return new Token(token, TokenType.StructuredReference);
			if (_nameValueProvider != null && _nameValueProvider.IsNamedValue(token, worksheet))
				return new Token(token, TokenType.NameValue);
			if (_functionNameProvider.IsFunctionName(token))
				return new Token(token, TokenType.Function);
			if (tokenList.Count > 0 && tokenList[tokenList.Count - 1].TokenType == TokenType.OpeningEnumerable)
				return new Token(token, TokenType.Enumerable);

			var addressType = ExcelAddress.IsValid(token);
			if (addressType == ExcelAddress.AddressType.InternalAddress)
				return new Token(token, TokenType.ExcelAddress);
			if (addressType == ExcelAddress.AddressType.ExternalAddress)
				return new Token(token, TokenType.InvalidReference);
			else if (addressType == ExcelAddress.AddressType.ExternalName)
				return new Token(token, TokenType.InvalidReference);
			if (addressType == ExcelAddress.AddressType.Invalid)
				return new Token(token, TokenType.InvalidReference);
			return new Token(token, TokenType.Unrecognized);
		}

		/// <summary>
		/// Creates a token with the specified <paramref name="explicitTokenType"/>.
		/// </summary>
		/// <param name="token">The value of the token to create.</param>
		/// <param name="explicitTokenType">The <see cref="TokenType"/> to assign to the new token.</param>
		/// <returns>The created token.</returns>
		public Token Create(string token, TokenType explicitTokenType)
		{
			return new Token(token, explicitTokenType);
		}
		#endregion
	}
}
