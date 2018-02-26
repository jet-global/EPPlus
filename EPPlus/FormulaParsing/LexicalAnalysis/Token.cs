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

using System;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
	public class Token
	{
		#region Properties
		/// <summary>
		/// Gets or sets the value of the token.
		/// </summary>
		public string Value { get; internal set; }

		/// <summary>
		/// Gets or sets the <see cref="TokenType"/> of the token.
		/// </summary>
		public TokenType TokenType { get; internal set; }

		/// <summary>
		/// Gets a value indicating that the token is negated.
		/// </summary>
		public bool IsNegated { get; private set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="Token"/>.
		/// </summary>
		/// <param name="token">The value of the <see cref="Token"/>.</param>
		/// <param name="tokenType">The type of the <see cref="Token"/>.</param>
		public Token(string token, TokenType tokenType)
		{
			Value = token;
			TokenType = tokenType;
		}
		#endregion

		#region Publid Methods
		/// <summary>
		/// Marks the token as negated.
		/// </summary>
		public void Negate()
		{

			this.IsNegated = this.TokenType == TokenType.Decimal
				|| this.TokenType == TokenType.Integer
				|| this.TokenType == TokenType.ExcelAddress
				|| this.TokenType == TokenType.StructuredReference;
		}
		#endregion

		#region Object Overrides
		public override string ToString()
		{
			return this.TokenType.ToString() + ", " + Value;
		}
		#endregion
	}
}
