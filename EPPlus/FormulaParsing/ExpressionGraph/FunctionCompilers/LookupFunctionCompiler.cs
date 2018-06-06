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
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	/// <summary>
	/// Compiles expressions for a lookup-type function.
	/// </summary>
	public class LookupFunctionCompiler : FunctionCompiler
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="LookupFunctionCompiler"/>.
		/// </summary>
		/// <param name="function">The <see cref="ExcelFunction"/> to compile.</param>
		public LookupFunctionCompiler(ExcelFunction function) : base(function) { }
		#endregion

		#region FunctionCompiler Overrides
		/// <summary>
		/// Compiles the provided child <see cref="Expression"/>s and invokes the <see cref="ExcelFunction"/>.
		/// </summary>
		/// <param name="children">The child <see cref="Expression"/>s to compile.</param>
		/// <param name="context">The <see cref="ParsingContext"/> for this function compiler.</param>
		/// <returns>The <see cref="CompileResult"/> from the function.</returns>
		public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
		{
			var args = new List<FunctionArgument>();
			base.Function.BeforeInvoke(context);
			int i = 0;
			foreach (var child in children)
			{
				if (base.Function is LookupFunction lookupFunction && lookupFunction.LookupArgumentIndicies.Contains(i))
					this.ConfigureExcelAddressExpressionToResolveAsRange(child.Children);
				var arg = child.Compile();
				base.BuildFunctionArguments(arg?.Result, arg?.DataType ?? DataType.Unknown, args);
				i++;
			}
			return base.Function.Execute(args, context);
		}
		#endregion
	}
}
