/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	/// <summary>
	/// For each of the children of the given function, if that child came from a cell reference,
	/// this compiler will set that child's ResolveAsRange flag to true.
	/// </summary>
	public class ResolveCellReferencesAsRangeCompiler : DefaultCompiler
	{
		/// <summary>
		/// Initializes a new ResolveCellReferencesAsRangeCompiler.
		/// </summary>
		/// <param name="function">The function to be compiled and executed.</param>
		public ResolveCellReferencesAsRangeCompiler(ExcelFunction function) : base(function) { }

		/// <summary>
		/// Set each of this function's arguments to be resolved as a cell reference if that argument is a cell reference.
		/// </summary>
		/// <param name="children">The uncompiled arguments for this function.</param>
		/// <param name="context">The context for this function.</param>
		/// <returns>Returns the result of executing this function.</returns>
		public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
		{
			// EPPlus handles operators as members of the child expression instead of as functions of their own.
			// We want to exclude those children whose values are part of an operator expression (since they will be resolved by the operator).
			var ignoreOperators = children.Where(child => child.Children.All(grandkid => grandkid.Operator == null));
			// Typically the Expressions will be FunctionArgumentExpressions, equivalent to the NimbusExcelFormulaCell,
			// so any of their children will be the actual expression arguments to compile, most notably this will
			// be the ExcelAddressExpression who's results we want to manipulate for resolving arguments.
			var childrenToResolveAsRange = ignoreOperators.SelectMany(child => child.Children).Where(child => child is ExcelAddressExpression);
			foreach (ExcelAddressExpression excelAddress in childrenToResolveAsRange)
			{
				excelAddress.ResolveAsRange = true;
			}
			return base.Compile(children, context);
		}
	}
}
