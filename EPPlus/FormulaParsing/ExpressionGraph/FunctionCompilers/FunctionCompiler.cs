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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	public abstract class FunctionCompiler
	{
		#region Properties
		/// <summary>
		/// Gets the function to compile.
		/// </summary>
		protected ExcelFunction Function { get; private set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Default constructor for the <see cref="FunctionCompiler"/>.
		/// </summary>
		/// <param name="function">The <see cref="ExcelFunction"/> to compile.</param>
		public FunctionCompiler(ExcelFunction function)
		{
			if (function == null)
				throw new ArgumentNullException(nameof(function));
			this.Function = function;
		}
		#endregion

		#region Public Abstract Methods
		/// <summary>
		/// Compiles the function.
		/// </summary>
		/// <param name="children">The children of the function to compile.</param>
		/// <param name="context">The context to compile within.</param>
		/// <returns>A <see cref="CompileResult"/>.</returns>
		public abstract CompileResult Compile(IEnumerable<Expression> children, ParsingContext context);
		#endregion

		#region Protected Methods
		protected void BuildFunctionArguments(object result, DataType dataType, List<FunctionArgument> args)
		{
			if (result is IEnumerable<object> && !(result is ExcelDataProvider.IRangeInfo))
			{
				var argList = new List<FunctionArgument>();
				var objects = result as IEnumerable<object>;
				foreach (var arg in objects)
				{
					this.BuildFunctionArguments(arg, dataType, argList);
				}
				args.Add(new FunctionArgument(argList));
			}
			else
				args.Add(new FunctionArgument(result, dataType));
		}

		protected void ConfigureExcelAddressExpressionToResolveAsRange(IEnumerable<Expression> children)
		{
			// EPPlus handles operators as members of the child expression instead of as functions of their own.
			// We want to exclude those children whose values are part of an operator expression (since they will be resolved by the operator).
			if (children.Any(grandkid => grandkid.Operator != null))
				return;
			// Typically the Expressions will be FunctionArgumentExpressions, equivalent to the NimbusExcelFormulaCell,
			// so any of their children will be the actual expression arguments to compile, most notably this will
			// be the ExcelAddressExpression who's results we want to manipulate for resolving arguments.
			var childrenToResolveAsRange = children.Where(child => child is ExcelAddressExpression);
			foreach (ExcelAddressExpression excelAddress in childrenToResolveAsRange)
			{
				excelAddress.ResolveAsRange = true;
			}
		}
		#endregion
	}
}
