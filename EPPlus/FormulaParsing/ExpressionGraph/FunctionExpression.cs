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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Expression that handles execution of a function.
	/// </summary>
	public class FunctionExpression : AtomicExpression
	{
		#region Properties
		/// <summary>
		/// Gets a value indicating whether the formula expression has a first child with children of its own.
		/// </summary>
		public override bool HasChildren
		{
			get
			{
				return (this.Children.Any() && this.Children.First().Children.Any());
			}
		}
		private ParsingContext ParsingContext { get; }
		private FunctionCompilerFactory FunctionCompilerFactory { get; }
		private bool IsNegated { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="expression">should be the of the function</param>
		/// <param name="parsingContext"></param>
		/// <param name="isNegated">True if the numeric result of the function should be negated.</param>
		public FunctionExpression(string expression, ParsingContext parsingContext, bool isNegated)
			 : base(expression)
		{
			this.ParsingContext = parsingContext;
			this.FunctionCompilerFactory = new FunctionCompilerFactory(parsingContext.Configuration.FunctionRepository);
			this.IsNegated = isNegated;
			base.AddChild(new FunctionArgumentExpression(this));
		}
		#endregion

		#region Public Expression Overrides
		/// <summary>
		/// Compiles the expression string.
		/// </summary>
		/// <returns>A <see cref="CompileResult"/> containing either the resulting compiled value or an error.</returns>
		public override CompileResult Compile()
		{
			try
			{
				var function = this.ParsingContext.Configuration.FunctionRepository.GetFunction(this.ExpressionString);
				if (function == null)
				{
					if (this.ParsingContext.Debug)
					{
						this.ParsingContext.Configuration.Logger.Log(this.ParsingContext, string.Format("'{0}' is not a supported function", this.ExpressionString));
					}
					return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);
				}
				if (this.ParsingContext.Debug)
				{
					this.ParsingContext.Configuration.Logger.LogFunction(this.ExpressionString);
				}
				var compiler = this.FunctionCompilerFactory.Create(function);
				var result = compiler.Compile(this.Children.Any() ? this.Children : Enumerable.Empty<Expression>(), this.ParsingContext);
				if (this.IsNegated)
				{
					if (!result.IsNumeric)
					{
						if (this.ParsingContext.Debug)
						{
							var msg = string.Format("Trying to negate a non-numeric value ({0}) in function '{1}'",
								 result.Result, this.ExpressionString);
							this.ParsingContext.Configuration.Logger.Log(this.ParsingContext, msg);
						}
						if (result.ResultValue is ExcelErrorValue errorValue)
							return new CompileResult(errorValue);
						return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
					}
					return new CompileResult(result.ResultNumeric * -1, result.DataType);
				}
				return result;
			}
			catch (ExcelErrorValueException e)
			{
				if (this.ParsingContext.Debug)
				{
					this.ParsingContext.Configuration.Logger.Log(this.ParsingContext, e);
				}
				return new CompileResult(e.ErrorValue, DataType.ExcelError);
			}
		}

		/// <summary>
		/// Adds the current instance as a child of the base class.
		/// </summary>
		/// <returns>Returns the newly added child <see cref="Expression"/>.</returns>
		public override Expression PrepareForNextChild()
		{
			return base.AddChild(new FunctionArgumentExpression(this));
		}

		/// <summary>
		/// Adds the given <paramref name="child"/> to the last element in Children.
		/// </summary>
		/// <param name="child">The expression to be added to Children.</param>
		/// <returns>Returns the given <paramref name="child"/>.</returns>
		public override Expression AddChild(Expression child)
		{
			this.Children.Last().AddChild(child);
			return child;
		}
		#endregion
	}
}
