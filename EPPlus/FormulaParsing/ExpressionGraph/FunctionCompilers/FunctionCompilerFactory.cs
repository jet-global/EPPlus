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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	public class FunctionCompilerFactory
	{
		private readonly Dictionary<Type, FunctionCompiler> _specialCompilers = new Dictionary<Type, FunctionCompiler>();

		public FunctionCompilerFactory(FunctionRepository repository)
		{
			_specialCompilers.Add(typeof(If), new IfFunctionCompiler(repository.GetFunction("if")));
			_specialCompilers.Add(typeof(IfError), new IfErrorFunctionCompiler(repository.GetFunction("iferror")));
			_specialCompilers.Add(typeof(IfNa), new IfNaFunctionCompiler(repository.GetFunction("ifna")));
			_specialCompilers.Add(typeof(Sum), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("sum")));
			_specialCompilers.Add(typeof(SumIf), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("sumif")));
			_specialCompilers.Add(typeof(SumIfs), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("sumifs")));
			_specialCompilers.Add(typeof(Count), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("count")));
			_specialCompilers.Add(typeof(CountA), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("counta")));
			_specialCompilers.Add(typeof(CountIf), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("countif")));
			_specialCompilers.Add(typeof(CountIfs), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("countifs")));
			_specialCompilers.Add(typeof(CountBlank), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("countblank")));
			_specialCompilers.Add(typeof(Average), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("average")));
			_specialCompilers.Add(typeof(AverageA), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("averagea")));
			_specialCompilers.Add(typeof(AverageIf), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("averageif")));
			_specialCompilers.Add(typeof(AverageIfs), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("averageifs")));
			_specialCompilers.Add(typeof(StdevP), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("stdev.p")));
			_specialCompilers.Add(typeof(StdevS), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("stdev.s")));
			_specialCompilers.Add(typeof(Stdevpa), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("stdevpa")));
			_specialCompilers.Add(typeof(Stdeva), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("stdeva")));
			_specialCompilers.Add(typeof(VarP), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("var.p")));
			_specialCompilers.Add(typeof(VarS), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("var.s")));
			_specialCompilers.Add(typeof(Vara), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("vara")));
			_specialCompilers.Add(typeof(Varpa), new ResolveCellReferencesAsRangeCompiler(repository.GetFunction("varpa")));
			foreach (var key in repository.CustomCompilers.Keys)
			{
				_specialCompilers.Add(key, repository.CustomCompilers[key]);
			}
		}

		private FunctionCompiler GetCompilerByType(ExcelFunction function)
		{
			var funcType = function.GetType();
			if (_specialCompilers.ContainsKey(funcType))
			{
				return _specialCompilers[funcType];
			}
			return new DefaultCompiler(function);
		}
		public virtual FunctionCompiler Create(ExcelFunction function)
		{
			if (function is LookupFunction) return new LookupFunctionCompiler(function);
			if (function is ErrorHandlingFunction) return new ErrorHandlingFunctionCompiler(function);
			return GetCompilerByType(function);
		}
	}
}
