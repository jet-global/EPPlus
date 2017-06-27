﻿using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formulas for the RANK, RANK.EQ, and RANK.AVG Excel Functions. Based on what is passed into the 
	/// constructor is the function that is executed (Specifically RANK and RANK.EQ v.s. RANK.AVG).
	/// </summary>
	public class Rank : ExcelFunction
	{
		bool _isAvg;
		/// <summary>
		/// If the RANK.AVG Function is to be executed then true will be passed into the constructor. 
		/// </summary>
		/// <param name="isAvg"></param>
		public Rank(bool isAvg = false)
		{
			_isAvg = isAvg;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			if (arguments.ElementAt(0).ValueFirst == null)
				return new CompileResult(eErrorType.NA);

			var number = ArgToDecimal(arguments, 0);
			var refer = arguments.ElementAt(1);
			bool asc = false;
			if (arguments.Count() > 2)
			{
				if (arguments.ElementAt(2).Value is string)
					return new CompileResult(eErrorType.Value);
				asc = base.ArgToBool(arguments, 2);
			}
			var l = new List<double>();

			foreach (var c in refer.ValueAsRangeInfo)
			{
				var v = Utils.ConvertUtil.GetValueDouble(c.Value, false, true);
				if (c.Value is bool)
					continue;
				if (!double.IsNaN(v))
					l.Add(v);
			}
			l.Sort();
			double ix;
			if (asc)
			{
				ix = l.IndexOf(number) + 1;
				if (_isAvg)
				{
					int st = Convert.ToInt32(ix);
					while (l.Count > st && l[st] == number) st++;
					if (st > ix) ix = ix + ((st - ix) / 2D);
				}
			}
			else
			{
				ix = l.LastIndexOf(number);
				if (_isAvg)
				{
					int st = Convert.ToInt32(ix) - 1;
					while (0 <= st && l[st] == number) st--;
					if (st + 1 < ix) ix = ix - ((ix - st - 1) / 2D);
				}
				ix = l.Count - ix;
			}
			if (ix <= 0 || ix > l.Count)
			{
				return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
			}
			else
			{
				return CreateResult(ix, DataType.Decimal);
			}
		}
	}
}
