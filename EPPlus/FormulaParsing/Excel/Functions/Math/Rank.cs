using System;
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
		/// Takse the user specified arguments and returns the rank of the first argument; the position of the first argument
		/// in relation to the other numbers in a list. 
		/// </summary>
		/// <param name="arguments">The user specified arguments, the first being the number whos rank is to be found and
		/// the second is the list of numbers as a cell reference.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The rank of the given number.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			if (arguments.ElementAt(0).ValueFirst == null)
				return new CompileResult(eErrorType.NA);

			var number = ArgToDecimal(arguments, 0);
			var reference = arguments.ElementAt(1);
			bool ascendingOrder = false;

			if (arguments.Count() > 2)
			{
				if (arguments.ElementAt(2).Value is string)
					return new CompileResult(eErrorType.Value);
				ascendingOrder = base.ArgToBool(arguments, 2);
			}

			var numberList = new List<double>();
			foreach (var cell in reference.ValueAsRangeInfo)
			{
				if (cell.Value is bool)
					continue;
				var valAsDouble = Utils.ConvertUtil.GetValueDouble(cell.Value, false, true);
				if (!double.IsNaN(valAsDouble))
					numberList.Add(valAsDouble);
			}
			numberList.Sort();
			double rankResult;
			if (ascendingOrder)
			{
				rankResult = numberList.IndexOf(number) + 1;
				if (_isAvg)
				{
					int st = Convert.ToInt32(rankResult);
					while (numberList.Count > st && numberList[st] == number) st++;
					if (st > rankResult) rankResult = rankResult + ((st - rankResult) / 2D);
				}
			}
			else
			{
				rankResult = numberList.LastIndexOf(number);
				if (_isAvg)
				{
					int intRank = Convert.ToInt32(rankResult) - 1;
					while (0 <= intRank && numberList[intRank] == number) intRank--;
					if (intRank + 1 < rankResult) rankResult = rankResult - ((rankResult - intRank - 1) / 2D);
				}
				rankResult = numberList.Count - rankResult;
			}
			if (rankResult <= 0 || rankResult > numberList.Count)
			{
				return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
			}
			else
			{
				return CreateResult(rankResult, DataType.Decimal);
			}
		}
	}
}
