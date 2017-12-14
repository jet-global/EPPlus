using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formulas for the RANK, RANK.EQ, and RANK.AVG Excel Functions. The function to be executed is 
	/// determined dynamically based on the arguments to the constructor (RANK and RANK.EQ vs. RANK.AVG).
	/// </summary>
	public class Rank : ExcelFunction
	{
		private bool _isAvg;
		/// <summary>
		/// Creates a new RANK Function.
		/// </summary>
		/// <param name="isAvg">Indicates if the function should use the RANK.AVG variant instad of the default RANK.EQ
		/// implementation.</param>
		public Rank(bool isAvg = false)
		{
			this._isAvg = isAvg;
		}
		/// <summary>
		/// Takse the user specified arguments and returns the  the position of the first argument in relation to the 
		/// other numbers in a sorted list. For example in a scrambled list of the numbers 1-100 the number 99 would have 
		/// a rank of 2. 
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
				ascendingOrder = base.ArgToBool(arguments.ElementAt(2));
			}

			var numberList = new List<double>();
			foreach (var cell in reference.ValueAsRangeInfo)
			{
				// Boolean and string values in a referenced range are ignored by the function. 
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
					int duplicateResultCount = Convert.ToInt32(rankResult);
					while (numberList.Count > duplicateResultCount && numberList[duplicateResultCount] == number) duplicateResultCount++;
					if (duplicateResultCount > rankResult) rankResult = rankResult + ((duplicateResultCount - rankResult) / 2D);
				}
			}
			else
			{
				rankResult = numberList.LastIndexOf(number);
				if (_isAvg)
				{
					int duplicateResultCount = Convert.ToInt32(rankResult) - 1;
					while (0 <= duplicateResultCount && numberList[duplicateResultCount] == number) duplicateResultCount--;
					if (duplicateResultCount + 1 < rankResult) rankResult = rankResult - ((rankResult - duplicateResultCount - 1) / 2D);
				}
				rankResult = numberList.Count - rankResult;
			}
			if (rankResult <= 0 || rankResult > numberList.Count)
			{
				return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
			}
			else
			{
				return this.CreateResult(rankResult, DataType.Decimal);
			}
		}
	}
}
