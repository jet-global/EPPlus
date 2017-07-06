using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class CountIf : ExcelFunction
	{
		private readonly ExpressionEvaluator _expressionEvaluator;

		public CountIf()
			 : this(new ExpressionEvaluator())
		{

		}

		public CountIf(ExpressionEvaluator evaluator)
		{
			Require.That(evaluator).Named("evaluator").IsNotNull();
			_expressionEvaluator = evaluator;
		}

		private bool Evaluate(object obj, string expression)
		{
			double? candidate = default(double?);
			if (IsNumeric(obj))
			{
				candidate = ConvertUtil.GetValueDouble(obj);
			}
			if (candidate.HasValue)
			{
				return _expressionEvaluator.Evaluate(candidate.Value, expression);
			}
			return _expressionEvaluator.Evaluate(obj, expression);
		}

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);

			var cellRangeToCheck = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (cellRangeToCheck == null)
				return new CompileResult(eErrorType.Value);

			string criteriaString = null;
			if (arguments.ElementAt(1).Value is ExcelDataProvider.IRangeInfo criteriaRange)
			{
				if (criteriaRange.IsMulti)
				{
					criteriaString = IfHelper.CalculateCriteria(arguments, context.ExcelDataProvider.GetRange(context.Scopes.Current.Address.Worksheet, 1, 1, "A1").Worksheet, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol).ToString().ToUpper();
				}
				else
					criteriaString = this.GetFirstArgument(arguments.ElementAt(1).ValueFirst).ToString().ToUpper();
			}
			else
				criteriaString = this.GetFirstArgument(arguments.ElementAt(1)).ValueFirst.ToString().ToUpper();

			var count = 0d;
			foreach (var currentCell in cellRangeToCheck)
			{
				if (IfHelper.ObjectMatchesCriteria(currentCell.Value, criteriaString))
					count++;
			}

			return this.CreateResult(count, DataType.Integer);

			//var range = functionArguments.ElementAt(0);
			//var criteria = GetFirstArgument(functionArguments.ElementAt(1)).ValueFirst != null ? GetFirstArgument(functionArguments.ElementAt(1)).ValueFirst.ToString() : string.Empty;
			//double result = 0d;
			//if (range.IsExcelRange)
			//{
			//	ExcelDataProvider.IRangeInfo rangeInfo = range.ValueAsRangeInfo;
			//	for (int row = rangeInfo.Address.Start.Row; row < rangeInfo.Address.End.Row + 1; row++)
			//	{
			//		for (int col = rangeInfo.Address.Start.Column; col < rangeInfo.Address.End.Column + 1; col++)
			//		{
			//			if (criteria != null && Evaluate(GetFirstArgument(rangeInfo.Worksheet.GetValue(row, col)), criteria))
			//			{
			//				result++;
			//			}
			//		}
			//	}
			//}
			//else if (range.Value is IEnumerable<FunctionArgument>)
			//{
			//	foreach (var arg in (IEnumerable<FunctionArgument>)range.Value)
			//	{
			//		if (Evaluate(arg.Value, criteria))
			//		{
			//			result++;
			//		}
			//	}
			//}
			//else
			//{
			//	if (Evaluate(range.Value, criteria))
			//	{
			//		result++;
			//	}
			//}
			//return CreateResult(result, DataType.Integer);
		}
	}
}
