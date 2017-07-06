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
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);

			var cellRangeToCheck = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (cellRangeToCheck == null)
				return new CompileResult(eErrorType.Value);

			if (!IfHelper.TryExtractCriteriaString(arguments.ElementAt(1), context, out string criteriaString))
				return this.CreateResult(0d, DataType.Integer);

			var count = 0d;
			foreach (var cell in cellRangeToCheck)
			{
				if (IfHelper.ObjectMatchesCriteria(this.GetFirstArgument(cell.Value), criteriaString))
					count++;
			}

			return this.CreateResult(count, DataType.Integer);
		}
	}
}
