using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class CountBlank : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var arg = arguments.First();
			if (!arg.IsExcelRange) throw new InvalidOperationException("CountBlank only support ranges as arguments");
			var result = arg.ValueAsRangeInfo.GetNCells();
			foreach (var cell in arg.ValueAsRangeInfo)
			{
				if (!(cell.Value == null || cell.Value.ToString() == string.Empty))
				{
					result--;
				}
			}
			return CreateResult(result, DataType.Integer);
		}
	}
}
