using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class IsEven : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
			if(ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var arg1 = GetFirstValue(arguments);//arguments.ElementAt(0);
            if (!ConvertUtil.IsNumeric(arg1))
            {
				return new CompileResult(eErrorType.Value);
            }
            var number = (int)System.Math.Floor(ConvertUtil.GetValueDouble(arg1));
            return CreateResult(number % 2 == 0, DataType.Boolean);
        }
    }
}
