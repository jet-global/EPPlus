using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Index : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if(ValidateArguments(arguments, 2) == false)
            	return new CompileResult(eErrorType.Value);
            var arg1 = arguments.ElementAt(0);
            var args = arg1.Value as IEnumerable<FunctionArgument>;
            var crf = new CompileResultFactory();
            if (args != null)
            {
                var index = ArgToInt(arguments, 1);
                if (index > args.Count())
                {
                    throw new ExcelErrorValueException(eErrorType.Ref);
                }
                var candidate = args.ElementAt(index - 1);
                return crf.Create(candidate.Value);
            }
            if (arg1.IsExcelRange)
            {
                var row = ArgToInt(arguments, 1);                 
                var col = arguments.Count()>2 ? ArgToInt(arguments, 2) : 1;
                var ri=arg1.ValueAsRangeInfo;
                if (row > ri.Address._toRow - ri.Address._fromRow + 1 ||
                    col > ri.Address._toCol - ri.Address._fromCol + 1)
                {
					return new CompileResult(eErrorType.Value);
                }
                var candidate = ri.GetOffset(row-1, col-1);
                return crf.Create(candidate);
            }
            throw new NotImplementedException();
        }
    }
}
