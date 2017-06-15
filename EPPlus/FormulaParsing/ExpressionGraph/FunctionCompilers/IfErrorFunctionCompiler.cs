using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	public class IfErrorFunctionCompiler : FunctionCompiler
	{
		public IfErrorFunctionCompiler(ExcelFunction function)
			 : base(function)
		{
			Require.That(function).Named("function").IsNotNull();

		}

		public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
		{
			if (children.Count() != 2)
				return new CompileResult(eErrorType.Value);
			var args = new List<FunctionArgument>();
			Function.BeforeInvoke(context);
			var firstChild = children.ElementAt(0);
			var lastChild = children.ElementAt(1);
			args.Add(new FunctionArgument(firstChild.Compile().Result));
			args.Add(new FunctionArgument(lastChild.Compile().Result));
			return Function.Execute(args, context);
		}
	}
}
