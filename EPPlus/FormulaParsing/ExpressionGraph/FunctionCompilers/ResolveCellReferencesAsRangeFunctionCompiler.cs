using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	class ResolveCellReferencesAsRangeFunctionCompiler : DefaultCompiler
	{
		public ResolveCellReferencesAsRangeFunctionCompiler(ExcelFunction function) : base(function) { }

		public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
		{
			// EPPlus handles operators as members of the child expression instead of as functions of their own.
			// We want to exclude those children whose values are part of an operator expression (since they will be resolved by the operator).
			var ignoreOperators = children.Where(child => child.Children.All(grandkid => grandkid.Operator == null));
			// Typically the Expressions will be FunctionArgumentExpressions, equivalent to the NimbusExcelFormulaCell,
			// so any of their children will be the actual expression arguments to compile, most notably this will
			// be the ExcelAddressExpression who's results we want to manipulate for resolving arguments.
			var childrenToResolveAsRange = ignoreOperators.SelectMany(child => child.Children).Where(child => child is ExcelAddressExpression);
			foreach (ExcelAddressExpression excelAddress in childrenToResolveAsRange)
			{
				excelAddress.ResolveAsRange = true;
			}
			return base.Compile(children, context);
		}
	}
}
