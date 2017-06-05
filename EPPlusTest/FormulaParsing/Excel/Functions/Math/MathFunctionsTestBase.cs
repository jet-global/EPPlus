using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public abstract class MathFunctionsTestBase
	{
		#region Properties
		protected ParsingContext ParsingContext { get; } = ParsingContext.Create();
		#endregion
	}
}
