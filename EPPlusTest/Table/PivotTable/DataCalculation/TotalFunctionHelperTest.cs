using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.DataCalculation;

namespace EPPlusTest.Table.PivotTable.DataCalculation
{
	[TestClass]
	public class TotalFunctionHelperTest
	{
		#region Calculate Tests
		[TestMethod]
		public void CalculateAverageDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Average, values);
			Assert.AreEqual(7.571428571, Math.Round((double)result, 9));
		}

		[TestMethod]
		public void CalculateCountDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Count, values);
			Assert.AreEqual(7, (double)result);
		}

		[TestMethod]
		public void CalculateMaxDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Max, values);
			Assert.AreEqual(22, (double)result);
		}

		[TestMethod]
		public void CalculateMinDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Min, values);
			Assert.AreEqual(1, (double)result);
		}

		[TestMethod]
		public void CalculateSumDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.None, values);
			Assert.AreEqual(53, (double)result);
		}

		[TestMethod]
		public void CalculateProductDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Product, values);
			Assert.AreEqual(92400, (double)result);
		}

		[TestMethod]
		public void CalculateStdDevDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.StdDev, values);
			Assert.AreEqual(7.044078905, Math.Round((double)result, 9));
		}

		[TestMethod]
		public void CalculateStdDevPDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.StdDevP, values);
			Assert.AreEqual(6.521549835, Math.Round((double)result, 9));
		}

		[TestMethod]
		public void CalculateVarDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Var, values);
			Assert.AreEqual(49.61904762, Math.Round((double)result, 8));
		}

		[TestMethod]
		public void CalculateVarPDataFieldFunctionTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var values = new List<object>() { 5, 2, 1, 6, 10, 22, 7 };
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.VarP, values);
			Assert.AreEqual(42.53061224, Math.Round((double)result, 8));
		}

		[TestMethod]
		public void CalculateNullValuesTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Average, null);
			Assert.IsNull(result);
		}

		[TestMethod]
		public void CalculateEmptyValuesTest()
		{
			var totalFunctionHelper = new TotalsFunctionHelper();
			var result = totalFunctionHelper.Calculate(OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Average, new List<object>());
			Assert.IsNull(result);
		}
		#endregion

		#region GetTokenNameValues Test
		[TestMethod]
		public void GetTokenNameValuesTest()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				helper.AddNames(new HashSet<string> { "Cost", "Count" });
				var tokens = helper.Tokenize("Count*Cost");
				Assert.AreEqual("Count", tokens.ElementAt(0).Value);
				Assert.AreEqual("*", tokens.ElementAt(1).Value);
				Assert.AreEqual("Cost", tokens.ElementAt(2).Value);
			}
		}

		[TestMethod]
		public void GetTokenNameValuesWithFunctionTest()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				helper.AddNames(new HashSet<string> { "Cost", "Count", "Item" });
				var tokens = helper.Tokenize("SUM(Count)*COUNT(Cost) - Item");
				Assert.AreEqual("SUM", tokens.ElementAt(0).Value);
				Assert.AreEqual("(", tokens.ElementAt(1).Value);
				Assert.AreEqual("Count", tokens.ElementAt(2).Value);
				Assert.AreEqual(")", tokens.ElementAt(3).Value);
				Assert.AreEqual("*", tokens.ElementAt(4).Value);
				Assert.AreEqual("COUNT", tokens.ElementAt(5).Value);
				Assert.AreEqual("(", tokens.ElementAt(6).Value);
				Assert.AreEqual("Cost", tokens.ElementAt(7).Value);
				Assert.AreEqual(")", tokens.ElementAt(8).Value);
				Assert.AreEqual("-", tokens.ElementAt(9).Value);
				Assert.AreEqual("Item", tokens.ElementAt(10).Value);
			}
		}

		[TestMethod]
		public void GetTokenNameValuesEmptyFormulaTest()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				helper.AddNames(new HashSet<string> { "Count", "Item" });
				var tokens = helper.Tokenize(string.Empty);
				Assert.IsNull(tokens);
			}
		}

		[TestMethod]
		public void GetTokenNameValuesNullFormulaTest()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				helper.AddNames(new HashSet<string> { "Count", "Item" });
				var tokens = helper.Tokenize(null);
				Assert.IsNull(tokens);
			}
		}
		#endregion

		#region EvaluateCalculatedFieldFormula Tests
		[TestMethod]
		public void EvaluateCalculatedFieldFormula()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				var fieldValues = new Dictionary<string, List<object>>
				{
					{ "Count", new List<object> { 1, 2, 3, 4} },
					{ "Cost", new List<object> {5, 6, 7, 8 } }
				};
				helper.AddNames(new HashSet<string>(fieldValues.Keys));
				var result = helper.EvaluateCalculatedFieldFormula(fieldValues, "Count*Cost");
				Assert.AreEqual(260d, result);
			}
		}

		[TestMethod]
		public void EvaluateCalculatedFieldFormulaWithFunction()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				var fieldValues = new Dictionary<string, List<object>>
				{
					{ "Count", new List<object> { 1, 2, 3, 4} },
					{ "Cost", new List<object> {5, 6, 7, 8 } }
				};
				helper.AddNames(new HashSet<string>(fieldValues.Keys));
				var result = helper.EvaluateCalculatedFieldFormula(fieldValues, "COUNT(Count)*Cost");
				Assert.AreEqual(26d, result);
			}
		}

		[TestMethod]
		public void EvaluateCalculatedFieldFormulaSingleValue()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				var fieldValues = new Dictionary<string, List<object>>
				{
					{ "Count", new List<object> { 1, 2, 3, 4} },
					{ "Cost", new List<object> {5, 6, 7, 8 } }
				};
				helper.AddNames(new HashSet<string>(fieldValues.Keys));
				var result = helper.EvaluateCalculatedFieldFormula(fieldValues, "Count");
				Assert.AreEqual(10d, result);
			}
		}

		[TestMethod]
		public void EvaluateCalculatedFieldFormulaEmptyValue()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				var fieldValues = new Dictionary<string, List<object>>
				{
					{ "Count", new List<object> { 1, 2, 3, 4} },
					{ "Cost", new List<object> {5, 6, 7, 8 } }
				};
				helper.AddNames(new HashSet<string>(fieldValues.Keys));
				var result = helper.EvaluateCalculatedFieldFormula(fieldValues, string.Empty);
				Assert.AreEqual(null, result);
			}
		}

		[TestMethod]
		public void EvaluateCalculatedFieldFormulaNullValue()
		{
			using (var helper = new TotalsFunctionHelper())
			{
				var fieldValues = new Dictionary<string, List<object>>
				{
					{ "Count", new List<object> { 1, 2, 3, 4} },
					{ "Cost", new List<object> {5, 6, 7, 8 } }
				};
				helper.AddNames(new HashSet<string>(fieldValues.Keys));
				var result = helper.EvaluateCalculatedFieldFormula(fieldValues, null);
				Assert.AreEqual(null, result);
			}
		}
		#endregion
	}
}
