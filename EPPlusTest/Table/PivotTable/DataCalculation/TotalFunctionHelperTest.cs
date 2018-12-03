using System;
using System.Collections.Generic;
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
	}
}
