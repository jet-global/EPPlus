using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
	[TestClass]
	public class ValueMatcherTests
	{
		private ValueMatcher _matcher;

		[TestInitialize]
		public void Setup()
		{
			_matcher = new ValueMatcher();
		}

		[TestMethod]
		public void ShouldReturn1WhenFirstParamIsSomethingAndSecondParamIsNull()
		{
			object o1 = 1;
			object o2 = null;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturnMinus1WhenFirstParamIsNullAndSecondParamIsSomething()
		{
			object o1 = null;
			object o2 = 1;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}

		[TestMethod]
		public void ShouldReturn0WhenBothParamsAreNull()
		{
			object o1 = null;
			object o2 = null;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn0WhenBothParamsAreEqual()
		{
			object o1 = 1d;
			object o2 = 1d;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturnMinus1WhenFirstParamIsLessThanSecondParam()
		{
			object o1 = 1d;
			object o2 = 5d;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}

		[TestMethod]
		public void ShouldReturn1WhenFirstParamIsGreaterThanSecondParam()
		{
			object o1 = 3d;
			object o2 = 1d;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturn0WhenWhenParamsAreEqualStrings()
		{
			object o1 = "T";
			object o2 = "T";
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn0WhenParamsAreEqualButDifferentTypes()
		{
			object o1 = "2";
			object o2 = 2d;
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result, "IsMatch did not return 0 as expected when first param is a string and second a double");

			o1 = 2d;
			o2 = "2";
			result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result, "IsMatch did not return 0 as expected when first param is a double and second a string");
		}

		[TestMethod]
		public void ShouldReturnNullWhenTypesDifferAndStringConversionToDoubleFails()
		{
			object o1 = 2d;
			object o2 = "T";
			var result = _matcher.IsMatch(o1, o2);
			Assert.IsNull(result);
		}

		[TestMethod]
		public void ShouldReturn0ForEqualDates()
		{
			object o1 = DateTime.Parse("1/1/2017");
			object o2 = DateTime.Parse("1/1/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn1ForGreaterFirstDate()
		{
			object o1 = DateTime.Parse("1/2/2017");
			object o2 = DateTime.Parse("1/1/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturnNegative1ForLesserFirstDate()
		{
			object o1 = DateTime.Parse("1/1/2017");
			object o2 = DateTime.Parse("1/2/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}

		[TestMethod]
		public void ShouldReturn0ForEqualDatesFirstIsString()
		{
			object o1 = "1/1/2017";
			object o2 = DateTime.Parse("1/1/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn1ForGreaterFirstDateFirstIsString()
		{
			object o1 = "1/2/2017";
			object o2 = DateTime.Parse("1/1/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturnNegative1ForLesserFirstDateFirstIsString()
		{
			object o1 = "1/1/2017";
			object o2 = DateTime.Parse("1/2/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}

		[TestMethod]
		public void ShouldReturn0ForEqualDatesSecondIsString()
		{
			object o1 = DateTime.Parse("1/1/2017");
			object o2 = "1/1/2017";
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn1ForGreaterFirstDateSecondIsString()
		{
			object o1 = DateTime.Parse("1/2/2017");
			object o2 = "1/1/2017";
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturnNegative1ForLesserFirstDateSecondIsString()
		{
			object o1 = DateTime.Parse("1/1/2017");
			object o2 = "1/2/2017";
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}

		[TestMethod]
		public void ShouldReturn0ForEqualDateFirstDateIsOADate()
		{
			object o1 = DateTime.Parse("1/1/2017").ToOADate();
			object o2 = DateTime.Parse("1/1/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn1ForGreaterFirstDateFirstDateIsOADate()
		{
			object o1 = DateTime.Parse("1/2/2017").ToOADate();
			object o2 = DateTime.Parse("1/1/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturnNegative1ForLesserFirstDateFirstDateIsOADate()
		{
			object o1 = DateTime.Parse("1/1/2017").ToOADate();
			object o2 = DateTime.Parse("1/2/2017");
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}

		[TestMethod]
		public void ShouldReturn0ForEqualDateSecondDateIsOADate()
		{
			object o1 = DateTime.Parse("1/1/2017");
			object o2 = DateTime.Parse("1/1/2017").ToOADate();
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(0, result);
		}

		[TestMethod]
		public void ShouldReturn1ForGreaterFirstDateSecondDateIsOADate()
		{
			object o1 = DateTime.Parse("1/2/2017");
			object o2 = DateTime.Parse("1/1/2017").ToOADate();
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void ShouldReturnNegative1ForLesserFirstDateSecondDateIsOADate()
		{
			object o1 = DateTime.Parse("1/1/2017");
			object o2 = DateTime.Parse("1/2/2017").ToOADate();
			var result = _matcher.IsMatch(o1, o2);
			Assert.AreEqual(-1, result);
		}
	}
}
