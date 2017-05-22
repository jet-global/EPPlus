using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Utils
{
	[TestClass]
	public class ConvertUtilTest
	{
		[TestMethod]
		public void TryParseNumericString()
		{
			double result;
			object numericString = null;
			double expected = 0;
			Assert.IsFalse(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			expected = 1442.0;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // 1.442E+003
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // 1442
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f2", CultureInfo.CurrentCulture); // 1442.00
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("n", CultureInfo.CurrentCulture); // 1,442.0
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			expected = -0.00526;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // -5.26E-003
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // -0
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(0.0, result);
			numericString = expected.ToString("f3", CultureInfo.CurrentCulture); // -0.005
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(-0.005, result);
			numericString = expected.ToString("n6", CultureInfo.CurrentCulture); // -0.005260
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
		}

		[TestMethod]
		public void TryParseDateString()
		{
			DateTime result;
			object dateString = null;
			DateTime expected = DateTime.MinValue;
			Assert.IsFalse(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			expected = new DateTime(2013, 1, 15);
			dateString = expected.ToString("d", CultureInfo.CurrentCulture); // 1/15/2013
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("D", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 12:00:00 AM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 12:00 AM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			expected = new DateTime(2013, 1, 15, 15, 26, 32);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 3:26:32 PM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 3:26 PM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(new DateTime(2013, 1, 15, 15, 26, 0), result);
		}

		[TestMethod]
		public void IsNumeric()
		{
			Assert.Fail("This will fail until IsNumeric is fixed to not consider chars numbers.");
			Assert.IsFalse(ConvertUtil.IsNumeric(null));
			Assert.IsTrue(ConvertUtil.IsNumeric((byte)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((short)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((int)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((long)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((Single)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((double)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((decimal)5));
			Assert.IsTrue(ConvertUtil.IsNumeric(true));
			Assert.IsFalse(ConvertUtil.IsNumeric('5'));
			Assert.IsFalse(ConvertUtil.IsNumeric("5"));
			// Excel treats dates as numeric, but not date strings.
			Assert.IsFalse(ConvertUtil.IsNumeric("1/1/2000"));
			Assert.IsTrue(ConvertUtil.IsNumeric(new DateTime(2000, 1, 1)));
			Assert.IsTrue(ConvertUtil.IsNumeric(new TimeSpan(5, 0, 0, 0)));
		}

		[TestMethod]
		public void GetValueDouble()
		{
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((byte)5));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((short)5));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((int)5));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((long)5));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((Single)5));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((double)5));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble((decimal)5));
			Assert.AreEqual(1, ConvertUtil.GetValueDouble(true));
			Assert.AreEqual(0, ConvertUtil.GetValueDouble(true, true));
			Assert.AreEqual(36526d, ConvertUtil.GetValueDouble(new DateTime(2000, 1, 1)));
			Assert.AreEqual(double.NaN, ConvertUtil.GetValueDouble("1/1/2000", retNaN: true));
			Assert.AreEqual(5d, ConvertUtil.GetValueDouble(new TimeSpan(5, 0, 0, 0)));
			Assert.AreEqual(0d, ConvertUtil.GetValueDouble('a'));
			Assert.AreEqual(double.NaN, ConvertUtil.GetValueDouble('a', false, true));
			Assert.AreEqual(0d, ConvertUtil.GetValueDouble("Not a number"));
			Assert.AreEqual(double.NaN, ConvertUtil.GetValueDouble("Not a number", false, true));
		}

		#region TryParseDateObject Tests
		[TestMethod]
		public void TryParseDateObjectParsesInt()
		{
			var isValidDate = ConvertUtil.TryParseDateObject(42874, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(42874, resultDate.ToOADate());
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectParsesIntWithinString()
		{
			var isValidDate = ConvertUtil.TryParseDateObject("42874", out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(42874, resultDate.ToOADate());
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectParsesDateAsString()
		{
			var isValidDate = ConvertUtil.TryParseDateObject("5/19/2017", out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(42874, resultDate.ToOADate());
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectDoesNotParseNegativeDouble()
		{
			var isValidDate = ConvertUtil.TryParseDateObject(-1.5, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(0, resultDate.ToOADate());
			Assert.AreEqual(eErrorType.Num, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectDoesNotParseNegativeDoubleWithinString()
		{
			var isValidDate = ConvertUtil.TryParseDateObject("-1.5", out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(0, resultDate.ToOADate());
			Assert.AreEqual(eErrorType.Num, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectDoesNotParseNonDateString()
		{
			var isValidDate = ConvertUtil.TryParseDateObject("word", out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(0, resultDate.ToOADate());
			Assert.AreEqual(eErrorType.Value, resultError);
		}
		#endregion
	}
}
