using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Utils
{
	[TestClass]
	public class ConvertUtilTest
	{
		#region TryParse Tests
		[TestMethod]
		public void ValidateNumberGroupSizesTest()
		{
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes(null, CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1.442E+003", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1442", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1442.00", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1,442.0", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-5.26E-003", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-0", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-0.005", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-0.005260", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("2,345,656,567", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("123", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("12345", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-12345", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-12,345", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("-122,345", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("2,3,45,656,567", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("23,45,656,567", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("2345656,567", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("2345656567,", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("2,345,656,567,", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes(",567", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("-,567", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("2,345,656,7.45E-003", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("2,345,656,567E+003", CultureInfo.CurrentCulture.NumberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("2,345,656,567.45E-003", CultureInfo.CurrentCulture.NumberFormat));

			var numberFormat = new NumberFormatInfo { NumberGroupSizes = new int[4] { 2, 3, 4, 0 }, NumberGroupSeparator = "*" };
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("12", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("123", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1*23", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("123*23", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("6*123*23", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("9876*123*23", numberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("99876*123*23", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("9*9876*123*23", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("1231239*9876*123*23", numberFormat));

			numberFormat = new NumberFormatInfo { NumberGroupSizes = new int[2] { 2, 3,}, NumberGroupSeparator = "*", NumberDecimalSeparator = "#", NegativeSign = "_" };
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("_123*23#23", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("6*123*23#", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("126*123*23#8989", numberFormat));
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("9126*123*23", numberFormat));

			numberFormat = new NumberFormatInfo { NumberGroupSizes = new int[0] };
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("2,345", numberFormat));
			Assert.IsTrue(ConvertUtil.ValidateNumberGroupSizes("2345", numberFormat));

			// In German, the "." is both the date seperator and the number group seperator.
			numberFormat = new CultureInfo("de-DE").NumberFormat;
			Assert.IsFalse(ConvertUtil.ValidateNumberGroupSizes("1.1", numberFormat));
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
		#endregion

		#region IsNumeric Tests
		[TestMethod]
		public void IsNumeric()
		{
			Assert.IsFalse(ConvertUtil.IsNumeric(null));
			Assert.IsTrue(ConvertUtil.IsNumeric((byte)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((short)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((int)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((long)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((Single)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((double)5));
			Assert.IsTrue(ConvertUtil.IsNumeric((decimal)5));
			Assert.IsTrue(ConvertUtil.IsNumeric(true));
			// We've never seen a char come through Excel, so we'll let it be a number for now.
			Assert.IsTrue(ConvertUtil.IsNumeric('5'));
			Assert.IsFalse(ConvertUtil.IsNumeric("5"));
			// Excel treats dates as numeric, but not date strings.
			Assert.IsFalse(ConvertUtil.IsNumeric("1/1/2000"));
			Assert.IsTrue(ConvertUtil.IsNumeric(new DateTime(2000, 1, 1)));
			Assert.IsTrue(ConvertUtil.IsNumeric(new TimeSpan(5, 0, 0, 0)));
		}
		#endregion

		#region GetValueDouble Tests
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
		#endregion

		#region TryParseObjectToDecimal Tests
		[TestMethod]
		public void TryParseObjectToDecimalParsesDateTimeObject()
		{
			var date = new DateTime(1900, 3, 1);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(date, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesInt()
		{
			var expectedDate = new DateTime(1900, 3, 1);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(61, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesDouble()
		{
			var expectedDate = new DateTime(1900, 3, 1, 12, 0, 0);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(61.5, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesDecimal()
		{
			var expectedDate = new DateTime(1900, 3, 1, 12, 0, 0);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal((decimal)61.5, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesNegativeInt()
		{
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(-1, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(-1, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesNegativeDouble()
		{
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(-1.5, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(-1.5, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesZero()
		{
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(0, out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(0, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesIntWithinString()
		{
			var expectedDate = new DateTime(1900, 3, 1);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal("61", out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesDoubleWithinString()
		{
			var expectedDate = new DateTime(1900, 3, 1, 12, 0, 0);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal("61.5", out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesNegativeIntWithinString()
		{
			var isValidDate = ConvertUtil.TryParseObjectToDecimal("-1", out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(-1, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesNegativeDoubleWithinString()
		{
			var isValidDate = ConvertUtil.TryParseObjectToDecimal("-1.5", out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(-1.5, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalDoesNotParseNonDateString()
		{
			var isValidDate = ConvertUtil.TryParseObjectToDecimal("word", out double OADate);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(-1.0, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesDateAsString()
		{
			var expectedDate = new DateTime(1900, 3, 1, 5, 56, 59);
			var isValidDate = ConvertUtil.TryParseObjectToDecimal("3/1/1900 5:56:59", out double OADate);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesDoublesInStringsAsDoublesCorrectly()
		{
			var testNumber = "1.11";
			var isValidOADate = ConvertUtil.TryParseObjectToDecimal(testNumber, out double OADate);
			Assert.IsTrue(isValidOADate);
			Assert.AreEqual(1.11, OADate);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesMilitaryTime()
		{
			Assert.IsTrue(ConvertUtil.TryParseObjectToDecimal("23:59", out double oaDate));
			Assert.AreEqual(0.99, oaDate, 0.01);
		}

		[TestMethod]
		public void TryParseObjectToDecimalParsesStringsCorrectly()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				{
					// This should parse as a decimal value under the US culture.
					var decimalValue = "1.11";
					var isValidDate = ConvertUtil.TryParseObjectToDecimal(decimalValue, out double parseResult);
					Assert.IsTrue(isValidDate);
					Assert.AreEqual(1.11, parseResult);
					// DateTime parses this as a date (M.DD.YYYY) under the US culture,
					// but Excel does not recognize this as a valid date format under the US culture.
					var dateValue = "1.11.2017";
					var expectedDate = new DateTime(2017, 1, 11);
					isValidDate = ConvertUtil.TryParseObjectToDecimal(dateValue, out parseResult);
					Assert.IsTrue(isValidDate);
					Assert.AreEqual(expectedDate.ToOADate(), parseResult);
					// DateTime parses this as a valid date under the US culture,
					// but Excel does not recognize this as a valid date format under the US culture.
					var USShortDate = "1,11";
					expectedDate = new DateTime(DateTime.Today.Year, 1, 11);
					isValidDate = ConvertUtil.TryParseObjectToDecimal(USShortDate, out parseResult);
					Assert.IsTrue(isValidDate);
					Assert.AreEqual(expectedDate.ToOADate(), parseResult);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				{
					// This should parse as a date (D.MM.CurrentYear) under the German culture.
					var GermanShortDate = "1.11";
					var isValidDate = ConvertUtil.TryParseObjectToDecimal(GermanShortDate, out double parseResult);
					var expectedDate = new DateTime(DateTime.Today.Year, 11, 1);
					Assert.IsTrue(isValidDate);
					Assert.AreEqual(expectedDate.ToOADate(), parseResult);
					// This should parse as a date (D.MM.YYYY) under the German culture.
					var GermanDate = "1.11.2017";
					expectedDate = new DateTime(2017, 11, 1);
					isValidDate = ConvertUtil.TryParseObjectToDecimal(GermanDate, out parseResult);
					Assert.IsTrue(isValidDate);
					Assert.AreEqual(expectedDate.ToOADate(), parseResult);
					// This should parse as a decimal value under the German culture.
					var GermanDecimalValue = "1,11";
					isValidDate = ConvertUtil.TryParseObjectToDecimal(GermanDecimalValue, out parseResult);
					Assert.IsTrue(isValidDate);
					Assert.AreEqual(1.11, parseResult);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void TryParseDateObjectAsOADateParsesPeriodFormatDateInString()
		{
			// Note that although System.DateTime considers "1.11.2017" as a valid date format under
			// the US culture, Excel does not. EPPlus therefore does not replicate Excel's behavior in
			// that regard. It is currently considered too much work for too little value to
			// properly replicate Excel's behavior with dates of this format in EPPlus.
			var expectedDate = new DateTime(2017, 1, 11);
			var testNumber = "1.11.2017";
			var isValidDate = ConvertUtil.TryParseObjectToDecimal(testNumber, out double OADate);
			Assert.AreEqual(expectedDate.ToOADate(), OADate);
		}
		#endregion

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

		[TestMethod]
		public void TryParseDateObjectWithDateTimeObjectReturnsCorrectResult()
		{
			var date = new DateTime(2017, 5, 23);
			var isValidDate = ConvertUtil.TryParseDateObject(date, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectWithDateTimeObjectWithOADateLessThan61ReturnsCorrectResult()
		{
			var date = new DateTime(1900, 2, 28);
			var isValidDate = ConvertUtil.TryParseDateObject(date, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectHandlesOffByOneErrorFor27February1900()
		{
			var date = new DateTime(1900, 2, 27);
			Assert.AreEqual(59, date.ToOADate()); // The OADate from System.DateTime for 2/27/1900 is 59.
												  // The OADate from Excel for 2/27/1900 is 58.
			var isValidDate = ConvertUtil.TryParseDateObject(58, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			// Calculate the date using Excel's OADates, not System.DateTime's OADates.
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectHandlesOffByOneErrorFor28February1900()
		{
			var date = new DateTime(1900, 2, 28);
			Assert.AreEqual(60, date.ToOADate()); // The OADate from System.DateTime for 2/28/1900 is 60.
												  // The OADate from Excel for 2/28/1900 is 59.
			var isValidDate = ConvertUtil.TryParseDateObject(59, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectTreatsNonExistentDate29February1900As1March1900()
		{
			var date = new DateTime(1900, 3, 1);
			Assert.AreEqual(61, date.ToOADate()); // The OADate from System.DateTime for 3/1/1900 is 61.
												  // The OADate from Excel for 2/29/1900 is 60, a day that doesn't exist.
												  // System.DateTime uses OADate 60 for 2/28/1900 and considers 2/29/1900 an invalid date.
												  // Since 60 cannot be parsed as 2/29/1900, it is instead parsed as 3/1/1900.
												  // So using 60 or 61 as the input object in TryParseDateObject will produce the date 3/1/1900.
			var isValidDate = ConvertUtil.TryParseDateObject(60, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectUsesSameOADatesForExcelAndDateTimeFor1March1900()
		{
			// This test is a reminder that Excel and System.DateTime have their OADates sync back
			// up at the OADate 61, which is 3/1/1900; all dates after 3/1/1900 have the same OADate
			// in Excel and System.DateTime.
			var date = new DateTime(1900, 3, 1);
			Assert.AreEqual(61, date.ToOADate());
			var isValidDate = ConvertUtil.TryParseDateObject(61, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
			
		}

		[TestMethod]
		public void TryParseDateObjectUsesSameOADatesForExcelAndDateTimeAfter1March1900()
		{
			// This test is to confirm that OADates after 3/1/1900 are the same
			// in Excel and System.DateTime.
			var date = new DateTime(1900, 3, 2);
			Assert.AreEqual(62, date.ToOADate());
			var isValidDate = ConvertUtil.TryParseDateObject(62, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectWithDateAsDoubleNear1March1900ReturnsCorrectYearMonthDay()
		{
			// Test the case where a time and day close to 3/1/1900 returns
			// the correct result. Note that Excel represents 2/28/1900 as the
			// OADate 59.
			var date = new DateTime(1900, 2, 28);
			var isValidDate = ConvertUtil.TryParseDateObject(59.99999, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(1900, resultDate.Year);
			Assert.AreEqual(2, resultDate.Month);
			Assert.AreEqual(28, resultDate.Day);
			Assert.AreNotEqual(date.TimeOfDay, resultDate.TimeOfDay);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectWithOADate1ReturnsCorrectResult()
		{
			var date = new DateTime(1900, 1, 1);
			var isValidDate = ConvertUtil.TryParseDateObject(1, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(true, isValidDate);
			Assert.AreEqual(date, resultDate);
			Assert.AreEqual(null, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectWithOADate0ReturnsInvalidDate()
		{
			var isValidDate = ConvertUtil.TryParseDateObject(0, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(eErrorType.Num, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectWithTimeComponentForOADate0ReturnsInvalidDate()
		{
			var isValidDate = ConvertUtil.TryParseDateObject(0.5, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(eErrorType.Num, resultError);
		}

		[TestMethod]
		public void TryParseDateObjectWithNegativeOADate0WithTimeComponentReturnsInvalidDate()
		{
			var isValidDate = ConvertUtil.TryParseDateObject(0.5, out DateTime resultDate, out eErrorType? resultError);
			Assert.AreEqual(false, isValidDate);
			Assert.AreEqual(eErrorType.Num, resultError);
		}
		#endregion

		#region ConvertObjectToXmlAttributeString Tests
		[TestMethod]
		public void ConvertObjectToXmlAttributeString()
		{
			var dateTime = DateTime.Now;
			Assert.AreEqual(dateTime.ToString("yyyy’-‘MM’-‘dd’T’HH’:’mm’:’ss"), ConvertUtil.ConvertObjectToXmlAttributeString(dateTime));
			Assert.IsNull(ConvertUtil.ConvertObjectToXmlAttributeString(null));
			Assert.AreEqual("832", ConvertUtil.ConvertObjectToXmlAttributeString(832));
			Assert.AreEqual("832.382", ConvertUtil.ConvertObjectToXmlAttributeString(832.382));
			Assert.AreEqual("jet", ConvertUtil.ConvertObjectToXmlAttributeString("jet"));
			Assert.AreEqual("0", ConvertUtil.ConvertObjectToXmlAttributeString(false));
			Assert.AreEqual("#NAME?", ConvertUtil.ConvertObjectToXmlAttributeString(ExcelErrorValue.Create(eErrorType.Name)));
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void ConvertObjectToXmlAttributeStringInvalidType()
		{
			ConvertUtil.ConvertObjectToXmlAttributeString(new List<int>());
		}
		#endregion
	}
}