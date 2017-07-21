/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.

* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Utils
{
	[TestClass]
	public class InternationalizationUtilityTests
	{
		#region InternationalizationUtility Tests
		[TestMethod]
		public void InternationaliztionUtilityParsesEnglishErrorValues()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-us");
				var valueErrorString = "#VALUE!";
				var numErrorString = "#NUM!";
				var div0ErrorString = "#DIV/0!";
				var nameErrorString = "#NAME?";
				var naErrorString = "#N/A";
				var refErrorString = "#REF!";
				var nullErrorString = "#NULL!";
				bool isErrorValue = false;
				ExcelErrorValue error = null;
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Value, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(numErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Num, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(div0ErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Div0, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nameErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Name, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(naErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.NA, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(refErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Ref, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nullErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Null, error.Type);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationaliztionUtilityParsesGermanErrorValues()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-de");
				var valueErrorString = "#WERT!";
				var numErrorString = "#ZAHL!";
				var div0ErrorString = "#DIV/0!";
				var nameErrorString = "#NAME?";
				var naErrorString = "#NV";
				var refErrorString = "#BEZUG!";
				var nullErrorString = "#NULL!";
				bool isErrorValue = false;
				ExcelErrorValue error = null;
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Value, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(numErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Num, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(div0ErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Div0, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nameErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Name, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(naErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.NA, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(refErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Ref, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nullErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Null, error.Type);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationaliztionUtilityParsesPolishErrorValues()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("pl-pl");
				var valueErrorString = "#ARG!";
				var numErrorString = "#LICZBA!";
				var div0ErrorString = "#DZIEL/0!";
				var nameErrorString = "#NAZWA?";
				var naErrorString = "#N/D!";
				var refErrorString = "#ADR!";
				var nullErrorString = "#ZERO!";
				bool isErrorValue = false;
				ExcelErrorValue error = null;
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Value, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(numErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Num, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(div0ErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Div0, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nameErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Name, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(naErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.NA, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(refErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Ref, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nullErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Null, error.Type);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationaliztionUtilityParsesRussianErrorValues()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("ru-ru");
				var valueErrorString = "#ЗНАЧ!";
				var numErrorString = "#ЧИСЛО!";
				var div0ErrorString = "#ДЕЛ/0!";
				var nameErrorString = "#ИМЯ?";
				var naErrorString = "#Н/Д";
				var refErrorString = "#ССЫЛКА!";
				var nullErrorString = "#ПУСТО!";
				bool isErrorValue = false;
				ExcelErrorValue error = null;
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Value, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(numErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Num, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(div0ErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Div0, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nameErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Name, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(naErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.NA, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(refErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Ref, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nullErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Null, error.Type);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationaliztionUtilityParsesGreekErrorValues()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("el-gr");
				var valueErrorString = "#ΤΙΜΗ!";
				var numErrorString = "#ΑΡΙΘ!";
				var div0ErrorString = "#ΔΙΑΙΡ./0!";
				var nameErrorString = "#ΟΝΟΜΑ?";
				var naErrorString = "#Δ/Υ";
				var refErrorString = "#ΑΝΑΦ!";
				var nullErrorString = "#ΚΕΝΟ!";
				bool isErrorValue = false;
				ExcelErrorValue error = null;
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Value, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(numErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Num, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(div0ErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Div0, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nameErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Name, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(naErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.NA, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(refErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Ref, error.Type);
				isErrorValue = InternationalizationUtility.TryParseLocalErrorValue(nullErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isErrorValue, true);
				Assert.AreEqual(eErrorType.Null, error.Type);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationalizationUtilityParsesEnglishBooleanStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-us");
				var trueString = "tRuE";
				var falseString = "fAlSe";
				var isBoolean = InternationalizationUtility.TryParseLocalBoolean(trueString, CultureInfo.CurrentCulture, out bool booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, true);
				isBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, false);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationalizationUtilityParsesGermanBooleanStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-de");
				var trueString = "wAhR";
				var falseString = "fAlScH";
				var isBoolean = InternationalizationUtility.TryParseLocalBoolean(trueString, CultureInfo.CurrentCulture, out bool booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, true);
				isBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, false);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationalizationUtilityParsesPolishBooleanStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("pl-pl");
				var trueString = "PRAWDA";
				var falseString = "FAŁSZ";
				var isBoolean = InternationalizationUtility.TryParseLocalBoolean(trueString, CultureInfo.CurrentCulture, out bool booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, true);
				isBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, false);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationalizationUtilityParsesRussianBooleanStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("ru-ru");
				var trueString = "ИСТИНА";
				var falseString = "ЛОЖЬ";
				var isBoolean = InternationalizationUtility.TryParseLocalBoolean(trueString, CultureInfo.CurrentCulture, out bool booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, true);
				isBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, false);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationalizationUtilityParsesGreekBooleanStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("el-gr");
				var trueString = "TRUE";
				var falseString = "FALSE";
				var isBoolean = InternationalizationUtility.TryParseLocalBoolean(trueString, CultureInfo.CurrentCulture, out bool booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, true);
				isBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out booleanValue);
				Assert.AreEqual(isBoolean, true);
				Assert.AreEqual(booleanValue, false);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void InternationalizationUtilityWithSupportedLanguageCodeAndNonSupportedCultureCode()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("es-mx"); // Spanish-Mexico
				var valueErrorString = "#¡VALOR!";
				var isValueError = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out ExcelErrorValue error);
				Assert.AreEqual(isValueError, true);
				Assert.AreEqual(error.Type, eErrorType.Value);
				var falseString = "FALSO";
				var isFalseBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out bool falseBool);
				Assert.AreEqual(isFalseBoolean, true);
				Assert.AreEqual(falseBool, false);
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-at"); // German-Austria
				valueErrorString = "#WERT!";
				isValueError = InternationalizationUtility.TryParseLocalErrorValue(valueErrorString, CultureInfo.CurrentCulture, out error);
				Assert.AreEqual(isValueError, true);
				Assert.AreEqual(error.Type, eErrorType.Value);
				falseString = "FALSCH";
				isFalseBoolean = InternationalizationUtility.TryParseLocalBoolean(falseString, CultureInfo.CurrentCulture, out falseBool);
				Assert.AreEqual(isFalseBoolean, true);
				Assert.AreEqual(falseBool, false);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion
	}
}
