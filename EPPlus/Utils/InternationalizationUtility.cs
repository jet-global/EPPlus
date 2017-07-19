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
using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.Utils
{
	class InternationalizationUtility
	{
		#region Localized Boolean Dictionaries
		private static readonly Dictionary<string, string> trueStrings = new Dictionary<string, string>()
		{
			{"en", "TRUE"},			// English
			{"de", "WAHR"},			// German
			{"zh", "TRUE"},			// Chinese
			{"da", "SAND"},			// Danish
			{"nl", "WAAR"},			// Dutch
			{"fi", "TOSI"},			// Finnish
			{"fr", "VRAI"},			// French
			{"it", "VERO"},			// Italian
			{"ja", "TRUE"},			// Japanese
			{"ko", "TRUE"},			// Korean
			{"nb", "SANN"},			// Norwegian
			{"pl", "PRAWDA"},		// Polish
			{"pt", "VERDADEIRO"},	// Portuguese
			{"ru", "ИСТИНА"},		// Russian
			{"es", "VERDADERO"},	// Spanish
			{"sv", "SANT"},			// Swedish
			{"hr", "TRUE"},			// Croatian
			{"cs", "PRAVDA"},		// Czech
			{"el", "TRUE"},			// Greek
			{"hu", "IGAZ"},			// Hungarian
			{"ms", "TRUE"},			// Malay
			{"ro", "TRUE"},			// Romanian
			{"sk", "TRUE"},			// Slovak
			{"sl", "TRUE"},			// Slovenian
			{"tr", "DOĞRU"}			// Turkish
		};

		private static readonly Dictionary<string, string> falseStrings = new Dictionary<string, string>()
		{
			{"en", "FALSE"},		// English
			{"de", "FALSCH"},		// German
			{"zh", "FALSE"},		// Chinese
			{"da", "FALSK"},		// Danish
			{"nl", "ONWAAR"},		// Dutch
			{"fi", "EPÄTOSI"},		// Finnish
			{"fr", "FAUX"},			// French
			{"it", "FALSO"},		// Italian
			{"ja", "FALSE"},		// Japanese
			{"ko", "FALSE"},		// Korean
			{"nb", "USANN"},		// Norwegian
			{"pl", "FAŁSZ"},		// Polish
			{"pt", "FALSO"},		// Portuguese
			{"ru", "ЛОЖЬ"},			// Russian
			{"es", "FALSO"},		// Spanish
			{"sv", "FALSKT"},		// Swedish
			{"hr", "FALSE"},		// Croatian
			{"cs", "NEPRAVDA"},		// Czech
			{"el", "FALSE"},		// Greek
			{"hu", "HAMIS"},		// Hungarian
			{"ms", "FALSE"},		// Malay
			{"ro", "FALSE"},		// Romanian
			{"sk", "FALSE"},		// Slovak
			{"sl", "FALSE"},		// Slovenian
			{"tr", "YANLIŞ" }		// Turkish
		};
		#endregion

		#region Localized Error Dictionaries
		private static readonly Dictionary<string, eErrorType> EnglishErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> GermanErrors = new Dictionary<string, eErrorType>()
		{
			{"#WERT!", eErrorType.Value},
			{"#ZAHL!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#NV", eErrorType.NA},
			{"#BEZUG!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> ChineseErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> DanishErrors = new Dictionary<string, eErrorType>()
		{
			{"#VÆRDI!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIVISION/0!", eErrorType.Div0},
			{"#NAVN?", eErrorType.Name},
			{"#I/T", eErrorType.NA},
			{"#REFERENCE!", eErrorType.Ref},
			{"#NUL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> DutchErrors = new Dictionary<string, eErrorType>()
		{
			{"#WAARDE!", eErrorType.Value},
			{"#GETAL!", eErrorType.Num},
			{"#DEEL/0!", eErrorType.Div0},
			{"#NAAM?", eErrorType.Name},
			{"#N/B", eErrorType.NA},
			{"#VERW!", eErrorType.Ref},
			{"#LEEG!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> FinnishErrors = new Dictionary<string, eErrorType>()
		{
			{"#ARVO!", eErrorType.Value},
			{"#LUKU!", eErrorType.Num},
			{"#JAKO/0!", eErrorType.Div0},
			{"#NIMI?", eErrorType.Name},
			{"#PUUTTUU!", eErrorType.NA},
			{"#VIITTAUS!", eErrorType.Ref},
			{"#TYHJÄ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> FrenchErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALEUR!", eErrorType.Value},
			{"#NOMBRE!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NOM?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NUL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> ItalianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALORE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NOME?", eErrorType.Name},
			{"#N/D", eErrorType.NA},
			{"#RIF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> JapaneseErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> KoreanErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> NorwegianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VERDI!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAVN?", eErrorType.Name},
			{"#I/T", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> PolishErrors = new Dictionary<string, eErrorType>()
		{
			{"#ARG!", eErrorType.Value},
			{"#LICZBA!", eErrorType.Num},
			{"#DZIEL/0!", eErrorType.Div0},
			{"#NAZWA?", eErrorType.Name},
			{"#N/D!", eErrorType.NA},
			{"#ADR!", eErrorType.Ref},
			{"#ZERO!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> PortugueseErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALOR!", eErrorType.Value},
			{"#NÚM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NOME?", eErrorType.Name},
			{"#N/D", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULO!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> RussianErrors = new Dictionary<string, eErrorType>()
		{
			{"#ЗНАЧ!", eErrorType.Value},
			{"#ЧИСЛО!", eErrorType.Num},
			{"#ДЕЛ/0!", eErrorType.Div0},
			{"#ИМЯ?", eErrorType.Name},
			{"#Н/Д", eErrorType.NA},
			{"#ССЫЛКА!", eErrorType.Ref},
			{"#ПУСТО!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> SpanishErrors = new Dictionary<string, eErrorType>()
		{
			{"#¡VALOR!", eErrorType.Value},
			{"#¡NUM!", eErrorType.Num},
			{"#¡DIV/0!", eErrorType.Div0},
			{"#¿NOMBRE?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#¡REF!", eErrorType.Ref},
			{"#¡NULO!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> SwedishErrors = new Dictionary<string, eErrorType>()
		{
			{"#VÄRDEFEL!", eErrorType.Value},
			{"#OGILTIGT!", eErrorType.Num},
			{"#DIVISION/0!", eErrorType.Div0},
			{"#NAMN?", eErrorType.Name},
			{"#SAKNAS!", eErrorType.NA},
			{"#REFERENS!", eErrorType.Ref},
			{"#SKÄRNING!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> CroatianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VRIJEDNOST!", eErrorType.Value},
			{"#BROJ!", eErrorType.Num},
			{"#DIJ/0!", eErrorType.Div0},
			{"#NAZIV?", eErrorType.Name},
			{"#N/D", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> CzechErrors = new Dictionary<string, eErrorType>()
		{
			{"#HODNOTA!", eErrorType.Value},
			{"#ČÍSLO!", eErrorType.Num},
			{"#DĚLENÍ_NULOU!", eErrorType.Div0},
			{"#NÁZEV?", eErrorType.Name},
			{"#NENÍ_K_DISPOZICI", eErrorType.NA},
			{"#ODKAZ!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> GreekErrors = new Dictionary<string, eErrorType>()
		{
			{"#ΤΙΜΗ!", eErrorType.Value},
			{"#ΑΡΙΘ!", eErrorType.Num},
			{"#ΔΙΑΙΡ./0!", eErrorType.Div0},
			{"#ΟΝΟΜΑ?", eErrorType.Name},
			{"#Δ/Υ", eErrorType.NA},
			{"#ΑΝΑΦ!", eErrorType.Ref},
			{"#ΚΕΝΟ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> HungarianErrors = new Dictionary<string, eErrorType>()
		{
			{"#ÉRTÉK!", eErrorType.Value},
			{"#SZÁM!", eErrorType.Num},
			{"#ZÉRÓOSZTÓ!", eErrorType.Div0},
			{"#NÉV?", eErrorType.Name},
			{"#HIÁNYZIK", eErrorType.NA},
			{"#HIV!", eErrorType.Ref},
			{"#NULLA!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> MalayErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> RomanianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALOARE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NUME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NUL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> SlovakErrors = new Dictionary<string, eErrorType>()
		{
			{"#HODNOTA!", eErrorType.Value},
			{"#ČÍSLO!", eErrorType.Num},
			{"#DELENIENULOU!", eErrorType.Div0},
			{"#NÁZOV?", eErrorType.Name},
			{"#NEDOSTUPNÝ", eErrorType.NA},
			{"#ODKAZ!", eErrorType.Ref},
			{"#NEPLATNÝ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> SlovenianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VREDN!", eErrorType.Value},
			{"#ŠTEV!", eErrorType.Num},
			{"#DEL/0!", eErrorType.Div0},
			{"#IME?", eErrorType.Name},
			{"#N/V", eErrorType.NA},
			{"#SKLIC!", eErrorType.Ref},
			{"#NIČ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> TurkishErrors = new Dictionary<string, eErrorType>()
		{
			{"#DEĞER!", eErrorType.Value},
			{"#SAYI!", eErrorType.Num},
			{"#SAYI/0!", eErrorType.Div0},
			{"#AD?", eErrorType.Name},
			{"#YOK", eErrorType.NA},
			{"#BAŞV!", eErrorType.Ref},
			{"#BOŞ!", eErrorType.Null}
		};
		#endregion

		private static readonly Dictionary<string, Dictionary<string, eErrorType>> errorDictionaries = new Dictionary<string, Dictionary<string, eErrorType>>()
		{
			{"en", EnglishErrors},		// English
			{"de", GermanErrors},		// German
			{"zh", ChineseErrors},		// Chinese
			{"da", DanishErrors},		// Danish
			{"nl", DutchErrors},		// Dutch
			{"fi", FinnishErrors},		// Finnish
			{"fr", FrenchErrors},		// French
			{"it", ItalianErrors},		// Italian
			{"ja", JapaneseErrors},		// Japanese
			{"ko", KoreanErrors},		// Korean
			{"nb", NorwegianErrors},	// Norwegian
			{"pl", PolishErrors},		// Polish
			{"pt", PortugueseErrors},	// Portuguese
			{"ru", RussianErrors},		// Russian
			{"es", SpanishErrors},		// Spanish
			{"sv", SwedishErrors},		// Swedish
			{"hr", CroatianErrors},		// Croatian
			{"cs", CzechErrors},		// Czech
			{"el", GreekErrors},		// Greek
			{"hu", HungarianErrors},	// Hungarian
			{"ms", MalayErrors},		// Malay
			{"ro", RomanianErrors},		// Romanian
			{"sk", SlovakErrors},		// Slovak
			{"sl", SlovenianErrors},	// Slovenian
			{"tr", TurkishErrors}		// Turkish
		};

		/// <summary>
		/// Attempt to parse the contents of <paramref name="errorCandidate"/> to an
		/// <see cref="ExcelErrorValue"/>. The <paramref name="errorCandidate"/> is compared
		/// to the error value strings for the given <paramref name="culture"/>.
		/// </summary>
		/// <param name="errorCandidate">The string to parse to an <see cref="ExcelErrorValue"/>.</param>
		/// <param name="culture">
		///		The <see cref="CultureInfo"/> that determines which error strings the 
		///		<paramref name="errorCandidate"/> are compared against.</param>
		/// <param name="errorValue">The resulting <see cref="ExcelErrorValue"/> from successfully parsing the <paramref name="errorCandidate"/>.</param>
		/// <returns>Returns true if the <paramref name="errorCandidate"/> was parsed to an <see cref="ExcelErrorValue"/>, and false otherwise.</returns>
		public static bool TryParseLocalErrorValue(string errorCandidate, CultureInfo culture, out ExcelErrorValue errorValue)
		{
			errorValue = null;
			errorCandidate = errorCandidate.ToUpper(culture);
			Dictionary<string, eErrorType> errorDictionary = null;
			if (!errorDictionaries.TryGetValue(culture.TwoLetterISOLanguageName, out errorDictionary))
				return false;
			if (!errorDictionary.TryGetValue(errorCandidate, out eErrorType errorType))
				return false;
			errorValue = ExcelErrorValue.Create(errorType);
			return true;
		}

		/// <summary>
		/// Attempt to parse the contents of <paramref name="booleanCandidate"/> to a
		/// bool. The <paramref name="booleanCandidate"/> is compared to the boolean value
		/// strings for the given <paramref name="culture"/>.
		/// </summary>
		/// <param name="booleanCandidate">The string to parse to a bool.</param>
		/// <param name="culture">
		///		The <see cref="CultureInfo"/> that determines which boolean value strings
		///		the <paramref name="booleanCandidate"/> are compared against.</param>
		/// <param name="booleanValue">The resulting <see cref="ExcelErrorValue"/> from successfully parsing the <paramref name="booleanCandidate"/>.</param>
		/// <returns>Returns true if the <paramref name="booleanCandidate"/> was parsed to a bool, and false otherwise.</returns>
		public static bool TryParseLocalBoolean(string booleanCandidate, CultureInfo culture, out bool booleanValue)
		{
			booleanValue = false;
			var cultureLanguageCode = culture.TwoLetterISOLanguageName;
			booleanCandidate = booleanCandidate.ToUpper(culture);
			if (trueStrings.TryGetValue(cultureLanguageCode, out string localTrueString) && booleanCandidate.Equals(localTrueString))
				booleanValue = true;
			else if (falseStrings.TryGetValue(cultureLanguageCode, out string localFalseString) && booleanCandidate.Equals(localFalseString))
				booleanValue = false;
			else
				return false;
			return true;
		}
	}
}
