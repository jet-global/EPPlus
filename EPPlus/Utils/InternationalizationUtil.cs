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
	class InternationalizationUtil
	{
		#region Language Error Dictionaries
		private static readonly Dictionary<string, eErrorType> englishErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> germanErrors = new Dictionary<string, eErrorType>()
		{
			{"#WERT!", eErrorType.Value},
			{"#ZAHL!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#NV", eErrorType.NA},
			{"#BEZUG!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> chineseErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> danishErrors = new Dictionary<string, eErrorType>()
		{
			{"#VÆRDI!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIVISION/0!", eErrorType.Div0},
			{"#NAVN?", eErrorType.Name},
			{"#I/T", eErrorType.NA},
			{"#REFERENCE!", eErrorType.Ref},
			{"#NUL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> dutchErrors = new Dictionary<string, eErrorType>()
		{
			{"#WAARDE!", eErrorType.Value},
			{"#GETAL!", eErrorType.Num},
			{"#DEEL/0!", eErrorType.Div0},
			{"#NAAM?", eErrorType.Name},
			{"#N/B", eErrorType.NA},
			{"#VERW!", eErrorType.Ref},
			{"#LEEG!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> finnishErrors = new Dictionary<string, eErrorType>()
		{
			{"#ARVO!", eErrorType.Value},
			{"#LUKU!", eErrorType.Num},
			{"#JAKO/0!", eErrorType.Div0},
			{"#NIMI?", eErrorType.Name},
			{"#PUUTTUU!", eErrorType.NA},
			{"#VIITTAUS!", eErrorType.Ref},
			{"#TYHJÄ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> frenchErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALEUR!", eErrorType.Value},
			{"#NOMBRE!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NOM?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NUL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> italianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALORE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NOME?", eErrorType.Name},
			{"#N/D", eErrorType.NA},
			{"#RIF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> japaneseErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> koreanErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> norwegianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VERDI!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAVN?", eErrorType.Name},
			{"#I/T", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> polishErrors = new Dictionary<string, eErrorType>()
		{
			{"#ARG!", eErrorType.Value},
			{"#LICZBA!", eErrorType.Num},
			{"#DZIEL/0!", eErrorType.Div0},
			{"#NAZWA?", eErrorType.Name},
			{"#N/D!", eErrorType.NA},
			{"#ADR!", eErrorType.Ref},
			{"#ZERO!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> portugueseErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALOR!", eErrorType.Value},
			{"#NÚM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NOME?", eErrorType.Name},
			{"#N/D", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULO!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> russianErrors = new Dictionary<string, eErrorType>()
		{
			{"#ЗНАЧ!", eErrorType.Value},
			{"#ЧИСЛО!", eErrorType.Num},
			{"#ДЕЛ/0!", eErrorType.Div0},
			{"#ИМЯ?", eErrorType.Name},
			{"#Н/Д", eErrorType.NA},
			{"#ССЫЛКА!", eErrorType.Ref},
			{"#ПУСТО!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> spanishErrors = new Dictionary<string, eErrorType>()
		{
			{"#¡VALOR!", eErrorType.Value},
			{"#¡NUM!", eErrorType.Num},
			{"#¡DIV/0!", eErrorType.Div0},
			{"#¿NOMBRE?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#¡REF!", eErrorType.Ref},
			{"#¡NULO!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> swedishErrors = new Dictionary<string, eErrorType>()
		{
			{"#VÄRDEFEL!", eErrorType.Value},
			{"#OGILTIGT!", eErrorType.Num},
			{"#DIVISION/0!", eErrorType.Div0},
			{"#NAMN?", eErrorType.Name},
			{"#SAKNAS!", eErrorType.NA},
			{"#REFERENS!", eErrorType.Ref},
			{"#SKÄRNING!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> croatianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VRIJEDNOST!", eErrorType.Value},
			{"#BROJ!", eErrorType.Num},
			{"#DIJ/0!", eErrorType.Div0},
			{"#NAZIV?", eErrorType.Name},
			{"#N/D", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> czechErrors = new Dictionary<string, eErrorType>()
		{
			{"#HODNOTA!", eErrorType.Value},
			{"#ČÍSLO!", eErrorType.Num},
			{"#DĚLENÍ_NULOU!", eErrorType.Div0},
			{"#NÁZEV?", eErrorType.Name},
			{"#NENÍ_K_DISPOZICI", eErrorType.NA},
			{"#ODKAZ!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> greekErrors = new Dictionary<string, eErrorType>()
		{
			{"#ΤΙΜΗ!", eErrorType.Value},
			{"#ΑΡΙΘ!", eErrorType.Num},
			{"#ΔΙΑΙΡ./0!", eErrorType.Div0},
			{"#ΟΝΟΜΑ?", eErrorType.Name},
			{"#Δ/Υ", eErrorType.NA},
			{"#ΑΝΑΦ!", eErrorType.Ref},
			{"#ΚΕΝΟ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> hungarianErrors = new Dictionary<string, eErrorType>()
		{
			{"#ÉRTÉK!", eErrorType.Value},
			{"#SZÁM!", eErrorType.Num},
			{"#ZÉRÓOSZTÓ!", eErrorType.Div0},
			{"#NÉV?", eErrorType.Name},
			{"#HIÁNYZIK", eErrorType.NA},
			{"#HIV!", eErrorType.Ref},
			{"#NULLA!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> malayErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALUE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NAME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NULL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> romanianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VALOARE!", eErrorType.Value},
			{"#NUM!", eErrorType.Num},
			{"#DIV/0!", eErrorType.Div0},
			{"#NUME?", eErrorType.Name},
			{"#N/A", eErrorType.NA},
			{"#REF!", eErrorType.Ref},
			{"#NUL!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> slovakErrors = new Dictionary<string, eErrorType>()
		{
			{"#HODNOTA!", eErrorType.Value},
			{"#ČÍSLO!", eErrorType.Num},
			{"#DELENIENULOU!", eErrorType.Div0},
			{"#NÁZOV?", eErrorType.Name},
			{"#NEDOSTUPNÝ", eErrorType.NA},
			{"#ODKAZ!", eErrorType.Ref},
			{"#NEPLATNÝ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> slovenianErrors = new Dictionary<string, eErrorType>()
		{
			{"#VREDN!", eErrorType.Value},
			{"#ŠTEV!", eErrorType.Num},
			{"#DEL/0!", eErrorType.Div0},
			{"#IME?", eErrorType.Name},
			{"#N/V", eErrorType.NA},
			{"#SKLIC!", eErrorType.Ref},
			{"#NIČ!", eErrorType.Null}
		};

		private static readonly Dictionary<string, eErrorType> turkishErrors = new Dictionary<string, eErrorType>()
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

		private static readonly Dictionary<CultureInfo, Dictionary<string, eErrorType>> errorDictionaries = new Dictionary<CultureInfo, Dictionary<string, eErrorType>>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), englishErrors},	// English
			{CultureInfo.CreateSpecificCulture("de-de"), germanErrors},		// German
			{CultureInfo.CreateSpecificCulture("zh-tw"), chineseErrors},	// Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), chineseErrors},	// Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), danishErrors},		// Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), dutchErrors},		// Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), finnishErrors},	// Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), frenchErrors},		// French
			{CultureInfo.CreateSpecificCulture("it-it"), italianErrors},	// Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), japaneseErrors},	// Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), koreanErrors},		// Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), norwegianErrors},	// Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), polishErrors},		// Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), portugueseErrors},	// Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), portugueseErrors},	// Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), russianErrors},	// Russian
			{CultureInfo.CreateSpecificCulture("es-es"), spanishErrors},	// Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), swedishErrors},	// Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), croatianErrors},	// Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), czechErrors},		// Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), greekErrors},		// Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), hungarianErrors},	// Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), malayErrors},		// Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), romanianErrors},	// Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), slovakErrors},		// Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), slovenianErrors},	// Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), turkishErrors}		// Turkish
		};

		private static readonly Dictionary<CultureInfo, string> trueStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), "TRUE"},		// English
			{CultureInfo.CreateSpecificCulture("de-de"), "WAHR"},		// German
			{CultureInfo.CreateSpecificCulture("zh-tw"), "TRUE"},		// Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), "TRUE"},		// Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "SAND"},		// Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "WAAR"},		// Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "TOSI"},		// Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), "VRAI"},		// French
			{CultureInfo.CreateSpecificCulture("it-it"), "VERO"},		// Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), "TRUE"},		// Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), "TRUE"},		// Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), "SANN"},		// Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "PRAWDA"},		// Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "VERDADEIRO"},	// Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "VERDADEIRO"},	// Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "ИСТИНА"},		// Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "VERDADERO"},	// Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "SANT"},		// Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "TRUE"},		// Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "PRAVDA"},		// Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "TRUE"},		// Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "IGAZ"},		// Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), "TRUE"},		// Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), "TRUE"},		// Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "TRUE"},		// Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "TRUE"},		// Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "DOĞRU"}		// Turkish
		};

		private static readonly Dictionary<CultureInfo, string> falseStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), "FALSE"},		// English
			{CultureInfo.CreateSpecificCulture("de-de"), "FALSCH"},		// German
			{CultureInfo.CreateSpecificCulture("zh-tw"), "FALSE"},		// Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), "FALSE"},		// Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "FALSK"},		// Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "ONWAAR"},		// Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "EPÄTOSI"},	// Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), "FAUX"},		// French
			{CultureInfo.CreateSpecificCulture("it-it"), "FALSO"},		// Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), "FALSE"},		// Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), "FALSE"},		// Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), "USANN"},		// Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "FAŁSZ"},		// Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "FALSO"},		// Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "FALSO"},		// Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "ЛОЖЬ"},		// Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "FALSO"},		// Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "FALSKT"},		// Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "FALSE"},		// Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "NEPRAVDA"},	// Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "FALSE"},		// Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "HAMIS"},		// Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), "FALSE"},		// Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), "FALSE"},		// Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "FALSE"},		// Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "FALSE"},		// Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "YANLIŞ" }		// Turkish
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
			if (!errorDictionaries.TryGetValue(culture, out errorDictionary))
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
			booleanCandidate = booleanCandidate.ToUpper(culture);
			if (trueStrings.TryGetValue(culture, out string localTrueString) && booleanCandidate.Equals(localTrueString))
				booleanValue = true;
			else if (falseStrings.TryGetValue(culture, out string localFalseString) && booleanCandidate.Equals(localFalseString))
				booleanValue = false;
			else
				return false;
			return true;
		}
	}
}
