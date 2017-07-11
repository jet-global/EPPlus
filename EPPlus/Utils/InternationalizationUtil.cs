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
		private static string defaultValueError = "#VALUE!";
		private static string defaultNumError = "#NUM!";
		private static string defaultDiv0Error = "#DIV/0!";
		private static string defaultNameError = "#NAME?";
		private static string defaultNAError = "#N/A";
		private static string defaultRefError = "#REF!";
		private static string defaultNullError = "#NULL!";

		private static readonly Dictionary<CultureInfo, string> valueErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultValueError}, // English
			{CultureInfo.CreateSpecificCulture("de-de"), "#WERT!"},			 // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultValueError}, // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultValueError}, // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "#VÆRDI!"},		 // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#WAARDE!"},		 // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#ARVO!"},			 // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), "#VALEUR!"},		 // French
			{CultureInfo.CreateSpecificCulture("it-it"), "#VALORE!"},		 // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultValueError}, // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultValueError}, // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), "#VERDI!"},		 // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#ARG!"},			 // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "#VALOR!"},		 // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "#VALOR!"},		 // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#ЗНАЧ!"},			 // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "#¡VALOR!"},		 // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#VÄRDEFEL!"},		 // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "#VRIJEDNOST!"},	 // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "#HODNOTA!"},		 // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#ΤΙΜΗ!"},			 // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#ÉRTÉK!"},		 // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultValueError}, // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), "#VALOARE!"},		 // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#HODNOTA!"},		 // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#VREDN!"},		 // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#DEĞER!"}			 // Turkish
		};

		private static readonly Dictionary<CultureInfo, string> numErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultNumError}, // English
			{CultureInfo.CreateSpecificCulture("de-de"), "#ZAHL!"},		   // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultNumError}, // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultNumError}, // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), defaultNumError}, // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#GETAL!"},	   // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#LUKU!"},		   // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), "#NOMBRE!"},	   // French
			{CultureInfo.CreateSpecificCulture("it-it"), defaultNumError}, // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultNumError}, // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultNumError}, // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), defaultNumError}, // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#LICZBA!"},	   // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "#NÚM!"},		   // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "#NÚM!"},		   // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#ЧИСЛО!"},	   // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "#¡NUM!"},		   // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#OGILTIGT!"},	   // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "#BROJ!"},		   // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "#ČÍSLO!"},	   // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#ΑΡΙΘ!"},		   // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#SZÁM!"},		   // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultNumError}, // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), defaultNumError}, // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#ČÍSLO!"},	   // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#ŠTEV!"},		   // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#SAYI!"}		   // Turkish
		};

		private static readonly Dictionary<CultureInfo, string> div0ErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultDiv0Error}, // English
			{CultureInfo.CreateSpecificCulture("de-de"), defaultDiv0Error}, // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultDiv0Error}, // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultDiv0Error}, // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "#DIVISION/0!"},   // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#DEEL/0!"},	    // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#JAKO/0!"},	    // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), defaultDiv0Error}, // French
			{CultureInfo.CreateSpecificCulture("it-it"), defaultDiv0Error}, // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultDiv0Error}, // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultDiv0Error}, // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), defaultDiv0Error}, // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#DZIEL/0!"},	    // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), defaultDiv0Error}, // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), defaultDiv0Error}, // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#ДЕЛ/0!"},	    // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "#¡DIV/0!"},	    // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#DIVISION/0!"},   // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "#DIJ/0!"},	    // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "#DĚLENÍ_NULOU!"}, // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#ΔΙΑΙΡ./0!"},	    // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#ZÉRÓOSZTÓ!"},    // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultDiv0Error}, // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), defaultDiv0Error}, // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#DELENIENULOU!"}, // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#DEL/0!"},	    // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#SAYI/0!"}	    // Turkish
		};

		private static readonly Dictionary<CultureInfo, string> nameErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultNameError}, // English
			{CultureInfo.CreateSpecificCulture("de-de"), defaultNameError}, // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultNameError}, // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultNameError}, // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "#NAVN?"},		    // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#NAAM?"},		    // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#NIMI?"},		    // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), "#NOM?"},		    // French
			{CultureInfo.CreateSpecificCulture("it-it"), "#NOME?"},		    // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultNameError}, // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultNameError}, // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), "#NAVN?"},		    // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#NAZWA?"},	    // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "#NOME?"},		    // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "#NOME?"},		    // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#ИМЯ?"},		    // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "#¿NOMBRE?"},	    // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#NAMN?"},		    // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "#NAZIV?"},	    // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "#NÁZEV?"},	    // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#ΟΝΟΜΑ?"},	    // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#NÉV?"},		    // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultNameError}, // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), "#NUME?"},		    // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#NÁZOV?"},	    // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#IME?"},		    // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#AD?"}		    // Turkish
		};

		private static readonly Dictionary<CultureInfo, string> naErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultNAError},	   // English
			{CultureInfo.CreateSpecificCulture("de-de"), "#NV"},			   // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultNAError},	   // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultNAError},	   // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "#I/T"},			   // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#N/B"},			   // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#PUUTTUU!"},		   // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), defaultNAError},	   // French
			{CultureInfo.CreateSpecificCulture("it-it"), "#N/D"},			   // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultNAError},	   // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultNAError},	   // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), "#I/T"},			   // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#N/D!"},			   // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "#N/D"},			   // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "#N/D"},			   // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#Н/Д"},			   // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), defaultNAError},	   // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#SAKNAS!"},		   // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), "#N/D"},			   // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "#NENÍ_K_DISPOZICI"}, // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#Δ/Υ"},			   // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#HIÁNYZIK"},		   // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultNAError},	   // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), defaultNAError},	   // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#NEDOSTUPNÝ"},	   // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#N/V"},			   // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#YOK"}			   // Turkish
		};

		private static readonly Dictionary<CultureInfo, string> refErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultRefError}, // English
			{CultureInfo.CreateSpecificCulture("de-de"), "#BEZUG!"},	   // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultRefError}, // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultRefError}, // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "#REFERENCE!"},   // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#VERW!"},		   // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#VIITTAUS!"},	   // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), defaultRefError}, // French
			{CultureInfo.CreateSpecificCulture("it-it"), "#RIF!"},		   // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultRefError}, // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultRefError}, // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), defaultRefError}, // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#ADR!"},		   // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), defaultRefError}, // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), defaultRefError}, // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#ССЫЛКА!"},	   // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "#¡REF!"},		   // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#REFERENS!"},	   // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), defaultRefError}, // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), "#ODKAZ!"},	   // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#ΑΝΑΦ!"},		   // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#HIV!"},		   // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultRefError}, // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), defaultRefError}, // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#ODKAZ!"},	   // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#SKLIC!"},	   // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#BAŞV!"}		   // Turkish
		};

		private static readonly Dictionary<CultureInfo, string> nullErrorStrings = new Dictionary<CultureInfo, string>()
		{
			{CultureInfo.CreateSpecificCulture("en-us"), defaultNullError}, // English
			{CultureInfo.CreateSpecificCulture("de-de"), defaultNullError}, // German
			{CultureInfo.CreateSpecificCulture("zh-tw"), defaultNullError}, // Chinese (Traditional)
			{CultureInfo.CreateSpecificCulture("zh-cn"), defaultNullError}, // Chinese (Simplified)
			{CultureInfo.CreateSpecificCulture("da-dk"), "#NUL!"},		    // Danish
			{CultureInfo.CreateSpecificCulture("nl-nl"), "#LEEG!"},		    // Dutch
			{CultureInfo.CreateSpecificCulture("fi-fi"), "#TYHJÄ!"},	    // Finnish
			{CultureInfo.CreateSpecificCulture("fr-fr"), "#NUL!"},		    // French
			{CultureInfo.CreateSpecificCulture("it-it"), defaultNullError}, // Italian
			{CultureInfo.CreateSpecificCulture("ja-jp"), defaultNullError}, // Japanese
			{CultureInfo.CreateSpecificCulture("ko-kr"), defaultNullError}, // Korean
			{CultureInfo.CreateSpecificCulture("nb-no"), defaultNullError}, // Norwegian
			{CultureInfo.CreateSpecificCulture("pl-pl"), "#ZERO!"},		    // Polish
			{CultureInfo.CreateSpecificCulture("pt-pt"), "#NULO!"},		    // Portuguese (Portugal)
			{CultureInfo.CreateSpecificCulture("pt-br"), "#NULO!"},		    // Portuguese (Brazil)
			{CultureInfo.CreateSpecificCulture("ru-ru"), "#ПУСТО!"},	    // Russian
			{CultureInfo.CreateSpecificCulture("es-es"), "#¡NULO!"},	    // Spanish (Spain)
			{CultureInfo.CreateSpecificCulture("sv-se"), "#SKÄRNING!"},	    // Swedish
			{CultureInfo.CreateSpecificCulture("hr-hr"), defaultNullError}, // Croatian
			{CultureInfo.CreateSpecificCulture("cs-cz"), defaultNullError}, // Czech
			{CultureInfo.CreateSpecificCulture("el-gr"), "#ΚΕΝΟ!"},		    // Greek
			{CultureInfo.CreateSpecificCulture("hu-hu"), "#NULLA!"},	    // Hungarian
			{CultureInfo.CreateSpecificCulture("ms-my"), defaultNullError}, // Malay
			{CultureInfo.CreateSpecificCulture("ro-ro"), "#NUL!"},		    // Romanian
			{CultureInfo.CreateSpecificCulture("sk-sk"), "#NEPLATNÝ!"},	    // Slovak
			{CultureInfo.CreateSpecificCulture("sl-si"), "#NIČ!"},		    // Slovenian
			{CultureInfo.CreateSpecificCulture("tr-tr"), "#BOŞ!"}		    // Turkish
		};

		public static bool TryParseLocalErrorValue(string errorCandidate, CultureInfo culture, out ExcelErrorValue errorValue)
		{
			errorValue = null;
			string errorString = null;
			if (valueErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.Value);
			else if (numErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.Num);
			else if (div0ErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.Div0);
			else if (nameErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.Name);
			else if (naErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.NA);
			else if (refErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.Ref);
			else if (nullErrorStrings.TryGetValue(culture, out errorString) && errorString.Equals(errorCandidate))
				errorValue = ExcelErrorValue.Create(eErrorType.Null);
			else
				return false;
			return true;
		}

		public static bool TryParseLocalBoolean(string booleanCandidate, CultureInfo culture, out bool booleanValue)
		{
			booleanValue = false;
			
			return false;
		}
	}
}
