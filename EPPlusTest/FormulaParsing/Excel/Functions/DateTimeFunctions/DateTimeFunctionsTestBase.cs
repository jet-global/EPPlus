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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public abstract class DateTimeFunctionsTestBase
	{
		#region Properties
		protected ParsingContext ParsingContext { get; } = ParsingContext.Create();
		#endregion

		#region Protected Methods
		/// <summary>
		/// Manually constructs an OADate for Today that represents the specified time.
		/// </summary>
		/// <param name="hour">The hour to construct a time from.</param>
		/// <param name="minute">The minute to construct a time from.</param>
		/// <param name="second">The second to construct a time from.</param>
		/// <returns>The OADate for today at the specified time.</returns>
		protected double GetTime(int hour, int minute, int second)
		{
			var secInADay = DateTime.Today.AddDays(1).Subtract(DateTime.Today).TotalSeconds;
			var secondsOfExample = (double)(hour * 60 * 60 + minute * 60 + second);
			return secondsOfExample / secInADay;
		}
		#endregion
	}
}
