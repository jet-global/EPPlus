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
*  * Author							Change						Date
* *******************************************************************************
* * Mats Alm   		                Added		                2013-12-03
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class converts a time into a serial number.
	/// </summary>
	public abstract class TimeBaseFunction : ExcelFunction
	{
		/// <summary>
		/// GetTimeSerialNumber returns the time as a serial number.
		/// </summary>
		/// <param name="seconds">The number of seconds.</param>
		/// <returns>The time as a serial number.</returns>
		protected double GetTimeSerialNumber(double seconds)
		{
			var SecondsInADay = 24 * 60 * 60;
			return seconds / SecondsInADay;
		}
	}
}