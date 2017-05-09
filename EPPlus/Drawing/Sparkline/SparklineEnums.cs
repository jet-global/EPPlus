/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * SparklineEnums.cs Copyright (C) 2016 Matt Delaney.
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
 * Code change notes:
 * 
 * Author					Change						                Date
 * ******************************************************************************
 * Matt Delaney		        Sparklines                                2016-05-20
 *******************************************************************************/

namespace OfficeOpenXml.Drawing.Sparkline
{
	/// <summary>
	/// Represents the values of the ST_SparklineType type as defined at https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx.
	/// </summary>
	public enum SparklineType
	{
		Line, Column, Stacked
	}

	/// <summary>
	/// Represents the values of the ST_SparklineAxisMinMax type as defined at https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx
	/// </summary>
	public enum SparklineAxisMinMax
	{
		Individual, Group, Custom
	}

	/// <summary>
	/// Represents the values of the ST_DispBlanksAs type as defined at https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx.
	/// </summary>
	public enum DispBlanksAs
	{
		Span, Gap, Zero
	}
}
