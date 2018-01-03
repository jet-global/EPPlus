﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * Author							Change						Date
 * ******************************************************************************
 * Eyal Seagull        Added       		  2012-04-03
 *******************************************************************************/
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// ExcelConditionalFormattingNextMonth
	/// </summary>
	public class ExcelConditionalFormattingNextMonth
	  : ExcelConditionalFormattingTimePeriodGroup
	{
		/****************************************************************************************/

		#region Constructors
		/// <summary>
		/// 
		/// </summary>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="itemElementNode"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingNextMonth(
		  ExcelAddressBase address,
		  int priority,
		  ExcelWorksheet worksheet,
		  XmlNode itemElementNode,
		  XmlNamespaceManager namespaceManager)
		  : base(
			 eExcelConditionalFormattingRuleType.NextMonth,
			 address,
			 priority,
			 worksheet,
			 itemElementNode,
			 (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
		{
			if (itemElementNode == null) //Set default values and create attributes if needed
			{
				TimePeriod = eExcelConditionalFormattingTimePeriodType.NextMonth;
				Formula = string.Format(
				  "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)), YEAR({0})=YEAR(EDATE(TODAY(),0+1)))",
				  Address.Start.Address);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="itemElementNode"></param>
		internal ExcelConditionalFormattingNextMonth(
		  ExcelAddressBase address,
		  int priority,
		  ExcelWorksheet worksheet,
		  XmlNode itemElementNode)
		  : this(
			 address,
			 priority,
			 worksheet,
			 itemElementNode,
			 null)
		{
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		internal ExcelConditionalFormattingNextMonth(
		  ExcelAddressBase address,
		  int priority,
		  ExcelWorksheet worksheet)
		  : this(
			 address,
			 priority,
			 worksheet,
			 null,
			 null)
		{
		}
		#endregion Constructors

		/****************************************************************************************/
	}
}