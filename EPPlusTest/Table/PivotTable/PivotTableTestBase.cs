/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau, Evan Schallerer, and others as noted in the source history.
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
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	public abstract class PivotTableTestBase
	{
		#region Helper Methods
		public CacheFieldNode GetTestCacheFieldNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" name=""Item"" numFmtId=""0""><sharedItems count=""2""><s v=""Bike""/><s v=""Car""/></sharedItems></cacheField>");
			var ns = TestUtility.CreateDefaultNSM();
			return new CacheFieldNode(ns, document.SelectSingleNode("//d:cacheField", ns));
		}

		public CacheRecordNode GetTestCacheRecordNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1""><r><n v=""20100076""/><x v=""0""/></r></pivotCacheRecords>");
			var ns = TestUtility.CreateDefaultNSM();
			return new CacheRecordNode(ns, document.SelectSingleNode("//d:r", ns));
		}
		#endregion
	}
}
