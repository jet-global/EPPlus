/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Filters;

namespace EPPlusTest.Table.PivotTable.Filters
{
	[TestClass]
	public class ExcelPivotTableAdvancedFilterTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ExcelPivotTableAdvancedFilterConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<filter fld=""2"" type=""captionNotEqual"" evalOrder=""-1"" id=""2"" stringValue1=""march"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
								<autoFilter ref=""A1"">
									<filterColumn colId=""0"">
										<customFilters>
											<customFilter operator=""notEqual"" val=""march""/>
										</customFilters>
									</filterColumn>
								</autoFilter>
							</filter>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var advancedFilter = new ExcelPivotTableAdvancedFilter(namespaceManager, document.SelectSingleNode("//d:filter", namespaceManager));
			Assert.IsNotNull(advancedFilter);
			Assert.AreEqual(2, advancedFilter.Field);
			Assert.AreEqual("captionNotEqual", advancedFilter.PivotFilterType);
			Assert.AreEqual("march", advancedFilter.StringValueOne);
			Assert.AreEqual(FieldFilter.Label, advancedFilter.FieldFilterType);
			Assert.IsNotNull(advancedFilter.CustomFilters);
			Assert.IsNull(advancedFilter.Filters);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelPivotTableAdvancedFilterConstructorNullNamespaceManagerTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<filter fld=""2"" type=""captionNotEqual"" evalOrder=""-1"" id=""2"" stringValue1=""march"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
								<autoFilter ref=""A1"">
									<filterColumn colId=""0"">
										<customFilters>
											<customFilter operator=""notEqual"" val=""march""/>
										</customFilters>
									</filterColumn>
								</autoFilter>
							</filter>");
			new ExcelPivotTableAdvancedFilter(null, document.SelectSingleNode("//d:filter", TestUtility.CreateDefaultNSM()));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelPivotTableAdvancedFilterConstructorNullNodeTest()
		{
			new ExcelPivotTableAdvancedFilter(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion

		#region MatchesFilterCriteriaResult Tests
		[TestMethod]
		public void MatchesFilterCriteriaResult()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<filter fld=""2"" type=""captionNotEqual"" evalOrder=""-1"" id=""2"" stringValue1=""march"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
								<autoFilter ref=""A1"">
									<filterColumn colId=""0"">
										<customFilters>
											<customFilter operator=""notEqual"" val=""march""/>
										</customFilters>
									</filterColumn>
								</autoFilter>
							</filter>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var advancedFilter = new ExcelPivotTableAdvancedFilter(namespaceManager, document.SelectSingleNode("//d:filter", namespaceManager));
			Assert.IsTrue(advancedFilter.MatchesFilterCriteriaResult("January", false));
			document.LoadXml(@"<filter fld=""2"" type=""captionBetween"" evalOrder=""-1"" id=""2"" stringValue1=""c"" stringValue2=""o"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
								<autoFilter ref=""A1"">
									<filterColumn colId=""0"">
										<customFilters and=""1"">
											<customFilter operator=""greaterThanOrEqual"" val=""c""/>
											<customFilter operator=""lessThanOrEqual"" val=""o""/>
										</customFilters>
									</filterColumn>
								</autoFilter>
							</filter>");
			advancedFilter = new ExcelPivotTableAdvancedFilter(namespaceManager, document.SelectSingleNode("//d:filter", namespaceManager));
			Assert.IsFalse(advancedFilter.MatchesFilterCriteriaResult("Apples", false));
		}
		#endregion
	}
}
