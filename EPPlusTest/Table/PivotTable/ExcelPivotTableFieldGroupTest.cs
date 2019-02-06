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
using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotTableFieldGroupTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ExcelPivotTableFieldGroupConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3"">
				<discretePr count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
				</discretePr>
				<groupItems count=""2"">
					<s v=""Group1""/>
					<s v=""Group2""/>
				</groupItems>
			</fieldGroup>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var fieldGroup = new ExcelPivotTableFieldGroup(namespaceManager, document.SelectSingleNode("//d:fieldGroup", namespaceManager));
			Assert.IsNotNull(fieldGroup.GroupItems);
			Assert.AreEqual(2, fieldGroup.GroupItems.Count);
			Assert.IsNotNull(fieldGroup.DiscreteGroupingProperties);
			Assert.AreEqual(4, fieldGroup.DiscreteGroupingProperties.Count);
			Assert.AreEqual(PivotFieldDateGrouping.None, fieldGroup.GroupBy);
		}

		[TestMethod]
		public void ExcelPivotTableFieldGroupConstructorNullGroupItemsTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3"">
				<discretePr count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
				</discretePr>
			</fieldGroup>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var fieldGroup = new ExcelPivotTableFieldGroup(namespaceManager, document.SelectSingleNode("//d:fieldGroup", namespaceManager));
			Assert.IsNull(fieldGroup.GroupItems);
			Assert.IsNotNull(fieldGroup.DiscreteGroupingProperties);
			Assert.AreEqual(4, fieldGroup.DiscreteGroupingProperties.Count);
			Assert.AreEqual(PivotFieldDateGrouping.None, fieldGroup.GroupBy);
		}

		[TestMethod]
		public void ExcelPivotTableFieldGroupConstructorNullDiscreteGroupingPropertiesTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3"">
				<groupItems count=""2"">
					<s v=""Group1""/>
					<s v=""Group2""/>
				</groupItems>
			</fieldGroup>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var fieldGroup = new ExcelPivotTableFieldGroup(namespaceManager, document.SelectSingleNode("//d:fieldGroup", namespaceManager));
			Assert.IsNotNull(fieldGroup.GroupItems);
			Assert.AreEqual(2, fieldGroup.GroupItems.Count);
			Assert.IsNull(fieldGroup.DiscreteGroupingProperties);
			Assert.AreEqual(PivotFieldDateGrouping.None, fieldGroup.GroupBy);
		}

		[TestMethod]
		public void ExcelPivotTableFieldGroupConstructorNullCollectionTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3""></fieldGroup>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var fieldGroup = new ExcelPivotTableFieldGroup(namespaceManager, document.SelectSingleNode("//d:fieldGroup", namespaceManager));
			Assert.IsNull(fieldGroup.GroupItems);
			Assert.IsNull(fieldGroup.DiscreteGroupingProperties);
			Assert.AreEqual(PivotFieldDateGrouping.None, fieldGroup.GroupBy);
		}

		[TestMethod]
		public void ExcelPivotTableFieldGroupConstructorGroupByTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3"">
				<rangePr groupBy=""months""/>
				<groupItems count=""2"">
					<s v=""Group1""/>
					<s v=""Group2""/>
				</groupItems>
			</fieldGroup>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var fieldGroup = new ExcelPivotTableFieldGroup(namespaceManager, document.SelectSingleNode("//d:fieldGroup", namespaceManager));
			Assert.IsNotNull(fieldGroup.GroupItems);
			Assert.AreEqual(2, fieldGroup.GroupItems.Count);
			Assert.IsNull(fieldGroup.DiscreteGroupingProperties);
			Assert.AreEqual(PivotFieldDateGrouping.Months, fieldGroup.GroupBy);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelPivotTableFieldGroupConstructorNullNamespaceManagerTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3"">
				<discretePr count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
				</discretePr>
				<groupItems count=""2"">
					<s v=""Group1""/>
					<s v=""Group2""/>
				</groupItems>
			</fieldGroup>");
			new ExcelPivotTableFieldGroup(null, document.SelectSingleNode("//d:fieldGroup", TestUtility.CreateDefaultNSM()));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelPivotTableFieldGroupConstructorNullNodeTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<fieldGroup xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" base=""3"">
				<discretePr count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
				</discretePr>
				<groupItems count=""2"">
					<s v=""Group1""/>
					<s v=""Group2""/>
				</groupItems>
			</fieldGroup>");
			new ExcelPivotTableFieldGroup(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion
	}
}
