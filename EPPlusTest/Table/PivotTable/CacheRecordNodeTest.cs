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
using System.Linq;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheRecordNodeTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		public void CacheRecordNodeConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1""><r><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></r></pivotCacheRecords>");
			var ns = TestUtility.CreateDefaultNSM();
			var node = new CacheRecordNode(ns, document.SelectSingleNode("//d:r", ns));
			Assert.AreEqual(6, node.Items.Count);
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.b));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.x));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.d));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.e));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.m));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.n));
			Assert.AreEqual(0, node.Items.Count(i => i.Type == PivotCacheRecordType.s));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldNodeTestNullNode()
		{
			new CacheFieldNode(TestUtility.CreateDefaultNSM(), null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldNodeTestNullNamespaceManager()
		{
			var xml = new XmlDocument();
			xml.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			new CacheFieldNode(null, xml.FirstChild);
		}
		#endregion
	}
}