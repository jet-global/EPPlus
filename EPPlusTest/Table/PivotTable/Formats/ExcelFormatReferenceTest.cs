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
using OfficeOpenXml.Table.PivotTable.Formats;

namespace EPPlusTest.Table.PivotTable.Formats
{
	[TestClass]
	public class ExcelFormatReferenceTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ExcelFormatReferenceConstructorTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"<reference field=""6"" count =""1"">
								<x v=""0""/>
							</reference>");
			var reference = new ExcelFormatReference(TestUtility.CreateDefaultNSM(), xmlDoc.FirstChild);
			Assert.AreEqual(6, reference.FieldIndex);
			Assert.AreEqual(1, reference.ItemIndexCount);
			Assert.IsTrue(reference.Selected);
			Assert.AreEqual(1, reference.SharedItems.Count);
		}

		[TestMethod]
		public void ExcelFormatReferenceConstructorWithNoSharedItemsTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"<reference field=""6"" count =""0""/>");
			var reference = new ExcelFormatReference(TestUtility.CreateDefaultNSM(), xmlDoc.FirstChild);
			Assert.AreEqual(6, reference.FieldIndex);
			Assert.AreEqual(0, reference.ItemIndexCount);
			Assert.IsTrue(reference.Selected);
			Assert.AreEqual(0, reference.SharedItems.Count);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelFormatReferenceConstructorNullNamespaceManagerTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"<reference field=""6"" count =""1"">
								<x v=""0""/>
							</reference>");
			new ExcelFormatReference(null, xmlDoc.FirstChild);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelFormatReferenceConstructorNullNodeTest()
		{
			new ExcelFormatReference(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion
	}
}
