/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Michelle Lau and others as noted in the source history.
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
using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotTableFieldTest
	{
		#region DisableDefaultSubtotal Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableBackedByExcelTable.xlsx")]
		public void DisableDefaultSubtotalRemovesDefaultItem()
		{
			var file = new FileInfo("PivotTableBackedByExcelTable.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					Assert.IsTrue(pivotTable.Fields.All(f => f.DefaultSubtotal));
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					var field = pivotTable.Fields[2];
					Assert.AreEqual(4, field.Items.Count);
					Assert.AreEqual("default", field.Items[3].T);
					field = pivotTable.Fields[3];
					Assert.AreEqual(5, field.Items.Count);
					Assert.AreEqual("default", field.Items[4].T);
					foreach (var ptField in pivotTable.Fields)
					{
						ptField.SubTotalFunctions = OfficeOpenXml.Table.PivotTable.eSubTotalFunctions.None;
					}
					Assert.IsFalse(pivotTable.Fields.All(f => f.DefaultSubtotal));
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					field = pivotTable.Fields[2];
					Assert.AreEqual(3, field.Items.Count);
					Assert.IsTrue(string.IsNullOrEmpty(field.Items[2].T));
					field = pivotTable.Fields[3];
					Assert.AreEqual(4, field.Items.Count);
					Assert.IsTrue(string.IsNullOrEmpty(field.Items[3].T));
					package.SaveAs(newFile.File);
				}
				using (var package = new ExcelPackage(newFile.File))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					Assert.IsFalse(pivotTable.Fields.All(f => f.DefaultSubtotal));
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					var field = pivotTable.Fields[2];
					Assert.AreEqual(3, field.Items.Count);
					Assert.IsTrue(string.IsNullOrEmpty(field.Items[2].T));
					field = pivotTable.Fields[3];
					Assert.AreEqual(4, field.Items.Count);
					Assert.IsTrue(string.IsNullOrEmpty(field.Items[3].T));
					package.SaveAs(newFile.File);
				}
			}
		}
		#endregion
	}
}
