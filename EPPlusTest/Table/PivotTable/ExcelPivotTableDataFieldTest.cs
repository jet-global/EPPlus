/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Evan Schallerer and others as noted in the source history.
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
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotTableDataFieldTest
	{
		#region ShowDataAs TestMethods
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void GetSetShowDataAs()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var dataField = pivotTable.DataFields.First(f => f.Index == 4);
					Assert.AreEqual(ShowDataAs.PercentOfTotal, dataField.ShowDataAs);

					dataField.ShowDataAs = ShowDataAs.PercentOfRow;
					Assert.AreEqual(ShowDataAs.PercentOfRow, dataField.ShowDataAs);

					dataField.ShowDataAs = ShowDataAs.NoCalculation;
					Assert.AreEqual(ShowDataAs.NoCalculation, dataField.ShowDataAs);
					Assert.IsNull(dataField.TopNode.Attributes["showDataAs"]);

					Assert.IsNull(dataField.TopNode.SelectSingleNode("d:extLst/d:ext/x14:dataField", dataField.NameSpaceManager));
					dataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					Assert.AreEqual(ShowDataAs.PercentOfParentRow, dataField.ShowDataAs);
					var extDataFieldNode = dataField.TopNode.SelectSingleNode("d:extLst/d:ext/x14:dataField/@pivotShowAs", dataField.NameSpaceManager);
					Assert.IsNotNull(extDataFieldNode);
				}
			}
		}
		#endregion
	}
}
