/*******************************************************************************
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
 * Jan Källman		    Initial Release		        2010-01-28
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 * Eyal Seagull		    Conditional Formatting      2012-04-03
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing.Sparkline;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;

namespace OfficeOpenXml
{
	/// <summary>
	/// A range of cells 
	/// </summary>
	public class ExcelRangeBase : ExcelAddress, IExcelCell, IEnumerable<ExcelRangeBase>, IEnumerator<ExcelRangeBase>
	{
		#region Class Variables
		/// <summary>
		/// Reference to the worksheet.
		/// </summary>
		protected ExcelWorksheet myWorksheet;
		protected ExcelRichTextCollection myExcelRichTextCollection = null;
		internal ExcelWorkbook myWorkbook = null;
		private delegate void _changeProp(ExcelRangeBase range, _setValue method, object value);
		private delegate void _setValue(ExcelRangeBase range, object value, int row, int col);
		private _changeProp myChangePropMethod;
		private int myStyleId;
		#endregion

		#region Nested Classes
		private class CopiedCell
		{
			internal int Row { get; set; }
			internal int Column { get; set; }
			internal object Value { get; set; }
			internal string Type { get; set; }
			internal object Formula { get; set; }
			internal int? StyleID { get; set; }
			internal Uri HyperLink { get; set; }
			internal ExcelComment Comment { get; set; }
			internal byte Flag { get; set; }
		}
		#endregion

		#region Constructors
		internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
		{
			this.myWorksheet = xlWorksheet;
			this._ws = this.myWorksheet.Name;
			this.myWorkbook = this.myWorksheet.Workbook;
			SetDelegate();
		}

		/// <summary>
		/// On change address handler
		/// </summary>
		protected internal override void ChangeAddress()
		{
			if (this.IsTableAddress)
				SetRCFromTable(this.myWorkbook.Package, null);
			if (this.myWorksheet.Name != this.WorkSheet)
				this.myWorksheet = this.myWorkbook.Worksheets[this.WorkSheet];
			SetDelegate();
		}

		internal ExcelRangeBase(ExcelWorksheet xlWorksheet, string address) : base(address)
		{
			if (!string.IsNullOrEmpty(xlWorksheet?.Name))
				_ws = xlWorksheet.Name;
			this.myWorksheet = xlWorksheet;
			this.myWorkbook = this.myWorksheet.Workbook;
			base.SetRCFromTable(this.myWorksheet.Package, null);
			if (string.IsNullOrEmpty(_ws))
				_ws = myWorksheet == null ? string.Empty : myWorksheet.Name;
			SetDelegate();
		}

		internal ExcelRangeBase(ExcelWorkbook wb, ExcelWorksheet xlWorksheet, string address, bool isName) : base(address, isName)
		{
			if (!string.IsNullOrEmpty(xlWorksheet?.Name))
				_ws = xlWorksheet.Name;
			SetRCFromTable(wb.Package, null);
			this.myWorksheet = xlWorksheet;
			this.myWorkbook = wb;
			if (string.IsNullOrEmpty(_ws))
				_ws = xlWorksheet?.Name;
			SetDelegate();
		}
		#endregion

		#region Set Value Delegates
		private static _changeProp _setUnknownProp = SetUnknown;
		private static _changeProp _setSingleProp = SetSingle;
		private static _changeProp _setRangeProp = SetRange;
		private static _changeProp _setMultiProp = SetMultiRange;

		private void SetDelegate()
		{
			if (this._fromRow == -1)
			{
				this.myChangePropMethod = SetUnknown;
			}
			//Single cell
			else if (this._fromRow == this._toRow && this._fromCol == this._toCol && this.Addresses == null)
			{
				this.myChangePropMethod = SetSingle;
			}
			//Range (ex A1:A2)
			else if (this.Addresses == null)
			{
				this.myChangePropMethod = SetRange;
			}
			//Multi Range (ex A1:A2,C1:C2)
			else
			{
				this.myChangePropMethod = SetMultiRange;
			}
		}

		/// <summary>
		/// We dont know the address yet. Set the delegate first time a property is set.
		/// </summary>
		/// <param name="range"></param>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private static void SetUnknown(ExcelRangeBase range, _setValue valueMethod, object value)
		{
			//Address is not set use, selected range
			if (range._fromRow == -1)
			{
				range.SetToSelectedRange();
			}
			range.SetDelegate();
			range.myChangePropMethod(range, valueMethod, value);
		}

		/// <summary>
		/// Set a single cell
		/// </summary>
		/// <param name="range"></param>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private static void SetSingle(ExcelRangeBase range, _setValue valueMethod, object value)
		{
			valueMethod(range, value, range._fromRow, range._fromCol);
		}

		/// <summary>
		/// Set a range
		/// </summary>
		/// <param name="range"></param>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private static void SetRange(ExcelRangeBase range, _setValue valueMethod, object value)
		{
			range.SetValueAddress(range, valueMethod, value);
		}

		/// <summary>
		/// Set a multirange (A1:A2,C1:C2)
		/// </summary>
		/// <param name="range"></param>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private static void SetMultiRange(ExcelRangeBase range, _setValue valueMethod, object value)
		{
			range.SetValueAddress(range, valueMethod, value);
			foreach (var address in range.Addresses)
			{
				range.SetValueAddress(address, valueMethod, value);
			}
		}

		/// <summary>
		/// Set the property for an address
		/// </summary>
		/// <param name="address"></param>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private void SetValueAddress(ExcelAddress address, _setValue valueMethod, object value)
		{
			IsRangeValid(string.Empty);
			if (this._fromRow == 1 && this._fromCol == 1 && this._toRow == ExcelPackage.MaxRows && this._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
				throw (new ArgumentException("Can't reference all cells. Please use the indexer to set the range"));
			else
			{
				if (value is object[,] && valueMethod == Set_Value)
				{
					// only simple set value is supported for bulk copy
					this.myWorksheet.SetRangeValueInner(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, (object[,])value);
				}
				else
				{
					for (int col = address.Start.Column; col <= address.End.Column; col++)
					{
						for (int row = address.Start.Row; row <= address.End.Row; row++)
						{
							valueMethod(this, value, row, col);
						}
					}
				}
			}
		}
		#endregion

		#region Properties
		/// <summary>
		/// Converts an <see cref="ExcelRangeBase"/> to a string.
		/// </summary>
		/// <param name="excelRangeBase">The <see cref="ExcelRangeBase"/> to convert to a string.</param>
		public static implicit operator string(ExcelRangeBase excelRangeBase)
		{
			return excelRangeBase.ToString();
		}

		/// <summary>
		/// The styleobject for the range.
		/// </summary>
		public ExcelStyle Style
		{
			get
			{
				IsRangeValid("styling");
				int s = 0;
				if (!this.myWorksheet.ExistsStyleInner(this._fromRow, this._fromCol, ref s)) // Cell exists.
				{
					if (!this.myWorksheet.ExistsStyleInner(this._fromRow, 0, ref s)) // Check Row style.
					{
						var column = this.Worksheet.GetColumn(this._fromCol);
						s = column == null ? 0 : column.StyleID;
					}
				}
				return this.myWorksheet.Workbook.Styles.GetStyleObject(s, this.myWorksheet.PositionID, this.Address);
			}
		}

		/// <summary>
		/// The named style.
		/// </summary>
		public string StyleName
		{
			get
			{
				IsRangeValid("styling");
				int xfId;
				if (this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows)
					xfId = GetColumnStyle(this._fromCol);
				else if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns)
				{
					xfId = 0;
					if (!this.myWorksheet.ExistsStyleInner(this._fromRow, 0, ref xfId))
						xfId = GetColumnStyle(this._fromCol);
				}
				else
				{
					xfId = 0;
					if (!this.myWorksheet.ExistsStyleInner(this._fromRow, this._fromCol, ref xfId) && !this.myWorksheet.ExistsStyleInner(this._fromRow, 0, ref xfId))
						xfId = GetColumnStyle(this._fromCol);
				}
				int nsID;
				if (xfId <= 0)
					nsID = this.Style.Styles.CellXfs[0].XfId;
				else
					nsID = this.Style.Styles.CellXfs[xfId].XfId;
				foreach (var ns in this.Style.Styles.NamedStyles)
				{
					if (ns.StyleXfId == nsID)
						return ns.Name;
				}
				return string.Empty;
			}
			set
			{
				this.myStyleId = this.myWorksheet.Workbook.Styles.GetStyleIdFromName(value);
				int col = this._fromCol;
				if (this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows)    //Full column
				{
					ExcelColumn column;
					var c = this.myWorksheet.GetValue(0, this._fromCol);
					if (c == null)
						column = this.myWorksheet.Column(this._fromCol);
					else
						column = (ExcelColumn)c;

					column.StyleName = value;
					column.StyleID = this.myStyleId;

					var cols = this.myWorksheet._values.GetEnumerator(0, this._fromCol + 1, 0, this._toCol);
					if (cols.MoveNext())
					{
						col = this._fromCol;
						while (column.ColumnMin <= this._toCol)
						{
							if (column.ColumnMax > this._toCol)
							{
								var newCol = this.myWorksheet.CopyColumn(column, this._toCol + 1, column.ColumnMax);
								column.ColumnMax = this._toCol;
							}

							column._styleName = value;
							column.StyleID = this.myStyleId;

							if (cols.Value._value == null)
								break;
							else
							{
								var nextCol = (ExcelColumn)cols.Value._value;
								if (column.ColumnMax < nextCol.ColumnMax - 1)
									column.ColumnMax = nextCol.ColumnMax - 1;
								column = nextCol;
								cols.MoveNext();
							}
						}
					}
					if (column.ColumnMax < this._toCol)
						column.ColumnMax = this._toCol;

					if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns) //FullRow
					{
						var rows = this.myWorksheet._values.GetEnumerator(1, 0, ExcelPackage.MaxRows, 0);
						rows.MoveNext();
						while (rows.Value._value != null)
						{
							this.myWorksheet.SetStyleInner(rows.Row, 0, this.myStyleId);
							if (!rows.MoveNext())
								break;
						}
					}
				}
				else if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns) //FullRow
				{
					for (int r = this._fromRow; r <= this._toRow; r++)
					{
						this.myWorksheet.Row(r)._styleName = value;
						this.myWorksheet.Row(r).StyleID = this.myStyleId;
					}
				}

				if (!((this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows) || (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns))) //Cell specific
				{
					for (int column = this._fromCol; column <= this._toCol; column++)
					{
						for (int row = this._fromRow; row <= this._toRow; row++)
						{
							this.myWorksheet.SetStyleInner(row, column, this.myStyleId);
						}
					}
				}
				else // Only set name on created cells. (uncreated cells is set on full row or full column).
				{
					var cells = this.myWorksheet._values.GetEnumerator(this._fromRow, this._fromCol, this._toRow, this._toCol);
					while (cells.MoveNext())
					{
						this.myWorksheet.SetStyleInner(cells.Row, cells.Column, this.myStyleId);
					}
				}
			}
		}

		private int GetColumnStyle(int col)
		{
			object c = null;
			if (this.myWorksheet.ExistsValueInner(0, col, ref c))
				return (c as ExcelColumn).StyleID;
			else
			{
				int row = 0;
				if (this.myWorksheet._values.PrevCell(ref row, ref col))
				{
					var column = this.myWorksheet.GetValueInner(row, col) as ExcelColumn;
					if (column.ColumnMax >= col)
						return this.myWorksheet.GetStyleInner(row, col);
				}
			}
			return 0;
		}

		/// <summary>
		/// The style ID. 
		/// It is not recomended to use this one. Use Named styles as an alternative.
		/// If you do, make sure that you use the Style.UpdateXml() method to update any new styles added to the workbook.
		/// </summary>
		public int StyleID
		{
			get
			{
				int s = 0;
				if (!this.myWorksheet.ExistsStyleInner(this._fromRow, this._fromCol, ref s))
				{
					if (!this.myWorksheet.ExistsStyleInner(this._fromRow, 0, ref s))
						s = this.myWorksheet.GetStyleInner(0, this._fromCol);
				}
				return s;
			}
			set
			{
				this.myChangePropMethod(this, _setStyleIdDelegate, value);
			}
		}

		/// <summary>
		/// Set the range to a specific value.
		/// </summary>
		public object Value
		{
			get
			{
				if (this._fromRow == this._toRow && this._fromCol == this._toCol)
					return this.myWorksheet.GetValue(this._fromRow, this._fromCol);
				else
					return GetValueArray();
			}
			set
			{
				this.myChangePropMethod(this, _setValueDelegate, value);
			}
		}

		/// <summary>
		/// Returns the formatted value.
		/// </summary>
		public string Text
		{
			get
			{
				return GetFormattedText(false);
			}
		}

		/// <summary>
		/// Gets or sets a formula for a range.
		/// </summary>
		public string Formula
		{
			get
			{
				return this.myWorksheet.GetFormula(this._fromRow, this._fromCol);
			}
			set
			{
				this.SetFormula(value, true);
			}
		}

		/// <summary>
		/// Gets or Set a formula in R1C1 format.
		/// </summary>
		public string FormulaR1C1
		{
			get
			{
				IsRangeValid("FormulaR1C1");
				return this.myWorksheet.GetFormulaR1C1(this._fromRow, this._fromCol);
			}
			set
			{
				IsRangeValid("FormulaR1C1");
				if (value.Length > 0 && value[0] == '=')
					value = value.Substring(1, value.Length - 1); // remove any starting equalsign.

				if (value == null || value.Trim() == string.Empty)
				{
					//Set the cells to null
					this.myWorksheet.Cells[ExcelCellBase.TranslateFromR1C1(value, this._fromRow, this._fromCol)].Value = null;
				}
				else if (this.Addresses == null)
					SetSharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, this._fromRow, this._fromCol), this, false);
				else
				{
					SetSharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, this._fromRow, this._fromCol), new ExcelAddress(this.WorkSheet, this.FirstAddress), false);
					foreach (var address in this.Addresses)
					{
						SetSharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, address.Start.Row, address.Start.Column), address, false);
					}
				}
			}
		}

		/// <summary>
		/// Set the hyperlink property for a range of cells
		/// </summary>
		public Uri Hyperlink
		{
			get
			{
				IsRangeValid("formulaR1C1");
				return this.myWorksheet._hyperLinks.GetValue(this._fromRow, this._fromCol);
			}
			set
			{
				this.myChangePropMethod(this, _setHyperLinkDelegate, value);
			}
		}

		/// <summary>
		/// If the cells in the range are merged.
		/// </summary>
		public bool Merge
		{
			get
			{
				IsRangeValid("merging");
				for (int col = this._fromCol; col <= this._toCol; col++)
				{
					for (int row = this._fromRow; row <= this._toRow; row++)
					{
						if (this.myWorksheet.MergedCells[row, col] == null)
							return false;
					}
				}
				return true;
			}
			set
			{
				IsRangeValid("merging");
				this.myWorksheet.MergedCells.Clear(this);
				if (value)
				{
					this.myWorksheet.MergedCells.Add(new ExcelAddress(this.FirstAddress), true);
					if (this.Addresses != null)
					{
						foreach (var address in this.Addresses)
						{
							this.myWorksheet.MergedCells.Clear(address); //Fixes issue 15482
							this.myWorksheet.MergedCells.Add(address, true);
						}
					}
				}
				else
				{
					if (this.Addresses != null)
					{
						foreach (var address in this.Addresses)
						{
							this.myWorksheet.MergedCells.Clear(address);
						}
					}

				}
			}
		}

		/// <summary>
		/// Set an autofilter for the range
		/// </summary>
		public bool AutoFilter
		{
			get
			{
				IsRangeValid("autofilter");
				ExcelAddress address = this.myWorksheet.AutoFilterAddress;
				if (address == null)
					return false;
				if (this._fromRow >= address.Start.Row
						&& this._toRow <= address.End.Row
						&& this._fromCol >= address.Start.Column
						&& this._toCol <= address.End.Column)
				{
					return true;
				}
				return false;
			}
			set
			{
				IsRangeValid("autofilter");
				if (this.myWorksheet.AutoFilterAddress != null)
				{
					var collision = this.Collide(this.myWorksheet.AutoFilterAddress);
					if (value == false && (collision == eAddressCollition.Partly || collision == eAddressCollition.No))
						throw (new InvalidOperationException("Can't remote Autofilter. Current autofilter does not match selected range."));
				}
				if (this.myWorksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
					this.myWorksheet.Names.Remove("_xlnm._FilterDatabase");
				if (value)
				{
					this.myWorksheet.AutoFilterAddress = this;
					var result = this.myWorksheet.Names.Add("_xlnm._FilterDatabase", this);
					result.IsNameHidden = true;
				}
				else
					this.myWorksheet.AutoFilterAddress = null;
			}
		}

		/// <summary>
		/// If the value is in richtext format.
		/// </summary>
		public bool IsRichText
		{
			get
			{
				IsRangeValid("richtext");
				return this.myWorksheet._flags.GetFlagValue(this._fromRow, this._fromCol, CellFlags.RichText);
			}
			set
			{
				this.myChangePropMethod(this, _setIsRichTextDelegate, value);
			}
		}

		/// <summary>
		/// Is the range a part of an Arrayformula
		/// </summary>
		public bool IsArrayFormula
		{
			get
			{
				IsRangeValid("arrayformulas");
				return this.myWorksheet._flags.GetFlagValue(this._fromRow, this._fromCol, CellFlags.ArrayFormula);
			}
		}

		/// <summary>
		/// Cell value is richtext formatted. 
		/// Richtext-property only apply to the left-top cell of the range.
		/// </summary>
		public ExcelRichTextCollection RichText
		{
			get
			{
				IsRangeValid("richtext");
				if (this.myExcelRichTextCollection == null)
					this.myExcelRichTextCollection = GetRichText(this._fromRow, this._fromCol);
				return this.myExcelRichTextCollection;
			}
		}

		private ExcelRichTextCollection GetRichText(int row, int col)
		{
			XmlDocument xml = new XmlDocument();
			var innerValue = this.myWorksheet.GetValueInner(row, col);
			var isRichText = this.myWorksheet._flags.GetFlagValue(row, col, CellFlags.RichText);
			if (innerValue != null)
			{
				if (isRichText)
					XmlHelper.LoadXmlSafe(xml, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + innerValue.ToString() + "</d:si>", Encoding.UTF8);
				else
					xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ><d:r><d:t>" + OfficeOpenXml.Utils.ConvertUtil.ExcelEscapeString(innerValue.ToString()) + "</d:t></d:r></d:si>");
			}
			else
				xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
			var richTextCollection = new ExcelRichTextCollection(this.myWorksheet.NameSpaceManager, xml.SelectSingleNode("d:si", this.myWorksheet.NameSpaceManager), this);
			return richTextCollection;
		}

		/// <summary>
		/// returns the comment object of the first cell in the range
		/// </summary>
		public ExcelComment Comment
		{
			get
			{
				IsRangeValid("comments");
				var i = -1;
				if (this.myWorksheet.Comments.Count > 0 && this.myWorksheet._commentsStore.Exists(this._fromRow, this._fromCol, out i))
						return this.myWorksheet.Comments[i] as ExcelComment;

				return null;
			}
		}

		/// <summary>
		/// WorkSheet object 
		/// </summary>
		public ExcelWorksheet Worksheet
		{
			get
			{
				return this.myWorksheet;
			}
		}

		/// <summary>
		/// Address including sheetname
		/// </summary>
		public string FullAddressAbsolute
		{
			get
			{
				string wbwsRef = string.IsNullOrEmpty(this._wb) ? this._ws : $"[{this._wb.Replace("'", "''")}]{this._ws}";
				if (this.Addresses == null)
					return ExcelCellBase.GetFullAddress(wbwsRef, GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol, true));
				string fullAddress = string.Empty;
				foreach (var address in this.Addresses)
				{
					if (address.Address == "#REF!")
						fullAddress += ExcelCellBase.GetFullAddress(wbwsRef, "#REF!") + ',';
					else
						fullAddress += ExcelCellBase.GetFullAddress(wbwsRef, GetAddress(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, true)) + ',';
				}
				return fullAddress.TrimEnd(',');
			}
		}
		#endregion

		#region Private Methods
		#region Set Property Methods
		private static _setValue _setStyleIdDelegate = Set_StyleID;
		private static _setValue _setValueDelegate = Set_Value;
		private static _setValue _setHyperLinkDelegate = Set_HyperLink;
		private static _setValue _setIsRichTextDelegate = Set_IsRichText;

		private static void Set_StyleID(ExcelRangeBase range, object value, int row, int col)
		{
			range.myWorksheet.SetStyleInner(row, col, (int)value);
		}

		private static void Set_StyleName(ExcelRangeBase range, object value, int row, int col)
		{
			range.myWorksheet.SetStyleInner(row, col, range.myStyleId);
		}

		private static void Set_Value(ExcelRangeBase range, object value, int row, int col)
		{
			var sfi = range.myWorksheet._formulas.GetValue(row, col);
			if (sfi is int)
				range.UpdateSharedFormulaAnchors(range.myWorksheet.Cells[row, col]);
			if (sfi != null)
				range.myWorksheet._formulas.SetValue(row, col, string.Empty);
			range.myWorksheet.SetValueInner(row, col, value);
		}

		private static void SetFormula(ExcelRangeBase range, string value, int row, int col, bool clearValue = true)
		{
			var formulaValue = range.myWorksheet._formulas.GetValue(row, col);
			var valueValue = range.myWorksheet._values.GetValue(row, col);
			if (formulaValue is int && (int)formulaValue >= 0)
				range.UpdateSharedFormulaAnchors(range.myWorksheet.Cells[row, col]);
			string formula = value ?? string.Empty;
			range.myWorksheet._formulas.SetValue(row, col, formula);
			if ((formula != string.Empty && clearValue) || valueValue.Equals(default(ExcelCoreValue)))
				range.myWorksheet.SetValueInner(row, col, null);
		}

		/// <summary>
		/// Handles shared formulas
		/// </summary>
		/// <param name="range">The range</param>
		/// <param name="value">The  formula</param>
		/// <param name="address">The address of the formula</param>
		/// <param name="IsArray">If the forumla is an array formula.</param>
		/// <param name="clearValue">Whether or not setting the formula is allowed to clear the cell's value.</param>
		private static void SetSharedFormula(ExcelRangeBase range, string value, ExcelAddress address, bool IsArray, bool clearValue = true)
		{
			if (range._fromRow == 1 && range._fromCol == 1 && range._toRow == ExcelPackage.MaxRows && range._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
			{
				throw (new InvalidOperationException("Can't set a formula for the entire worksheet"));
			}
			else if (address.Start.Row == address.End.Row && address.Start.Column == address.End.Column && !IsArray)             //is it really a shared formula? Arrayformulas can be one cell only
			{
				// Single cells get individual formulas.
				SetFormula(range, value, address.Start.Row, address.Start.Column, clearValue);
				return;
			}
			range.CheckAndSplitSharedFormula(address);
			ExcelWorksheet.Formulas formulas = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default)
			{
				Formula = value,
				Index = range.myWorksheet.GetMaxShareFunctionIndex(IsArray),
				Address = address.FirstAddress,
				StartCol = address.Start.Column,
				StartRow = address.Start.Row,
				IsArray = IsArray
			};
			range.myWorksheet._sharedFormulas.Add(formulas.Index, formulas);

			for (int col = address.Start.Column; col <= address.End.Column; col++)
			{
				for (int row = address.Start.Row; row <= address.End.Row; row++)
				{
					range.myWorksheet._formulas.SetValue(row, col, formulas.Index);
					range.myWorksheet.SetValueInner(row, col, null);
				}
			}
		}

		private static void Set_HyperLink(ExcelRangeBase range, object value, int row, int col)
		{
			if (value is Uri)
			{
				range.myWorksheet._hyperLinks.SetValue(row, col, (Uri)value);

				if (value is ExcelHyperLink)
					range.myWorksheet.SetValueInner(row, col, ((ExcelHyperLink)value).Display);
				else
				{
					var v = range.myWorksheet.GetValueInner(row, col);
					if (v == null || v.ToString() == string.Empty)
						range.myWorksheet.SetValueInner(row, col, ((Uri)value).OriginalString);
				}
			}
			else
			{
				range.myWorksheet._hyperLinks.SetValue(row, col, (Uri)null);
				range.myWorksheet.SetValueInner(row, col, (Uri)null);
			}
		}

		private static void Set_IsRichText(ExcelRangeBase range, object value, int row, int col)
		{
			range.myWorksheet._flags.SetFlagValue(row, col, (bool)value, CellFlags.RichText);
		}
		#endregion
		
		/// <summary>
		/// Set the value without altering the richtext property
		/// </summary>
		/// <param name="value">the value</param>
		internal void SetValueRichText(object value)
		{
			if (this._fromRow == 1 && this._fromCol == 1 && this._toRow == ExcelPackage.MaxRows && this._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
				SetValue(value, 1, 1);
			else
				SetValue(value, this._fromRow, this._fromCol);
		}

		private void SetValue(object value, int row, int col)
		{
			this.myWorksheet.SetValue(row, col, value);
			this.myWorksheet._formulas.SetValue(row, col, "");
		}

		internal void SetSharedFormulaID(int id)
		{
			for (int col = this._fromCol; col <= this._toCol; col++)
			{
				for (int row = this._fromRow; row <= this._toRow; row++)
				{
					if (this.myWorksheet._formulas.GetValue(row, col) is int)
						this.myWorksheet._formulas.SetValue(row, col, id);
				}
			}
		}

		private void CheckAndSplitSharedFormula(ExcelAddress address)
		{
			for (int col = address._fromCol; col <= address._toCol; col++)
			{
				for (int row = address._fromRow; row <= address._toRow; row++)
				{
					var formula = this.myWorksheet._formulas.GetValue(row, col);
					if (formula is int && (int)formula >= 0)
					{
						UpdateSharedFormulaAnchors(address);
						return;
					}
				}
			}
		}

		private void UpdateSharedFormulaAnchors(ExcelAddress address)
		{
			List<int> formulas = new List<int>();
			for (int col = address._fromCol; col <= address._toCol; col++)
			{
				for (int row = address._fromRow; row <= address._toRow; row++)
				{
					var formula = this.myWorksheet._formulas.GetValue(row, col);
					if (formula is int id && id >= 0 && !formulas.Contains(id))
					{
						if (this.myWorksheet._sharedFormulas[id].IsArray &&
								Collide(this.myWorksheet.Cells[this.myWorksheet._sharedFormulas[id].Address]) == eAddressCollition.Partly) // If the formula is an array formula and its on the inside the overwriting range throw an exception
						{
							throw (new InvalidOperationException("Can not overwrite a part of an array-formula"));
						}
						formulas.Add(id);
					}
				}
			}

			foreach (int ix in formulas)
			{
				UpdateSharedFormulaAnchor(address, ix);
			}
		}

		private void UpdateSharedFormulaAnchor(ExcelAddress address, int ix)
		{
			var sharedFormula = this.myWorksheet._sharedFormulas[ix];
			var fRange = this.myWorksheet.Cells[sharedFormula.Address];
			var collide = address.Collide(fRange);
			switch (collide)
			{
				case eAddressCollition.Equal:
				case eAddressCollition.Inside:
					this.myWorksheet._sharedFormulas.Remove(ix);
					break;
				case eAddressCollition.Partly:
					if (this.TryFindNewAnchorCoordinates(sharedFormula, address, out int row, out int column))
					{
						var r1c1 = ExcelCellBase.TranslateToR1C1(sharedFormula.Formula, sharedFormula.StartRow, sharedFormula.StartCol);
						sharedFormula.StartRow = row;
						sharedFormula.StartCol = column;
						sharedFormula.Formula = ExcelCellBase.TranslateFromR1C1(r1c1, sharedFormula.StartRow, sharedFormula.StartCol);
					}
					else
						this.myWorksheet._sharedFormulas.Remove(ix);
					break;
				default:
				case eAddressCollition.No:
					break;
			}
		}

		private bool TryFindNewAnchorCoordinates(ExcelWorksheet.Formulas currentFormula, ExcelAddress newAddress, out int row, out int column)
		{
			row = -1;
			column = -1;
			var currentAddress = new ExcelAddress(currentFormula.Address);
			for (int r = currentAddress._fromRow; r <= currentAddress._toRow; r++)
			{
				for (int c = currentAddress._fromCol; c <= currentAddress._toCol; c++)
				{
					var f = this.myWorksheet._formulas.GetValue(r, c) as int?;
					if (f == currentFormula.Index)
					{
						if (!newAddress.ContainsCoordinate(r, c))
						{
							row = r;
							column = c;
							return true;
						}
					}
				}
			}
			return false;
		}

		private object ConvertData(ExcelTextFormat Format, string value, int column, bool isText)
		{
			if (isText && (Format.DataTypes == null || Format.DataTypes.Length < column))
				return string.IsNullOrEmpty(value) ? null : value;

			double doubleValue;
			DateTime dateTime;
			if (Format.DataTypes == null || Format.DataTypes.Length <= column || Format.DataTypes[column] == eDataTypes.Unknown)
			{
				string v2 = value.EndsWith("%") ? value.Substring(0, value.Length - 1) : value;
				if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out doubleValue))
				{
					if (v2 == value)
						return doubleValue;
					else
						return doubleValue / 100;
				}
				if (DateTime.TryParse(value, Format.Culture, DateTimeStyles.None, out dateTime))
					return dateTime;
				else
					return value;
			}
			else
			{
				switch (Format.DataTypes[column])
				{
					case eDataTypes.Number:
						if (double.TryParse(value, NumberStyles.Any, Format.Culture, out doubleValue))
							return doubleValue;
						else
							return value;
					case eDataTypes.DateTime:
						if (DateTime.TryParse(value, Format.Culture, DateTimeStyles.None, out dateTime))
							return dateTime;
						else
							return value;
					case eDataTypes.Percent:
						string v2 = value.EndsWith("%") ? value.Substring(0, value.Length - 1) : value;
						if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out doubleValue))
							return doubleValue / 100;
						else
							return value;
					case eDataTypes.String:
						return value;
					default:
						return string.IsNullOrEmpty(value) ? null : value;

				}
			}
		}

		private bool IsCellInConditionalFormatAddress(ExcelRangeBase cell, ConditionalFormatting.Contracts.IExcelConditionalFormattingRule rule)
		{
			var addresses = rule.Address.AddressSpaceSeparated.Split(' ');
			foreach (var address in addresses)
			{
				var addressBase = new ExcelAddress(address);
				var collision = addressBase.Collide(cell, true);
				if (collision != eAddressCollition.No)
					return true;
			}
			return false;
		}

		private void SetMinWidth(double minimumWidth, int fromCol, int toCol)
		{
			var iterator = this.myWorksheet._values.GetEnumerator(0, fromCol, 0, toCol);
			var prevCol = fromCol;
			foreach (ExcelCoreValue val in iterator)
			{
				var col = (ExcelColumn)val._value;
				col.Width = minimumWidth;
				if (this.myWorksheet.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
				{
					var newCol = this.myWorksheet.Column(prevCol);
					newCol.ColumnMax = col.ColumnMin - 1;
					newCol.Width = minimumWidth;
				}
				prevCol = col.ColumnMax + 1;
			}
			if (this.myWorksheet.DefaultColWidth > minimumWidth && prevCol < toCol)
			{
				var newCol = this.myWorksheet.Column(prevCol);
				newCol.ColumnMax = toCol;
				newCol.Width = minimumWidth;
			}
		}

		private string GetTextForWidth(Font font)
		{
			string formattedText = this.GetFormattedText(true);
			if (this.Style.WrapText)
				return this.GetLongestStringFromWrappedText(formattedText, font);
			return formattedText;
		}

		private string GetLongestStringFromWrappedText(string text, Font font)
		{
			var rowHeight = this.myWorksheet.Row(this._fromRow).Height;
			var textHeight = ExcelFontXml.GetFontHeight(font.Name, font.Size) * 0.75;
			var leftoverRowHeight = rowHeight % textHeight;
			int textLines = (int)Math.Round((rowHeight - leftoverRowHeight) / textHeight);
			if (textLines <= 1)
				return text;
			var splitIndex = (int)Math.Round((double)(text.Length / textLines));
			var longestTextBuilder = new StringBuilder();
			foreach (var segment in text.Split(' '))
			{
				longestTextBuilder.Append(segment);
				if (longestTextBuilder.Length >= splitIndex)
					return longestTextBuilder.ToString();
				longestTextBuilder.Append(" ");
			}
			// Theoretically shouldn't hit this case.
			return text;
		}

		private string GetFormattedText(bool forWidthCalc)
		{
			object v = this.Value;
			if (v == null)
				return string.Empty;
			var styles = this.Worksheet.Workbook.Styles;
			var nfID = styles.CellXfs[this.StyleID].NumberFormatId;
			ExcelNumberFormatXml.ExcelFormatTranslator nf = null;
			for (int i = 0; i < styles.NumberFormats.Count; i++)
			{
				if (nfID == styles.NumberFormats[i].NumFmtId)
				{
					nf = styles.NumberFormats[i].FormatTranslator;
					break;
				}
			}

			string format, textFormat;
			if (forWidthCalc)
			{
				format = nf.NetFormatForWidth;
				textFormat = nf.NetTextFormatForWidth;
			}
			else
			{
				format = nf.NetFormat;
				textFormat = nf.NetTextFormat;
			}
			return FormatValue(v, nf, format, textFormat);
		}

		internal static string FormatValue(object v, ExcelNumberFormatXml.ExcelFormatTranslator nf, string format, string textFormat)
		{
			if (v is decimal || v.GetType().IsPrimitive)
			{
				double d;
				try
				{
					d = Convert.ToDouble(v);
				}
				catch
				{
					return string.Empty;
				}

				if (nf.DataType == ExcelNumberFormatXml.eFormatType.Number)
				{
					if (string.IsNullOrEmpty(nf.FractionFormat))
						return d.ToString(format, nf.Culture);
					else
						return nf.FormatFraction(d);
				}
				else if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
				{
					var date = DateTime.FromOADate(d);
					return date.ToString(format, nf.Culture);
				}
			}
			else if (v is DateTime)
			{
				if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
					return ((DateTime)v).ToString(format, nf.Culture);
				else
				{
					double d = ((DateTime)v).ToOADate();
					if (string.IsNullOrEmpty(nf.FractionFormat))
						return d.ToString(format, nf.Culture);
					else
						return nf.FormatFraction(d);
				}
			}
			else if (v is TimeSpan)
			{
				if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
					return new DateTime(((TimeSpan)v).Ticks).ToString(format, nf.Culture);
				else
				{
					double d = DateTime.FromOADate(0).Add((TimeSpan)v).ToOADate();
					if (string.IsNullOrEmpty(nf.FractionFormat))
						return d.ToString(format, nf.Culture);
					else
						return nf.FormatFraction(d);
				}
			}
			else
			{
				if (textFormat == string.Empty)
					return v.ToString();
				else
					return string.Format(textFormat, v);
			}
			return v.ToString();
		}

		private void SetToSelectedRange()
		{
			if (this.myWorksheet.View.SelectedRange == string.Empty)
				this.Address = "A1";
			else
				this.Address = this.myWorksheet.View.SelectedRange;
		}

		private void IsRangeValid(string type)
		{
			if (this._fromRow <= 0)
			{
				if (this._address == string.Empty)
					SetToSelectedRange();
				else
				{
					if (type == string.Empty)
						throw (new InvalidOperationException(string.Format("Range is not valid for this operation: {0}", this._address)));
					else
						throw (new InvalidOperationException(string.Format("Range is not valid for {0} : {1}", type, this._address)));
				}
			}
		}

		internal void UpdateAddress(string address)
		{
			throw new NotImplementedException();
		}

		private bool IsInfinityValue(object value)
		{
			double? valueAsDouble = value as double?;
			if (valueAsDouble.HasValue && (double.IsNegativeInfinity(valueAsDouble.Value) || double.IsPositiveInfinity(valueAsDouble.Value)))
				return true;
			return false;
		}

		private object GetValueArray()
		{
			ExcelAddress addr;
			if (this._fromRow == 1 && this._fromCol == 1 && this._toRow == ExcelPackage.MaxRows && this._toCol == ExcelPackage.MaxColumns)
			{
				addr = this.myWorksheet.Dimension;
				if (addr == null)
					return null;
			}
			else
				addr = this;
			object[,] v = new object[addr._toRow - addr._fromRow + 1, addr._toCol - addr._fromCol + 1];

			for (int col = addr._fromCol; col <= addr._toCol; col++)
			{
				for (int row = addr._fromRow; row <= addr._toRow; row++)
				{
					object o = null;
					if (this.myWorksheet.ExistsValueInner(row, col, ref o))
					{
						if (this.myWorksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
							v[row - addr._fromRow, col - addr._fromCol] = GetRichText(row, col).Text;
						else
							v[row - addr._fromRow, col - addr._fromCol] = o;
					}
				}
			}
			return v;
		}

		private ExcelAddress GetAddressDim(ExcelRangeBase addr)
		{
			int fromRow, fromCol, toRow, toCol;
			var d = this.myWorksheet.Dimension;
			fromRow = addr._fromRow < d._fromRow ? d._fromRow : addr._fromRow;
			fromCol = addr._fromCol < d._fromCol ? d._fromCol : addr._fromCol;

			toRow = addr._toRow > d._toRow ? d._toRow : addr._toRow;
			toCol = addr._toCol > d._toCol ? d._toCol : addr._toCol;

			if (addr._fromRow == fromRow && addr._fromCol == fromCol && addr._toRow == toRow && addr._toCol == this._toCol)
				return addr;
			else
			{
				if (this._fromRow > this._toRow || this._fromCol > this._toCol)
					return null;
				else
					return new ExcelAddress(fromRow, fromCol, toRow, toCol);
			}
		}

		private object GetSingleValue()
		{
			if (this.IsRichText)
				return this.RichText.Text;
			else
				return this.myWorksheet.GetValueInner(this._fromRow, this._fromCol);
		}

		private void DeleteCheckMergedCells(ExcelAddress range)
		{
			var removeItems = new List<string>();
			foreach (var addr in this.Worksheet.MergedCells)
			{
				var addrCol = range.Collide(new ExcelAddress(range.WorkSheet, addr));
				if (addrCol != eAddressCollition.No)
				{
					if (addrCol == eAddressCollition.Inside)
						removeItems.Add(addr);
					else
						throw (new InvalidOperationException("Can't remove/overwrite a part of cells that are merged"));
				}
			}
			foreach (var item in removeItems)
			{
				this.Worksheet.MergedCells.Remove(item);
			}
		}

		private void CopySparklines(ExcelRangeBase destination)
		{
			List<ExcelSparklineGroup> newSparklineGroups = new List<ExcelSparklineGroup>();
			foreach (var group in this.Worksheet.SparklineGroups.SparklineGroups)
			{
				ExcelSparklineGroup newGroup = null;
				foreach (var sparkline in group.Sparklines)
				{
					if (sparkline.HostCell.Collide(this) != eAddressCollition.No)
					{
						ExcelAddress newFormula = null;
						if (sparkline.Formula != null)
						{
							ExcelRangeBase.SplitAddress(sparkline.Formula.Address, out string workbook, out string worksheet, out string address);
							string newFormulaString = this.myWorkbook.Package.FormulaManager.UpdateFormulaReferences(address, destination._fromRow - this._fromRow, destination._fromCol - this._fromCol, 0, 0, this.WorkSheet, this.WorkSheet, true);
							if (newFormulaString != "#REF!")
							{
								if (this.Worksheet.Name == worksheet)
									worksheet = destination.Worksheet.Name;
								newFormulaString = string.IsNullOrEmpty(worksheet) ? newFormulaString : ExcelRangeBase.GetFullAddress(worksheet, newFormulaString);
								newFormula = new ExcelAddress(newFormulaString);
							}
						}
						var newHostCellAddress = this.myWorkbook.Package.FormulaManager.UpdateFormulaReferences(sparkline.HostCell.Address, destination._fromRow - this._fromRow, destination._fromCol - this._fromCol, 0, 0, this.WorkSheet, this.WorkSheet, true);
						if (newGroup == null)
						{
							newGroup = new ExcelSparklineGroup(destination.Worksheet, group.NameSpaceManager);
							newGroup.CopyNodeStyle(group.TopNode);
						}
						var newSparkline = new ExcelSparkline(new ExcelAddress(newHostCellAddress), newFormula, newGroup, sparkline.NameSpaceManager);
						newGroup.Sparklines.Add(newSparkline);
					}
				}
				if (newGroup != null)
					newSparklineGroups.Add(newGroup);
			}
			destination.Worksheet.SparklineGroups.SparklineGroups.AddRange(newSparklineGroups);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Compares the range of data.
		/// </summary>
		/// <param name="excelRange">The range being compared to.</param>
		/// <returns>A value indicating if this Excel range is equivalent to the specified range.</returns>
		public bool IsEquivalentRange(ExcelRangeBase excelRange)
		{
			return this.Worksheet.Equals(excelRange.Worksheet)
				&& this.Start.Row == excelRange.Start.Row
				&& this.Start.Column == excelRange.Start.Column
				&& this.End.Row == excelRange.End.Row
				&& this.End.Column == excelRange.End.Column;
		}

		/// <summary>
		/// Conditional Formatting for this range.
		/// </summary>
		public IRangeConditionalFormatting ConditionalFormatting
		{
			get
			{
				return new RangeConditionalFormatting(this.myWorksheet, new ExcelAddress(this.Address));
			}
		}

		/// <summary>
		/// Data validation for this range.
		/// </summary>
		public IRangeDataValidation DataValidation
		{
			get
			{
				return new RangeDataValidation(this.myWorksheet, this.Address);
			}
		}

		/// <summary>
		/// Load the data from the datareader starting from the top left cell of the range
		/// </summary>
		/// <param name="reader">The datareader to loadfrom</param>
		/// <param name="printHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
		/// <param name="tableName">The name of the table</param>
		/// <param name="tableStyle">The table style to apply to the data</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromDataReader(IDataReader reader, bool printHeaders, string tableName, TableStyles tableStyle = TableStyles.None)
		{
			var range = LoadFromDataReader(reader, printHeaders);
			int rows = range.Rows - 1;
			if (rows >= 0 && range.Columns > 0)
			{
				var table = this.myWorksheet.Tables.Add(new ExcelAddress(this._fromRow, this._fromCol, this._fromRow + (rows <= 0 ? 1 : rows), this._fromCol + range.Columns - 1), tableName);
				table.ShowHeader = printHeaders;
				table.TableStyle = tableStyle;
			}
			return range;
		}

		/// <summary>
		/// Load the data from the datareader starting from the top left cell of the range
		/// </summary>
		/// <param name="reader">The datareader to load from</param>
		/// <param name="printHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromDataReader(IDataReader reader, bool printHeaders)
		{
			if (reader == null)
				throw (new ArgumentNullException("Reader", "Reader can't be null"));
			int fieldCount = reader.FieldCount;
			int col = this._fromCol, row = this._fromRow;
			if (printHeaders)
			{
				for (int i = 0; i < fieldCount; i++)
				{
					// If no caption is set, the ColumnName property is called implicitly.
					this.myWorksheet.SetValueInner(row, col++, reader.GetName(i));
				}
				row++;
				col = this._fromCol;
			}
			while (reader.Read())
			{
				for (int i = 0; i < fieldCount; i++)
				{
					this.myWorksheet.SetValueInner(row, col++, reader.GetValue(i));
				}
				row++;
				col = this._fromCol;
			}
			return this.myWorksheet.Cells[this._fromRow, this._fromCol, row - 1, this._fromCol + fieldCount - 1];
		}

		/// <summary>
		/// Load the data from the datatable starting from the top left cell of the range
		/// </summary>
		/// <param name="table">The datatable to load</param>
		/// <param name="printHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
		/// <param name="tableStyle">The table style to apply to the data</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromDataTable(DataTable table, bool printHeaders, TableStyles tableStyle)
		{
			var range = LoadFromDataTable(table, printHeaders);
			int rows = (table.Rows.Count == 0 ? 1 : table.Rows.Count) + (printHeaders ? 1 : 0);
			if (rows >= 0 && table.Columns.Count > 0)
			{
				var excelTable = this.myWorksheet.Tables.Add(new ExcelAddress(this._fromRow, this._fromCol, this._fromRow + rows - 1, this._fromCol + table.Columns.Count - 1), table.TableName);
				excelTable.ShowHeader = printHeaders;
				excelTable.TableStyle = tableStyle;
			}
			return range;
		}
		/// <summary>
		/// Load the data from the datatable starting from the top left cell of the range
		/// </summary>
		/// <param name="table">The datatable to load</param>
		/// <param name="printHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromDataTable(DataTable table, bool printHeaders)
		{
			if (table == null)
				throw (new ArgumentNullException("Table can't be null"));
			if (table.Rows.Count == 0)
				return null;
			var rowArray = new List<object[]>();
			if (printHeaders)
				rowArray.Add(table.Columns.Cast<DataColumn>().Select((dc) => { return dc.Caption; }).ToArray());
			foreach (DataRow dr in table.Rows)
			{
				rowArray.Add(dr.ItemArray);
			}
			for (int column = 0; column < table.Columns.Count; column++)
				for (int row = 0; row < rowArray.Count; row++)
				{
					var val = rowArray[row][column];
					if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
						this.Worksheet.SetValue(row + this._fromRow, column + this._fromCol, val);
				}
			return this.myWorksheet.Cells[this._fromRow, this._fromCol, this._fromRow + rowArray.Count - 1, this._fromCol + table.Columns.Count - 1];
		}

		/// <summary>
		/// Loads data from the collection of arrays of objects into the range, starting from
		/// the top-left cell.
		/// </summary>
		/// <param name="data">The data.</param>
		public ExcelRangeBase LoadFromArrays(IEnumerable<object[]> data)
		{
			if (data == null)
				throw new ArgumentNullException("data");
			var rowArray = new List<object[]>();
			var maxColumn = 0;
			foreach (object[] item in data)
			{
				rowArray.Add(item);
				if (maxColumn < item.Length)
					maxColumn = item.Length;
			}
			for (int row = 0; row < rowArray.Count; row++)
			{
				var currentRow = rowArray[row];
				for (int column = 0; column < currentRow.Length; column++)
				{
					var val = currentRow[column];
					if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
					{
						this.Worksheet.SetValue(row + this._fromRow, column + this._fromCol, val);
					}
				}
			}
			return this.myWorksheet.Cells[this._fromRow, this._fromCol, this._fromRow + rowArray.Count - 1, this._fromCol + maxColumn - 1];
		}

		/// <summary>
		/// Load a collection into a the worksheet starting from the top left row of the range.
		/// </summary>
		/// <typeparam name="T">The datatype in the collection</typeparam>
		/// <param name="collection">The collection to load</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> collection)
		{
			return LoadFromCollection<T>(collection, false, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
		}

		/// <summary>
		/// Load a collection of T into the worksheet starting from the top left row of the range.
		/// Default option will load all public instance properties of T
		/// </summary>
		/// <typeparam name="T">The datatype in the collection</typeparam>
		/// <param name="collection">The collection to load</param>
		/// <param name="printHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> collection, bool printHeaders)
		{
			return LoadFromCollection<T>(collection, printHeaders, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
		}

		/// <summary>
		/// Load a collection of T into the worksheet starting from the top left row of the range.
		/// Default option will load all public instance properties of T
		/// </summary>
		/// <typeparam name="T">The datatype in the collection</typeparam>
		/// <param name="collection">The collection to load</param>
		/// <param name="printHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
		/// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> collection, bool printHeaders, TableStyles tableStyle)
		{
			return LoadFromCollection<T>(collection, printHeaders, tableStyle, BindingFlags.Public | BindingFlags.Instance, null);
		}

		/// <summary>
		/// Load a collection into the worksheet starting from the top left row of the range.
		/// </summary>
		/// <typeparam name="T">The datatype in the collection</typeparam>
		/// <param name="collection">The collection to load</param>
		/// <param name="printHeaders">Print the property names on the first row. Any underscore in the property name will be converted to a space. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
		/// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
		/// <param name="memberFlags">Property flags to use</param>
		/// <param name="members">The properties to output. Must be of type T</param>
		/// <returns>The filled range</returns>
		public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> collection, bool printHeaders, TableStyles tableStyle, BindingFlags memberFlags, MemberInfo[] members)
		{
			var type = typeof(T);
			if (members == null)
				members = type.GetProperties(memberFlags);
			else
			{
				foreach (var t in members)
				{
					if (t.DeclaringType != null && t.DeclaringType != type && !t.DeclaringType.IsSubclassOf(type))
						throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T");
				}
			}
			// create buffer
			object[,] values = new object[(printHeaders ? collection.Count() + 1 : collection.Count()), members.Count()];
			int col = 0, row = 0;
			if (members.Length > 0 && printHeaders)
			{
				foreach (var t in members)
				{
					var descriptionAttribute = t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() as DescriptionAttribute;
					var header = string.Empty;
					if (descriptionAttribute != null)
						header = descriptionAttribute.Description;
					else
					{
						var displayNameAttribute =
							t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as
							DisplayNameAttribute;
						if (displayNameAttribute != null)
							header = displayNameAttribute.DisplayName;
						else
							header = t.Name.Replace('_', ' ');
					}
					values[row, col++] = header;
				}
				row++;
			}

			if (!collection.Any() && (members.Length == 0 || printHeaders == false))
				return null;

			if (members.Length == 0)
			{
				foreach (var item in collection)
				{
					values[row, col++] = item;
				}
			}
			else
			{
				foreach (var item in collection)
				{
					col = 0;
					if (item is string || item is decimal || item is DateTime || item.GetType().IsPrimitive)
						values[row, col++] = item;
					else
					{
						foreach (var t in members)
						{
							if (t is PropertyInfo)
								values[row, col++] = ((PropertyInfo)t).GetValue(item, null);
							else if (t is FieldInfo)
								values[row, col++] = ((FieldInfo)t).GetValue(item);
							else if (t is MethodInfo)
								values[row, col++] = ((MethodInfo)t).Invoke(item, null);
						}
					}
					row++;
				}
			}

			this.myWorksheet.SetRangeValueInner(this._fromRow, this._fromCol, this._fromRow + row - 1, this._fromCol + col - 1, values);

			//Must have at least 1 row, if header is showen
			if (row == 1 && printHeaders)
				row++;

			var r = this.myWorksheet.Cells[this._fromRow, this._fromCol, this._fromRow + row - 1, this._fromCol + col - 1];

			if (tableStyle != TableStyles.None)
			{
				var tbl = this.myWorksheet.Tables.Add(r, "");
				tbl.ShowHeader = printHeaders;
				tbl.TableStyle = tableStyle;
			}
			return r;
		}

		#region LoadFromText
		/// <summary>
		/// Loads a CSV text into a range starting from the top left cell.
		/// Default settings is Comma separation
		/// </summary>
		/// <param name="text">The Text</param>
		/// <returns>The range containing the data</returns>
		public ExcelRangeBase LoadFromText(string text)
		{
			return LoadFromText(text, new ExcelTextFormat());
		}
		/// <summary>
		/// Loads a CSV text into a range starting from the top left cell.
		/// </summary>
		/// <param name="text">The Text</param>
		/// <param name="format">Information how to load the text</param>
		/// <returns>The range containing the data</returns>
		public ExcelRangeBase LoadFromText(string text, ExcelTextFormat format)
		{
			if (string.IsNullOrEmpty(text))
			{
				var range = this.myWorksheet.Cells[this._fromRow, this._fromCol];
				range.Value = string.Empty;
				return range;
			}
			if (format == null)
				format = new ExcelTextFormat();
			string splitRegex = string.Format("{0}(?=(?:[^{1}]*{1}[^{1}]*{1})*[^{1}]*$)", format.EOL, format.TextQualifier);
			string[] lines = Regex.Split(text, splitRegex);
			int row = 0;
			int col = 0;
			int maxCol = col;
			int lineNo = 1;
			var values = new List<object>[lines.Length];
			foreach (string line in lines)
			{
				var items = new List<object>();
				values[row] = items;

				if (lineNo > format.SkipLinesBeginning && lineNo <= lines.Length - format.SkipLinesEnd)
				{
					col = 0;
					string value = string.Empty;
					bool isText = false, isQualifier = false;
					int QCount = 0;
					int lineQCount = 0;
					foreach (char character in line)
					{
						if (format.TextQualifier != 0 && character == format.TextQualifier)
						{
							if (!isText && value != string.Empty)
								throw (new Exception(string.Format("Invalid Text Qualifier in line : {0}", line)));
							isQualifier = !isQualifier;
							QCount += 1;
							lineQCount++;
							isText = true;
						}
						else
						{
							if (QCount > 1 && !string.IsNullOrEmpty(value))
								value += new string(format.TextQualifier, QCount / 2);
							else if (QCount > 2 && string.IsNullOrEmpty(value))
								value += new string(format.TextQualifier, (QCount - 1) / 2);

							if (isQualifier)
								value += character;
							else
							{
								if (character == format.Delimiter)
								{
									items.Add(ConvertData(format, value, col, isText));
									value = string.Empty;
									isText = false;
									col++;
								}
								else
								{
									if (QCount % 2 == 1)
										throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
									value += character;
								}
							}
							QCount = 0;
						}
					}
					if (QCount > 1)
					{
						value += new string(format.TextQualifier, QCount / 2);
					}
					if (lineQCount % 2 == 1)
						throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
					items.Add(ConvertData(format, value, col, isText));
					if (col > maxCol)
						maxCol = col;
					row++;
				}
				lineNo++;
			}

			// Flush
			for (int myRowIndex = 0; row < values.Length; row++)
			{
				var element = values[myRowIndex];
				for (int column = 0; column < element.Count; column++)
				{
					var val = element[column];
					if (val != null)
					{
						this.Worksheet.SetValue(myRowIndex + this._fromRow, column + this._fromCol, val);
					}
				}
			}
			return this.myWorksheet.Cells[this._fromRow, this._fromCol, this._fromRow + row, this._fromCol + maxCol];
		}

		/// <summary>
		/// Loads a CSV text into a range starting from the top left cell.
		/// </summary>
		/// <param name="text">The Text</param>
		/// <param name="format">Information how to load the text</param>
		/// <param name="tableStyle">Create a table with this style</param>
		/// <param name="firstRowIsHeader">Use the first row as header</param>
		/// <returns></returns>
		public ExcelRangeBase LoadFromText(string text, ExcelTextFormat format, TableStyles tableStyle, bool firstRowIsHeader)
		{
			var range = LoadFromText(text, format);
			var table = this.myWorksheet.Tables.Add(range, "");
			table.ShowHeader = firstRowIsHeader;
			table.TableStyle = tableStyle;
			return range;
		}

		/// <summary>
		/// Loads a CSV file into a range starting from the top left cell.
		/// </summary>
		/// <param name="textFile">The Textfile</param>
		/// <returns></returns>
		public ExcelRangeBase LoadFromText(FileInfo textFile)
		{
			return LoadFromText(File.ReadAllText(textFile.FullName, Encoding.ASCII));
		}

		/// <summary>
		/// Loads a CSV file into a range starting from the top left cell.
		/// </summary>
		/// <param name="textFile">The Textfile</param>
		/// <param name="format">Information how to load the text</param>
		/// <returns></returns>
		public ExcelRangeBase LoadFromText(FileInfo textFile, ExcelTextFormat format)
		{
			return LoadFromText(File.ReadAllText(textFile.FullName, format.Encoding), format);
		}

		/// <summary>
		/// Loads a CSV file into a range starting from the top left cell.
		/// </summary>
		/// <param name="textFile">The Textfile</param>
		/// <param name="format">Information how to load the text</param>
		/// <param name="tableStyle">Create a table with this style</param>
		/// <param name="firstRowIsHeader">Use the first row as header</param>
		/// <returns></returns>
		public ExcelRangeBase LoadFromText(FileInfo textFile, ExcelTextFormat format, TableStyles tableStyle, bool firstRowIsHeader)
		{
			return LoadFromText(File.ReadAllText(textFile.FullName, format.Encoding), format, tableStyle, firstRowIsHeader);
		}
		#endregion

		#region GetValue
		/// <summary>
		/// Get the strongly typed value of the cell.
		/// </summary>
		/// <typeparam name="T">The type</typeparam>
		/// <returns>The value. If the value can't be converted to the specified type, the default value will be returned</returns>
		public T GetValue<T>()
		{
			return this.myWorksheet.GetTypedValue<T>(this.Value);
		}
		#endregion

		/// <summary>
		/// Get a range with an offset from the top left cell.
		/// The new range has the same dimensions as the current range
		/// </summary>
		/// <param name="rowOffset">Row Offset</param>
		/// <param name="columnOffset">Column Offset</param>
		/// <returns></returns>
		public ExcelRangeBase Offset(int rowOffset, int columnOffset)
		{
			if (this._fromRow + rowOffset < 1 || this._fromCol + columnOffset < 1 || this._fromRow + rowOffset > ExcelPackage.MaxRows || this._fromCol + columnOffset > ExcelPackage.MaxColumns)
				throw (new ArgumentOutOfRangeException("Offset value out of range"));
			string address = GetAddress(this._fromRow + rowOffset, this._fromCol + columnOffset, this._toRow + rowOffset, this._toCol + columnOffset);
			return new ExcelRangeBase(this.myWorksheet, address);
		}

		/// <summary>
		/// Get a range with an offset from the top left cell.
		/// </summary>
		/// <param name="rowOffset">Row Offset</param>
		/// <param name="columnOffset">Column Offset</param>
		/// <param name="numberOfRows">Number of rows. Minimum 1</param>
		/// <param name="numberOfColumns">Number of colums. Minimum 1</param>
		/// <returns></returns>
		public ExcelRangeBase Offset(int rowOffset, int columnOffset, int numberOfRows, int numberOfColumns)
		{
			if (numberOfRows < 1 || numberOfColumns < 1)
				throw (new Exception("Number of rows/columns must be greater than 0"));
			numberOfRows--;
			numberOfColumns--;
			if (this._fromRow + rowOffset < 1 || this._fromCol + columnOffset < 1 || this._fromRow + rowOffset > ExcelPackage.MaxRows 
				|| this._fromCol + columnOffset > ExcelPackage.MaxColumns || this._fromRow + rowOffset + numberOfRows < 1 || this._fromCol + columnOffset + numberOfColumns < 1 
				|| this._fromRow + rowOffset + numberOfRows > ExcelPackage.MaxRows || this._fromCol + columnOffset + numberOfColumns > ExcelPackage.MaxColumns)
			{
				throw (new ArgumentOutOfRangeException("Offset value out of range"));
			}
			string address = GetAddress(this._fromRow + rowOffset, this._fromCol + columnOffset, this._fromRow + rowOffset + numberOfRows, this._fromCol + columnOffset + numberOfColumns);
			return new ExcelRangeBase(this.myWorksheet, address);
		}

		/// <summary>
		/// Adds a new comment for the range.
		/// If this range contains more than one cell, the top left comment is returned by the method.
		/// </summary>
		/// <param name="text"></param>
		/// <param name="author"></param>
		/// <returns>A reference comment of the top left cell</returns>
		public ExcelComment AddComment(string text, string author)
		{
			if (string.IsNullOrEmpty(author))
				author = Thread.CurrentPrincipal.Identity.Name;
			if (!this.myWorksheet._commentsStore.Exists(this._fromRow, this._fromCol))
				this.myWorksheet.Comments.Add(new ExcelRangeBase(this.myWorksheet, ExcelRangeBase.GetAddress(this._fromRow, this._fromCol)), text, author);
			else
				throw new Exception("Comment already exists in cell.");
			return this.myWorksheet.Comments[new ExcelCellAddress(this._fromRow, this._fromCol)];
		}

		/// <summary>
		/// Adds a new comment for the range with the same style as the specified <paramref name="comment"/>.
		/// If this range contains more than one cell, the top left comment is returned by the method.
		/// </summary>
		/// <param name="comment">The comment to copy.</param>
		/// <returns>A reference comment of the top left cell.</returns>
		public ExcelComment AddComment(ExcelComment comment)
		{
			if (!this.myWorksheet._commentsStore.Exists(this._fromRow, this._fromCol))
				this.myWorksheet.Comments.Add(new ExcelRangeBase(this.myWorksheet, ExcelRangeBase.GetAddress(this._fromRow, this._fromCol)), comment);
			else
				throw new Exception("Comment already exists in cell.");
			return this.myWorksheet.Comments[new ExcelCellAddress(this._fromRow, this._fromCol)];
		}

		/// <summary>
		/// Sets the formula of the cell to the given formula, only allowing the value of the cell
		/// to be cleared if specified.
		/// </summary>
		/// <param name="formula">The formula to set.</param>
		/// <param name="clearValue">The formula to set.</param>
		public void SetFormula(string formula, bool clearValue)
		{
			if (!string.IsNullOrEmpty(formula) && formula[0] == '=')
				formula = formula.Substring(1);
			if ((formula == null || formula.Trim() == string.Empty) && clearValue)
				this.Value = null;
			else if (this._fromRow == this._toRow && this._fromCol == this._toCol)
				SetFormula(this, formula, this._fromRow, this._fromCol, clearValue);
			else
			{
				SetSharedFormula(this, formula, this, false, clearValue);
				if (this.Addresses != null)
				{
					foreach (var address in this.Addresses)
					{
						SetSharedFormula(this, formula, address, false, clearValue);
					}
				}
			}
		}

		/// <summary>
		/// Copies the range of cells to an other range
		/// </summary>
		/// <param name="Destination">The start cell where the range will be copied.</param>
		public void Copy(ExcelRangeBase Destination)
		{
			Copy(Destination, null);
		}

		/// <summary>
		/// Copies the range of cells to an other range
		/// </summary>
		/// <param name="destination">The start cell where the range will be copied.</param>
		/// <param name="excelRangeCopyOptionFlags">Cell parts that will not be copied. If Formulas are specified, the formulas will NOT be copied.</param>
		public void Copy(ExcelRangeBase destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
		{
			bool sameWorkbook = destination.myWorksheet.Workbook == this.myWorksheet.Workbook;
			ExcelStyles sourceStyles = this.myWorksheet.Workbook.Styles,
						styles = destination.myWorksheet.Workbook.Styles;
			Dictionary<int, int> styleCashe = new Dictionary<int, int>();
			//Clear all existing cells; 
			int toRow = this._toRow - this._fromRow + 1,
				toCol = this._toCol - this._fromCol + 1;

			int i = 0;
			object o = null;
			byte flag = 0;
			Uri hl = null;

			var excludeFormulas = excelRangeCopyOptionFlags.HasValue && (excelRangeCopyOptionFlags.Value & ExcelRangeCopyOptionFlags.ExcludeFormulas) == ExcelRangeCopyOptionFlags.ExcludeFormulas;
			var cse = this.myWorksheet._values.GetEnumerator(this._fromRow, this._fromCol, this._toRow, this._toCol);

			var copiedValue = new List<CopiedCell>();
			while (cse.MoveNext())
			{
				var row = cse.Row;
				var col = cse.Column;       //Issue 15070
				var cell = new CopiedCell
				{
					Row = destination._fromRow + (row - this._fromRow),
					Column = destination._fromCol + (col - this._fromCol),
					Value = cse.Value._value
				};

				if (!excludeFormulas && this.myWorksheet._formulas.Exists(row, col, out o))
				{
					if (o is int)
						cell.Formula = this.myWorksheet.GetFormula(cse.Row, cse.Column);
					else
						cell.Formula = o;
				}
				if (this.myWorksheet.ExistsStyleInner(row, col, ref i))
				{
					if (sameWorkbook)
						cell.StyleID = i;
					else
					{
						if (styleCashe.ContainsKey(i))
							i = styleCashe[i];
						else
						{
							var oldStyleID = i;
							i = styles.CloneStyle(sourceStyles, i);
							styleCashe.Add(oldStyleID, i);
						}
						cell.StyleID = i;
					}
				}

				if (this.myWorksheet._hyperLinks.Exists(row, col, out hl))
					cell.HyperLink = hl;

				if (this.myWorksheet._flags.Exists(row, col, out flag))
					cell.Flag = flag;
				copiedValue.Add(cell);
			}

			//Copy styles with no cell value
			var cses = this.myWorksheet._values.GetEnumerator(this._fromRow, this._fromCol, this._toRow, this._toCol);
			while (cses.MoveNext())
			{
				if (!this.myWorksheet.ExistsValueInner(cses.Row, cses.Column))
				{
					var row = destination._fromRow + (cses.Row - this._fromRow);
					var col = destination._fromCol + (cses.Column - this._fromCol);
					var cell = new CopiedCell
					{
						Row = row,
						Column = col,
						Value = null
					};

					i = cses.Value._styleId;
					if (sameWorkbook)
						cell.StyleID = i;
					else
					{
						if (styleCashe.ContainsKey(i))
							i = styleCashe[i];
						else
						{
							var oldStyleID = i;
							i = styles.CloneStyle(sourceStyles, i);
							styleCashe.Add(oldStyleID, i);
						}
						//Destination._worksheet.SetStyleInner(row, col, i);
						cell.StyleID = i;
					}
					copiedValue.Add(cell);
				}
			}
			var copiedMergedCells = new Dictionary<int, ExcelAddress>();
			var csem = this.myWorksheet.MergedCells.Cells.GetEnumerator(this._fromRow, this._fromCol, this._toRow, this._toCol);
			while (csem.MoveNext())
			{
				if (!copiedMergedCells.ContainsKey(csem.Value))
				{
					var adr = new ExcelAddress(this.myWorksheet.Name, this.myWorksheet.MergedCells.List[csem.Value]);
					if (this.Collide(adr) == eAddressCollition.Inside)
					{
						copiedMergedCells.Add(csem.Value, new ExcelAddress(
							destination._fromRow + (adr.Start.Row - this._fromRow),
							destination._fromCol + (adr.Start.Column - this._fromCol),
							destination._fromRow + (adr.End.Row - this._fromRow),
							destination._fromCol + (adr.End.Column - this._fromCol)));
					}
					else
					{
						//Partial merge of the address ignore.
						copiedMergedCells.Add(csem.Value, null);
					}
				}
			}
			var copiedCommentCells = new List<CopiedCell>();
			var csec = this.myWorksheet._commentsStore.GetEnumerator(this._fromRow, this._fromCol, this._toRow, this._toCol);
			while (csec.MoveNext())
			{
				var row = destination._fromRow + (csec.Row - this._fromRow);
				var column = destination._fromCol + (csec.Column - this._fromCol);
				var cell = new CopiedCell
				{
					Row = row,
					Column = column,
					Value = null
				};
				// Will just be null if no comment exists.
				cell.Comment = this.myWorksheet.Cells[csec.Row, csec.Column].Comment;
				if (cell.Comment != null)
					copiedCommentCells.Add(cell);
			}
			destination.myWorksheet._values.Clear(destination._fromRow, destination._fromCol, toRow, toCol);
			destination.myWorksheet._formulas.Clear(destination._fromRow, destination._fromCol, toRow, toCol);
			destination.myWorksheet._hyperLinks.Clear(destination._fromRow, destination._fromCol, toRow, toCol);
			destination.myWorksheet._flags.Clear(destination._fromRow, destination._fromCol, toRow, toCol);
			destination.myWorksheet._commentsStore.Clear(destination._fromRow, destination._fromCol, toRow, toCol);

			foreach (var cell in copiedValue)
			{
				destination.myWorksheet.SetValueInner(cell.Row, cell.Column, cell.Value);

				if (cell.StyleID != null)
					destination.myWorksheet.SetStyleInner(cell.Row, cell.Column, cell.StyleID.Value);

				if (cell.Formula != null)
				{
					cell.Formula = this.myWorkbook.Package.FormulaManager.UpdateFormulaReferences(cell.Formula.ToString(), destination._fromRow - this._fromRow, destination._fromCol - this._fromCol, 0, 0, destination.WorkSheet, destination.WorkSheet, true);
					destination.myWorksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
				}
				if (cell.HyperLink != null)
					destination.myWorksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);

				if (cell.Flag != 0)
					destination.myWorksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
			}

			//Add merged cells
			foreach (var mergedCell in copiedMergedCells.Values)
			{
				if (mergedCell != null)
					destination.myWorksheet.MergedCells.Add(mergedCell, true);
			}
			int rowOffset = destination._fromRow;
			// Add comment cells.
			foreach (var cell in copiedCommentCells)
			{
				destination.Worksheet.Cells[cell.Row, cell.Column].AddComment(cell.Comment);
			}
			this.CopySparklines(destination);
			if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns)
			{
				for (int row = 0; row < this.Rows; row++)
				{
					var destinationRow = destination.Worksheet.Row(destination.Start.Row + row);
					destinationRow.OutlineLevel = this.Worksheet.Row(this._fromRow + row).OutlineLevel;
				}
			}
			if (this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows)
			{
				for (int column = 0; column < this.Columns; column++)
				{
					var destinationCol = destination.Worksheet.Column(destination.Start.Column + column);
					destinationCol.OutlineLevel = this.Worksheet.Column(this._fromCol + column).OutlineLevel;
				}
			}
		}

		/// <summary>
		/// Clear all cells
		/// </summary>
		public void Clear()
		{
			Clear(this);
		}

		/// <summary>
		/// Creates an array-formula.
		/// </summary>
		/// <param name="arrayFormula">The formula</param>
		public void CreateArrayFormula(string arrayFormula)
		{
			if (this.Addresses != null)
				throw (new Exception("An Arrayformula can not have more than one address"));
			SetSharedFormula(this, arrayFormula, this, true);
		}

		internal void Clear(ExcelAddress range)
		{
			this.myWorksheet.MergedCells.Clear(range);
			//First find the start cell
			int fromRow, fromCol;
			var d = this.Worksheet.Dimension;
			if (d != null && range._fromRow <= d._fromRow && range._toRow >= d._toRow) //EntireRow?
				fromRow = 0;
			else
				fromRow = range._fromRow;
			if (d != null && range._fromCol <= d._fromCol && range._toCol >= d._toCol) //EntireRow?
				fromCol = 0;
			else
				fromCol = range._fromCol;

			var rows = range._toRow - fromRow + 1;
			var cols = range._toCol - fromCol + 1;

			this.myWorksheet._values.Clear(fromRow, fromCol, rows, cols);
			//_worksheet._types.Delete(fromRow, fromCol, rows, cols, shift);
			//_worksheet._styles.Delete(fromRow, fromCol, rows, cols, shift);
			this.myWorksheet._formulas.Clear(fromRow, fromCol, rows, cols);
			this.myWorksheet._hyperLinks.Clear(fromRow, fromCol, rows, cols);
			this.myWorksheet._flags.Clear(fromRow, fromCol, rows, cols);
			this.myWorksheet._commentsStore.Clear(fromRow, fromCol, rows, cols);

			//Clear multi addresses as well
			if (this.Addresses != null)
			{
				foreach (var sub in this.Addresses)
				{
					Clear(sub);
				}
			}
		}

		/// <summary>
		/// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
		/// Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine.
		/// Wrapped and merged cells are also ignored.
		/// </summary>
		public void AutoFitColumns()
		{
			AutoFitColumns(this.myWorksheet.DefaultColWidth);
		}

		/// <summary>
		/// Set the column width from the content of the range.
		/// Note: Cells containing formulas are ignored if no calculation is made.
		///       Wrapped and merged cells are also ignored.
		/// </summary>
		/// <remarks>This method will not work if you run in an environment that does not support GDI</remarks>
		/// <param name="minimumWidth">Minimum column width</param>
		public void AutoFitColumns(double minimumWidth)
		{
			AutoFitColumns(minimumWidth, double.MaxValue);
		}

		/// <summary>
		/// Set the column width from the content of the range.
		/// Note: Cells containing formulas are ignored if no calculation is made.
		///      Merged cells are also ignored.
		///      Hidden columns are left hidden.
		/// </summary>
		/// <param name="MinimumWidth">Minimum column width.</param>
		/// <param name="MaximumWidth">Maximum column width.</param>
		public void AutoFitColumns(double MinimumWidth, double MaximumWidth)
		{
			if (this.myWorksheet.Dimension == null)
				return;
			if (this._fromCol < 1 || this._fromRow < 1)
				SetToSelectedRange();
			var fontCache = new Dictionary<int, Font>();
			bool doAdjust = this.myWorksheet.Package.DoAdjustDrawings;
			this.myWorksheet.Package.DoAdjustDrawings = false;
			var drawWidths = this.myWorksheet.Drawings.GetDrawingWidths();

			var fromCol = this._fromCol > this.myWorksheet.Dimension._fromCol ? this._fromCol : this.myWorksheet.Dimension._fromCol;
			var toCol = this._toCol < this.myWorksheet.Dimension._toCol ? this._toCol : this.myWorksheet.Dimension._toCol;

			if (fromCol > toCol)
				return; //Issue 15383

			if (this.Addresses == null)
				SetMinWidth(MinimumWidth, fromCol, toCol);
			else
			{
				foreach (var addr in this.Addresses)
				{
					fromCol = addr._fromCol > this.myWorksheet.Dimension._fromCol ? addr._fromCol : this.myWorksheet.Dimension._fromCol;
					toCol = addr._toCol < this.myWorksheet.Dimension._toCol ? addr._toCol : this.myWorksheet.Dimension._toCol;
					SetMinWidth(MinimumWidth, fromCol, toCol);
				}
			}

			//Get any autofilter to widen these columns
			var afAddr = new List<ExcelAddress>();
			if (this.myWorksheet.AutoFilterAddress != null)
			{
				afAddr.Add(new ExcelAddress(this.myWorksheet.AutoFilterAddress._fromRow,
													this.myWorksheet.AutoFilterAddress._fromCol,
													this.myWorksheet.AutoFilterAddress._fromRow,
													this.myWorksheet.AutoFilterAddress._toCol));
				afAddr[afAddr.Count - 1]._ws = this.WorkSheet;
			}
			foreach (var table in this.myWorksheet.Tables)
			{
				if (table.AutoFilterAddress != null)
				{
					afAddr.Add(new ExcelAddress(table.AutoFilterAddress._fromRow,
																			table.AutoFilterAddress._fromCol,
																			table.AutoFilterAddress._fromRow,
																			table.AutoFilterAddress._toCol));
					afAddr[afAddr.Count - 1]._ws = this.WorkSheet;
				}
			}

			var styles = this.myWorksheet.Workbook.Styles;
			var normal = styles.Fonts[styles.CellXfs[0].FontId];
			var normalStyle = FontStyle.Regular;
			if (normal.Bold) normalStyle |= FontStyle.Bold;
			if (normal.UnderLine) normalStyle |= FontStyle.Underline;
			if (normal.Italic) normalStyle |= FontStyle.Italic;
			if (normal.Strike) normalStyle |= FontStyle.Strikeout;
			var defaultFont = new Font(normal.Name, normal.Size, normalStyle);

			using (Graphics g = Graphics.FromImage(new Bitmap(20, 20)))
			{
				var defaultCharacterWidth = Enumerable.Range(0, 10).Select(i => g.MeasureString(i.ToString(), defaultFont).Width).Average();
				foreach (var cell in this)
				{
					if (this.myWorksheet.Column(cell.Start.Column).Hidden)    //Issue 15338
						continue;

					if (cell.Merge == true) continue;
					var fntID = styles.CellXfs[cell.StyleID].FontId;
					Font font;
					if (fontCache.ContainsKey(fntID))
						font = fontCache[fntID];
					else
					{
						var fnt = styles.Fonts[fntID];
						var fs = FontStyle.Regular;
						if (fnt.Bold) fs |= FontStyle.Bold;
						if (fnt.UnderLine) fs |= FontStyle.Underline;
						if (fnt.Italic) fs |= FontStyle.Italic;
						if (fnt.Strike) fs |= FontStyle.Strikeout;
						font = new Font(fnt.Name, fnt.Size, fs);
						fontCache.Add(fntID, font);
					}

					var cellFonts = new List<Font> { font };
					foreach (var cFormat in this.myWorksheet.ConditionalFormatting.Where(format => this.IsCellInConditionalFormatAddress(cell, format)))
					{
						var cFormatFont = cFormat.Style?.Font;
						if (cFormatFont == null)
							continue;
						var fs = FontStyle.Regular;
						if (cFormatFont.Bold.HasValue && cFormatFont.Bold.Value) fs |= FontStyle.Bold;
						if (cFormatFont.Underline.HasValue && cFormatFont.Underline.Value != ExcelUnderLineType.None) fs |= FontStyle.Underline;
						if (cFormatFont.Italic.HasValue && cFormatFont.Italic.Value) fs |= FontStyle.Italic;
						if (cFormatFont.Strike.HasValue && cFormatFont.Strike.Value) fs |= FontStyle.Strikeout;
						// Conditional formatting doesn't modify font or size, just style
						cellFonts.Add(new Font(font.Name, font.Size, fs));
					}

					double r = styles.CellXfs[cell.StyleID].TextRotation;
					List<double> sizes = new List<double>();
					var cellStyle = styles.CellXfs[cell.StyleID];
					var cellPadding = cellStyle.WrapText ? 0 : 2;
					var cellIndent = cellStyle.Indent;
					foreach (var cellFont in cellFonts)
					{
						var textForWidth = cell.GetTextForWidth(cellFont);
						var characterCount = textForWidth.ToCharArray().Count() + (cellIndent > 0 && !string.IsNullOrEmpty(textForWidth) ? cellIndent : 0);
						var measurableText = new string('W', characterCount + cellPadding);
						var textWidth = g.MeasureString(measurableText, cellFont).Width / defaultCharacterWidth;
						if (r == 0)
							sizes.Add(textWidth);
						else
							sizes.Add(Convert.ToSingle(Math.Cos(Math.PI * (r <= 90 ? r : r - 90) / 180.0) * textWidth));
					}

					var width = sizes.Max();
					foreach (var a in afAddr)
					{
						if (a.Collide(cell) != eAddressCollition.No)
						{
							//width += 2.8;
							width += 2.25;
							break;
						}
					}
					if (width > this.myWorksheet.Column(cell._fromCol).Width)
						this.myWorksheet.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
				}
			}
			this.myWorksheet.Drawings.AdjustWidth(drawWidths);
			this.myWorksheet.Package.DoAdjustDrawings = doAdjust;
		}
		#endregion

		#region Enumerator
		ICellStoreEnumerator<ExcelCoreValue> cellEnum;

		/// <summary>
		/// Returns an enumerator that iterates through the collection.
		/// </summary>
		/// <returns>An IEnumerator object that can be used to iterate through the collection of <see cref="ExcelRangeBase"/>.</returns>
		public IEnumerator<ExcelRangeBase> GetEnumerator()
		{
			Reset();
			return this;
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			Reset();
			return this;
		}

		/// <summary>
		/// The current range when enumerating
		/// </summary>
		public ExcelRangeBase Current
		{
			get
			{
				return new ExcelRangeBase(this.myWorksheet, ExcelAddress.GetAddress(this.cellEnum.Row, this.cellEnum.Column));
			}
		}

		/// <summary>
		/// The current range when enumerating
		/// </summary>
		object IEnumerator.Current
		{
			get
			{
				return ((new ExcelRangeBase(this.myWorksheet, ExcelAddress.GetAddress(this.cellEnum.Row, this.cellEnum.Column))));
			}
		}

		public object FormatedText { get; private set; }

		int _enumAddressIx = -1;

		/// <summary>
		/// Advances the enumerator to the next element of the collection.
		/// </summary>
		/// <returns>True if the enumerator was successfully advanced to the next element; false if the enumerator has passed the end of the collection.</returns>
		public bool MoveNext()
		{
			if (this.cellEnum.MoveNext())
			{
				return true;
			}
			else if (this._addresses != null)
			{
				this._enumAddressIx++;
				if (this._enumAddressIx < this._addresses.Count)
				{
					this.cellEnum = this.myWorksheet._values.GetEnumerator(
						this._addresses[this._enumAddressIx]._fromRow,
						this._addresses[this._enumAddressIx]._fromCol,
						this._addresses[this._enumAddressIx]._toRow,
						this._addresses[this._enumAddressIx]._toCol);
					return MoveNext();
				}
				else
					return false;
			}
			return false;
		}

		/// <summary>
		/// Sets the enumerator to its initial position, which is before the first element in the collection.
		/// </summary>
		public void Reset()
		{
			this._enumAddressIx = -1;
			this.cellEnum = this.myWorksheet._values.GetEnumerator(this._fromRow, this._fromCol, this._toRow, this._toCol);
		}

		/// <summary>
		/// No-op dispose implementation for IEnumerator.
		/// </summary>
		public void Dispose() { }
		#endregion
	}
}