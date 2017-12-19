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
 * Jan Källman		Added this class		        2010-01-28
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/

using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml
{
	/// <summary>
	/// A named range.
	/// </summary>
	public sealed class ExcelNamedRange : ExcelRangeBase
	{
		#region Properties
		/// <summary>
		/// Gets or set the name of this <see cref="ExcelNamedRange"/>.
		/// </summary>
		public string Name { get; internal set; }

		/// <summary>
		/// Returns the "localSheetId" property, which is really the sheet's positionID minus one.
		/// </summary>
		public int LocalSheetID
		{
			get
			{
				return this.LocalSheet == null ? -1 : this.LocalSheet.PositionID - 1;
			}
		}

		/// <summary>
		/// Returns the worksheet's actual "SheetID" property, or -1 if it is not a local named range.
		/// </summary>
		public int ActualSheetID
		{
			get
			{
				return this.LocalSheet == null ? -1 : this.LocalSheet.SheetID;
			}
		}

		/// <summary>
		/// A comment for the Name
		/// </summary>
		public string NameComment { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this Named range is hidden.
		/// </summary>
		public bool IsNameHidden { get; set; }

		/// <summary>
		/// Gets the <see cref="ExcelWorksheet"/> that this named range is local to, or null if the named range has a workbook scope.
		/// </summary>
		internal ExcelWorksheet LocalSheet { get; private set; }

		/// <summary>
		/// Gets or sets the index value of this named range in its parent <see cref="ExcelNamedRangeCollection"/>.
		/// </summary>
		internal int Index { get; set; }

		/// <summary>
		/// Gets or sets the value of this Named Range.
		/// </summary>
		internal object NameValue { get; set; }

		/// <summary>
		/// Gets or sets the formula of this Named Range.
		/// </summary>
		internal string NameFormula { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// A named range
		/// </summary>
		/// <param name="name">The name of the range.</param>
		/// <param name="nameSheet">The sheet this named range is local to, or null for a global named range.</param>
		/// <param name="sheet">The sheet where the target address of the named range exists.</param>
		/// <param name="address">The address (range) this named range refers to.</param>
		/// <param name="index">The index of this named range in the parent <see cref="ExcelNamedRangeCollection"/>.</param>
		public ExcelNamedRange(string name, ExcelWorksheet nameSheet, ExcelWorksheet sheet, string address, int index) :
			 base(sheet, address)
		{
			this.Name = name;
			this.LocalSheet = nameSheet;
			this.Index = index;
		}

		internal ExcelNamedRange(string name, ExcelWorkbook wb, ExcelWorksheet nameSheet, int index) :
			 base(wb, nameSheet, name, true)
		{
			this.Name = name;
			this.LocalSheet = nameSheet;
			this.Index = index;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Gets the address relative to the specified row and column.
		/// </summary>
		/// <param name="relativeRow">The row from which the named range is used.</param>
		/// <param name="relativeColumn">The column from which the named range is used.</param>
		/// <returns>The address relative to the specified row and column.</returns>
		public string GetRelativeAddress(int relativeRow, int relativeColumn)
		{
			// Relative references are relative to cell A1. Offsets that cause the 
			// relative address to exceed the maximum row or column will wrap around.
			// Examples:
			//	$B2 means on column B and down one row from the relative address.
			//	D$5 means on row 5 and right three columns from the relative address.
			//	C3 means right two and down three from the relative address.
			int fromRow = this.GetRelativeLocation(_fromRowFixed, _fromRow, relativeRow, ExcelPackage.MaxRows);
			int fromColumn = this.GetRelativeLocation(_fromColFixed, _fromCol, relativeColumn, ExcelPackage.MaxColumns);
			int toRow = this.GetRelativeLocation(_toRowFixed, _toRow, relativeRow, ExcelPackage.MaxRows);
			int toColumn = this.GetRelativeLocation(_toColFixed, _toCol, relativeColumn, ExcelPackage.MaxColumns);
			var address = ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			return ExcelCellBase.GetFullAddress(this.Worksheet.Name, address);
		}

		/// <summary>
		/// Gets the formula of a named range relative to the specified <paramref name="relativeRow"/> and <paramref name="relativeColumn"/>.
		/// </summary>
		/// <param name="relativeRow">The row from which the named range is referenced.</param>
		/// <param name="relativeColumn">The column from which the named range is referenced.</param>
		/// <returns>The updated formula relative to the specified <paramref name="relativeRow"/> and <paramref name="relativeColumn"/>.</returns>
		public IEnumerable<Token> GetRelativeNameFormula(int relativeRow, int relativeColumn)
		{
			var tokens = this.myWorkbook.FormulaParser.Lexer.Tokenize(this.NameFormula);
			foreach (var token in tokens)
			{
				if (token.TokenType == TokenType.ExcelAddress)
				{
					var address = new ExcelAddress(token.Value);
					int fromRow = this.GetRelativeLocation(address._fromRowFixed, address._fromRow, relativeRow, ExcelPackage.MaxRows);
					int fromColumn = this.GetRelativeLocation(address._fromColFixed, address._fromCol, relativeColumn, ExcelPackage.MaxColumns);
					int toRow = this.GetRelativeLocation(address._toRowFixed, address._toRow, relativeRow, ExcelPackage.MaxRows);
					int toColumn = this.GetRelativeLocation(address._toColFixed, address._toCol, relativeColumn, ExcelPackage.MaxColumns);
					var updatedAddress = ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed);
					token.Value = ExcelCellBase.GetFullAddress(address.WorkSheet, updatedAddress);
				}
			}
			return tokens;
		}
		#endregion

		#region Private Methods
		private int GetRelativeLocation(bool fixedLocation, int current, int relative, int maximum)
		{
			if (fixedLocation)
				return current;
			int row = current + relative - 1;
			if (row > maximum)
				return row - maximum;
			return row;
		}
		#endregion

		#region System.Object Overrides
		public override string ToString()
		{
			return Name;
		}
		#endregion
	}
}
