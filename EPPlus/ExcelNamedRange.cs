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

using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml
{
	/// <summary>
	/// A named range.
	/// </summary>
	public sealed class ExcelNamedRange
	{
		#region Class Variables
		private string myFormula;
		#endregion

		#region Constants
		/// <summary>
		/// A list of known characters that are not allowed in Excel named ranges.
		/// </summary>
		public static readonly char[] IllegalCharacters = 
		{
			'~','!','@','#','$','%','^','&','*','(',')','+','-','=','[',']','{','}','|',
			'<','>','?',',','/',';','\'',':','"','¥','¢','£','«', '»','®','©',' ','	'
		};
		#endregion

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
		/// A comment for the Name
		/// </summary>
		public string NameComment { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this Named range is hidden.
		/// </summary>
		public bool IsNameHidden { get; set; }

		/// <summary>
		/// Gets the <see cref="ExcelWorksheet"/> that this named range is local to. Null if the named range is workbook scoped.
		/// </summary>
		public ExcelWorksheet LocalSheet { get; private set; }

		/// <summary>
		/// Gets the <see cref="ExcelWorkbook"/> that contains this named range.
		/// </summary>
		internal ExcelWorkbook Workbook { get; private set; }

		/// <summary>
		/// Gets or sets the index value of this named range in its parent <see cref="ExcelNamedRangeCollection"/>.
		/// This is used to create ID values for dependency chains.
		/// </summary>
		internal int Index { get; set; }

		/// <summary>
		/// Gets or sets the formula of this Named Range.
		/// </summary>
		public string NameFormula
		{
			get
			{
				return myFormula;
			}
			set
			{
				if (string.IsNullOrEmpty(value))
					throw new InvalidOperationException($"{nameof(this.NameFormula)} cannot be null or empty");
				myFormula = value;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Constructs a new <see cref="ExcelNamedRange"/> object.
		/// </summary>
		/// <param name="name">The name of the range.</param>
		/// <param name="workbook">The workbook that contains this named range.</param>
		/// <param name="nameSheet">The sheet this named range is local to, or null for a global named range.</param>
		/// <param name="formula">The address (range) this named range refers to.</param>
		/// <param name="index">The index of this named range in the parent <see cref="ExcelNamedRangeCollection"/>.</param>
		public ExcelNamedRange(string name, ExcelWorkbook workbook, ExcelWorksheet nameSheet, string formula, int index)
		{
			if (workbook == null)
				throw new ArgumentNullException(nameof(workbook));
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException(nameof(name));
			if (string.IsNullOrEmpty(formula))
				throw new ArgumentNullException(nameof(formula));
			this.Name = name;
			this.Workbook = workbook;
			this.LocalSheet = nameSheet;
			this.NameFormula = formula;
			this.Index = index;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Gets the formula of a named range relative to the specified <paramref name="relativeRow"/> and <paramref name="relativeColumn"/>.
		/// </summary>
		/// <param name="relativeRow">The row from which the named range is referenced.</param>
		/// <param name="relativeColumn">The column from which the named range is referenced.</param>
		/// <returns>The updated formula relative to the specified <paramref name="relativeRow"/> and <paramref name="relativeColumn"/>.</returns>
		public IEnumerable<Token> GetRelativeNameFormula(int relativeRow, int relativeColumn)
		{
			var tokens = this.Workbook.FormulaParser.Lexer.Tokenize(this.NameFormula);
			foreach (var token in tokens)
			{
				if (token.TokenType == TokenType.ExcelAddress)
				{
					var address = new ExcelAddress(token.Value);
					// Do not update external references.
					if (!string.IsNullOrEmpty(address?.Workbook))
						continue;
					int fromRow = address._fromRow;
					int fromColumn = address._fromCol;
					int toRow = address._toRow;
					int toColumn = address._toCol;
					if (!address._isFullColumn)
					{
						fromRow = this.GetRelativeLocation(address._fromRowFixed, address._fromRow, relativeRow, ExcelPackage.MaxRows);
						toRow = this.GetRelativeLocation(address._toRowFixed, address._toRow, relativeRow, ExcelPackage.MaxRows);
					}
					if (!address._isFullRow)
					{
						fromColumn = this.GetRelativeLocation(address._fromColFixed, address._fromCol, relativeColumn, ExcelPackage.MaxColumns);
						toColumn = this.GetRelativeLocation(address._toColFixed, address._toCol, relativeColumn, ExcelPackage.MaxColumns);
					}
					var updatedAddress = ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed);
					token.Value = ExcelCellBase.GetFullAddress(address.WorkSheet, updatedAddress);
				}
			}
			return tokens;
		}

		/// <summary>
		/// Updates the named range's <see cref="NameFormula"/> references according to the 
		/// rows and or columns being inserted and or deleted on the specified <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="rowFrom">The starting row to perform the operation at.</param>
		/// <param name="colFrom">The ending row to perform the operation at.</param>
		/// <param name="rows">The number of rows being inserted.</param>
		/// <param name="cols">The number of columns being inserted.</param>
		/// <param name="worksheet">The worksheet to update.</param>
		public void UpdateFormula(int rowFrom, int colFrom, int rows, int cols, ExcelWorksheet worksheet)
		{
			this.NameFormula = this.Workbook.Package.FormulaManager.UpdateFormulaReferences(
				this.NameFormula, rows, cols, rowFrom, colFrom, worksheet.Name, worksheet.Name, updateOnlyFixed: true);
		}

		/// <summary>
		/// Attempts to parse the <see cref="NameFormula"/> as an address, evaluating reference functions and
		/// nested named ranges as necessary.
		/// </summary>
		/// <returns>The formula as an <see cref="ExcelRangeBase"/> if it is an address, null otherwise.</returns>
		public ExcelRangeBase GetFormulaAsCellRange()
		{
			return AddressUtility.GetFormulaAsCellRange(this.Workbook, this.LocalSheet, this.NameFormula);
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
			return this.Name;
		}
		#endregion
	}
}
