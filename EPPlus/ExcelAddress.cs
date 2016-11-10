﻿/*******************************************************************************
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
*******************************************************************************
* Jan Källman		Added		18-MAR-2010
* Jan Källman		License changed GPL-->LGPL 2011-12-16
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml
{
	/// <summary>
	/// A range address
	/// </summary>
	/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
	public class ExcelAddressBase : ExcelCellBase
	{
		#region Class Variables
		protected internal int _fromRow = -1, _toRow, _fromCol, _toCol;
		protected internal bool _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed;
		protected internal string _wb;
		protected internal string _ws;
		protected internal string _address;
		protected ExcelCellAddress _start = null;
		protected ExcelCellAddress _end = null;
		private string _firstAddress;
		internal protected List<ExcelAddress> _addresses = null;

		private static readonly HashSet<char> FormulaCharacters = new HashSet<char>(new char[] { '(', ')', '+', '-', '*', '/', '=', '^', '&', '%', '\"' });
		#endregion

		#region Properties
		/// <summary>
		/// Gets the row and column of the top left cell.
		/// </summary>
		/// <value>The start row column.</value>
		public ExcelCellAddress Start
		{
			get
			{
				if (_start == null)
				{
					_start = new ExcelCellAddress(_fromRow, _fromCol);
				}
				return _start;
			}
		}

		/// <summary>
		/// Gets the row and column of the bottom right cell.
		/// </summary>
		/// <value>The end row column.</value>
		public ExcelCellAddress End
		{
			get
			{
				if (_end == null)
				{
					_end = new ExcelCellAddress(_toRow, _toCol);
				}
				return _end;
			}
		}

		/// <summary>
		/// Gets a boolean value indicating whether or not this represents a table address.
		/// </summary>
		public bool IsTableAddress { get { return this.Table != null; } }

		/// <summary>
		/// The address for the range
		/// </summary>
		public virtual string Address
		{
			get
			{
				return _address;
			}
		}

		/// <summary>
		/// Gets the count of rows in this range.
		/// </summary>
		public int Rows
		{
			get
			{
				return _toRow - _fromRow + 1;
			}
		}

		/// <summary>
		/// Gets the count of columns in this range.
		/// </summary>
		public int Columns
		{
			get
			{
				return _toCol - _fromCol + 1;
			}
		}

		/// <summary>
		/// Gets a boolean value indicating whether or not this range is a single cell.
		/// </summary>
		public bool IsSingleCell
		{
			get
			{
				return (_fromRow == _toRow && _fromCol == _toCol);
			}
		}

		/// <summary>
		/// If the address is a defined name
		/// </summary>
		public bool IsName
		{
			get
			{
				return _fromRow < 0;
			}
		}

		/// <summary>
		/// Gets the worksheet name if one exists.
		/// NOTE: Do not refactor this to "Worksheet" as that property exists on a subclass
		/// and is a different type. This is left for future cleanup.
		/// </summary>
		public string WorkSheet
		{
			get { return _ws; }
		}

		/// <summary>
		/// Gets the workbook name for this range.
		/// </summary>
		public string Workbook
		{
			get { return _wb; }
		}

		internal string FullAddress
		{
			get
			{
				if (Addresses == null)
					return GetFullAddress(_ws, _address);
				string fullAddress = string.Empty;
				foreach (var a in Addresses)
				{
					fullAddress += GetFullAddress(_ws, a.Address) + ",";
				}
				return fullAddress.TrimEnd(',');
			}
		}

		/// <summary>
		/// returns the first address if the address is a multi address.
		/// A1:A2,B1:B2 returns A1:A2
		/// </summary>
		internal string FirstAddress
		{
			get
			{
				if (string.IsNullOrEmpty(_firstAddress))
				{
					return _address;
				}
				else
				{
					return _firstAddress;
				}
			}
		}

		internal string AddressSpaceSeparated
		{
			get
			{
				return _address.Replace(',', ' '); //Conditional formatting and a few other places use space as separator for mulit addresses.
			}
		}

		internal virtual List<ExcelAddress> Addresses
		{
			get
			{
				return _addresses;
			}
		}

		private ExcelTableAddress Table { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an <see cref="ExcelAddressBase"/>.
		/// DO NOT DELETE: used by subclasses.
		/// </summary>
		public ExcelAddressBase() { }

		/// <summary>
		/// Creates an Address object
		/// </summary>
		/// <param name="fromRow">start row</param>
		/// <param name="fromCol">start column</param>
		/// <param name="toRow">End row</param>
		/// <param name="toColumn">End column</param>
		public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
		{
			_fromRow = fromRow;
			_toRow = toRow;
			_fromCol = fromCol;
			_toCol = toColumn;
			Validate();

			_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
		}

		/// <summary>
		/// Creates an Address object
		/// </summary>
		/// <param name="fromRow">start row</param>
		/// <param name="fromCol">start column</param>
		/// <param name="toRow">End row</param>
		/// <param name="toColumn">End column</param>
		/// <param name="fromRowFixed">start row fixed</param>
		/// <param name="fromColFixed">start column fixed</param>
		/// <param name="toRowFixed">End row fixed</param>
		/// <param name="toColFixed">End column fixed</param>
		public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed)
		{
			_fromRow = fromRow;
			_toRow = toRow;
			_fromCol = fromCol;
			_toCol = toColumn;
			_fromRowFixed = fromRowFixed;
			_fromColFixed = fromColFixed;
			_toRowFixed = toRowFixed;
			_toColFixed = toColFixed;
			Validate();

			_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, fromColFixed, _toRowFixed, _toColFixed);
		}

		/// <summary>
		/// Creates an Address object
		/// </summary>
		/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
		/// <param name="address">The Excel Address</param>
		public ExcelAddressBase(string address)
		{
			SetAddress(address);
		}

		/// <summary>
		/// Creates an Address object
		/// </summary>
		/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
		/// <param name="address">The Excel Address</param>
		/// <param name="pck">Reference to the package to find information about tables and names</param>
		/// <param name="referenceAddress">The address</param>
		public ExcelAddressBase(string address, ExcelPackage pck, ExcelAddressBase referenceAddress)
		{
			SetAddress(address);
			SetRCFromTable(pck, referenceAddress);
		}

		/// <summary>
		/// Address is an defined name
		/// </summary>
		/// <param name="address">the name</param>
		/// <param name="isName">Should always be true</param>
		protected ExcelAddressBase(string address, bool isName)
		{
			if (isName)
			{
				_address = address;
				_fromRow = -1;
				_fromCol = -1;
				_toRow = -1;
				_toCol = -1;
				_start = null;
				_end = null;
			}
			else
			{
				SetAddress(address);
			}
		}
		#endregion

		#region Public methods
		/// <summary>
		/// Changes the worksheet this range is associated with if the original sheet matches <paramref name="wsName"/>.
		/// </summary>
		/// <param name="wsName">The original worksheet name.</param>
		/// <param name="newWs">The new worksheet name.</param>
		public void ChangeWorksheet(string wsName, string newWs)
		{
			if (_ws == wsName) _ws = newWs;
			if (Addresses == null)
				_address = this.GetAddress();
			else
			{
				_address = string.Empty;
				foreach (var a in Addresses)
				{
					if (a._ws == wsName)
					{
						a._ws = newWs;
						_address += a.GetAddress() + ",";
					}
					else
						_address += a._address + ",";
				}
				if (_address.Length > 0)
					_address = _address.TrimEnd(',');
			}
		}

		/// <summary>
		/// Creates a shifted range if this range is after <paramref name="row"/>.
		/// </summary>
		/// <param name="row">The row to shift after.</param>
		/// <param name="rows">The number of rows to shift.</param>
		/// <param name="setFixed">Indicates whether or not treat the reference as fixed.</param>
		/// <returns>A modified <see cref="ExcelAddressBase"/>.</returns>
		public ExcelAddressBase AddRow(int row, int rows, bool setFixed = false)
		{
			if (row > _toRow)
				return this;
			else if (row <= _fromRow)
				return new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow + rows), _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			else
				return new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}

		/// <summary>
		/// Creates a shifted range if this range is after <paramref name="row"/>.
		/// </summary>
		/// <param name="row">The row to shift after.</param>
		/// <param name="rows">The number of rows to shift.</param>
		/// <param name="setFixed">Indicates whether or not treat the reference as fixed.</param>
		/// <returns>A modified <see cref="ExcelAddressBase"/>.</returns>
		public ExcelAddressBase DeleteRow(int row, int rows, bool setFixed = false)
		{
			if (row > _toRow) //After
				return this;
			else if (row + rows <= _fromRow) //Before
				return new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow - rows), _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			else if (row <= _fromRow && row + rows > _toRow) //Inside
				return null;
			else  //Partly
			{
				if (row <= _fromRow)
					return new ExcelAddressBase(row, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
				else
					return new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows < row ? row - 1 : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			}
		}

		/// <summary>
		/// Creates a shifted range if this range is after <paramref name="col"/>.
		/// </summary>
		/// <param name="col">The column to shift after.</param>
		/// <param name="cols">The number of columns to shift.</param>
		/// <param name="setFixed">Indicates whether or not treat the reference as fixed.</param>
		/// <returns>A modified <see cref="ExcelAddressBase"/>.</returns>
		public ExcelAddressBase AddColumn(int col, int cols, bool setFixed = false)
		{
			if (col > _toCol)
				return this;
			else if (col <= _fromCol)
				return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol + cols), _toRow, (setFixed && _toColFixed ? _toCol : _toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			else
				return new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed ? _toCol : _toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}

		/// <summary>
		/// Creates a shifted range if this range is after <paramref name="col"/>.
		/// </summary>
		/// <param name="col">The column to shift after.</param>
		/// <param name="cols">The number of columns to shift.</param>
		/// <param name="setFixed">Indicates whether or not treat the reference as fixed.</param>
		/// <returns>A modified <see cref="ExcelAddressBase"/>.</returns>
		public ExcelAddressBase DeleteColumn(int col, int cols, bool setFixed = false)
		{
			if (col > _toCol) //After
				return this;
			else if (col + cols <= _fromCol) //Before
				return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol - cols), _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			else if (col <= _fromCol && col + cols > _toCol) //Inside
				return null;
			else  //Partly
			{
				if (col <= _fromCol)
					return new ExcelAddressBase(_fromRow, col, _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
				else
					return new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols < col ? col - 1 : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			}
		}

		/// <summary>
		/// Determines whether or not this range is valid.
		/// </summary>
		/// <returns>True if the start and end addresses are correct; otherwise false.</returns>
		public bool IsValidRowCol()
		{
			return !(_fromRow > _toRow ||
				   _fromCol > _toCol ||
				   _fromRow < 1 ||
				   _fromCol < 1 ||
				   _toRow > ExcelPackage.MaxRows ||
				   _toCol > ExcelPackage.MaxColumns);
		}

		/// <summary>
		/// Determines whether or not the given coordinates are within this <see cref="ExcelAddressBase"/>.
		/// </summary>
		/// <param name="row">The row to check.</param>
		/// <param name="column">The column to check.</param>
		/// <returns>True if the row and column do not map to a cell in this range; otherwise false.</returns>
		public bool ContainsCoordinate(int row, int column)
		{
			return row >= this._fromRow && row <= this._toRow && column >= this._fromCol && column <= this._toCol;
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Parses the given <paramref name="address"/> and sets the class members appropriately.
		/// </summary>
		/// <param name="address">The address to set.</param>
		protected internal void SetAddress(string address)
		{
			address = address.Trim();
			if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(address, "'"))
			{
				int pos = address.IndexOf("'", 1);
				while (pos < address.Length && address[pos + 1] == '\'')
				{
					pos = address.IndexOf("'", pos + 2);
				}
				var wbws = address.Substring(1, pos - 1).Replace("''", "'");
				SetWbWs(wbws);
				_address = address.Substring(pos + 2);
			}
			else if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(address, "[")) //Remove any external reference
			{
				SetWbWs(address);
			}
			else
			{
				_address = address;
			}
			if (_address.IndexOfAny(new char[] { ',', '!', '[' }) > -1)
			{
				//Advanced address. Including Sheet or multi or table.
				ExtractAddress(_address);
			}
			else
			{
				//Simple address
				GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
				_addresses = null;
				_start = null;
				_end = null;
			}
			_address = address;
			Validate();
		}

		/// <summary>
		/// Changes the address.
		/// </summary>
		protected internal virtual void ChangeAddress() { }

		/// <summary>
		/// Sets the class members based on the given table <paramref name="referenceAddress"/>.
		/// </summary>
		/// <param name="pck">The <see cref="ExcelPackage"/>.</param>
		/// <param name="referenceAddress">The table address.</param>
		internal void SetRCFromTable(ExcelPackage pck, ExcelAddressBase referenceAddress)
		{
			if (string.IsNullOrEmpty(_wb) && Table != null)
			{
				foreach (var ws in pck.Workbook.Worksheets)
				{
					foreach (var t in ws.Tables)
					{
						if (t.Name.Equals(Table.Name, StringComparison.InvariantCultureIgnoreCase))
						{
							_ws = ws.Name;
							if (Table.IsAll)
							{
								_fromRow = t.Address._fromRow;
								_toRow = t.Address._toRow;
							}
							else
							{
								if (Table.IsThisRow)
								{
									if (referenceAddress == null)
									{
										_fromRow = -1;
										_toRow = -1;
									}
									else
									{
										_fromRow = referenceAddress._fromRow;
										_toRow = _fromRow;
									}
								}
								else if (Table.IsHeader && Table.IsData)
								{
									_fromRow = t.Address._fromRow;
									_toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
								}
								else if (Table.IsData && Table.IsTotals)
								{
									_fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
									_toRow = t.Address._toRow;
								}
								else if (Table.IsHeader)
								{
									_fromRow = t.ShowHeader ? t.Address._fromRow : -1;
									_toRow = t.ShowHeader ? t.Address._fromRow : -1;
								}
								else if (Table.IsTotals)
								{
									_fromRow = t.ShowTotal ? t.Address._toRow : -1;
									_toRow = t.ShowTotal ? t.Address._toRow : -1;
								}
								else
								{
									_fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
									_toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
								}
							}

							if (string.IsNullOrEmpty(Table.ColumnSpan))
							{
								_fromCol = t.Address._fromCol;
								_toCol = t.Address._toCol;
								return;
							}
							else
							{
								var col = t.Address._fromCol;
								var cols = Table.ColumnSpan.Split(':');
								foreach (var c in t.Columns)
								{
									if (_fromCol <= 0 && cols[0].Equals(c.Name, StringComparison.InvariantCultureIgnoreCase))   //Issue15063 Add invariant igore case
									{
										_fromCol = col;
										if (cols.Length == 1)
										{
											_toCol = _fromCol;
											return;
										}
									}
									else if (cols.Length > 1 && _fromCol > 0 && cols[1].Equals(c.Name, StringComparison.InvariantCultureIgnoreCase)) //Issue15063 Add invariant igore case
									{
										_toCol = col;
										return;
									}

									col++;
								}
							}
						}
					}
				}
			}
		}

		/// <summary>
		/// Determines the type of collision, if any, between this range and the given <paramref name="address"/>.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddressBase"/> to compare against.</param>
		/// <param name="ignoreWs">Indicates whether or not to ignore whitespace.</param>
		/// <returns>An <see cref="eAddressCollition"/> indicating the collision type.</returns>
		internal eAddressCollition Collide(ExcelAddressBase address, bool ignoreWs = false)
		{
			if (ignoreWs == false && address.WorkSheet != WorkSheet && address.WorkSheet != null)
			{
				return eAddressCollition.No;
			}

			if (address._fromRow > _toRow || address._fromCol > _toCol
				||
				_fromRow > address._toRow || _fromCol > address._toCol)
			{
				return eAddressCollition.No;
			}
			else if (address._fromRow == _fromRow && address._fromCol == _fromCol &&
					address._toRow == _toRow && address._toCol == _toCol)
			{
				return eAddressCollition.Equal;
			}
			else if (address._fromRow >= _fromRow && address._toRow <= _toRow &&
					 address._fromCol >= _fromCol && address._toCol <= _toCol)
			{
				return eAddressCollition.Inside;
			}
			else
				return eAddressCollition.Partly;
		}
		#endregion

		#region Private Methods
		private void SetAddress(ref string first, ref string second, ref bool hasSheet, bool isMulti)
		{
			string ws, address;
			if (hasSheet)
			{
				ws = first;
				address = second;
				first = "";
				second = "";
			}
			else
			{
				address = first;
				ws = "";
				first = "";
			}
			hasSheet = false;
			if (string.IsNullOrEmpty(_firstAddress))
			{
				if (string.IsNullOrEmpty(_ws) || !string.IsNullOrEmpty(ws)) _ws = ws;
				_firstAddress = address;
				GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
			}
			if (isMulti)
			{
				if (_addresses == null) _addresses = new List<ExcelAddress>();
				_addresses.Add(new ExcelAddress(_ws, address));
			}
			else
			{
				_addresses = null;
			}
		}

		private string GetAddress()
		{
			var adr = "";
			if (!string.IsNullOrEmpty(_wb))
			{
				adr = "[" + _wb + "]";
			}

			if (!string.IsNullOrEmpty(_ws))
			{
				adr += string.Format("'{0}'!", _ws);
			}
			if (IsName)
				adr += GetAddress(_fromRow, _fromCol, _toRow, _toCol);
			else
				adr += GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			return adr;
		}

		private void SetWbWs(string address)
		{
			int pos;
			if (address[0] == '[')
			{
				pos = address.IndexOf("]");
				_wb = address.Substring(1, pos - 1);
				_ws = address.Substring(pos + 1);
			}
			else
			{
				_wb = "";
				_ws = address;
			}
			pos = _ws.IndexOf("!");
			if (pos > -1)
			{
				_address = _ws.Substring(pos + 1);
				_ws = _ws.Substring(0, pos);
			}
		}

		private void Validate()
		{
			if (_fromRow > _toRow || _fromCol > _toCol)
				throw new ArgumentOutOfRangeException("Start cell Address must be less or equal to End cell address");
		}

		private bool ExtractAddress(string fullAddress)
		{
			var brackPos = new Stack<int>();
			var bracketParts = new List<string>();
			string first = "", second = "";
			bool isText = false, hasSheet = false;
			try
			{
				if (fullAddress == "#REF!")
				{
					SetAddress(ref fullAddress, ref second, ref hasSheet, false);
					return true;
				}
				else if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "!"))
				{
					// invalid address!
					return false;
				}
				bool isMulti = false;
				for (int i = 0; i < fullAddress.Length; i++)
				{
					var c = fullAddress[i];
					if (c == '\'')
					{
						if (isText && i + 1 < fullAddress.Length && fullAddress[i] == '\'')
						{
							if (hasSheet)
							{
								second += c;
							}
							else
							{
								first += c;
							}
						}
						isText = !isText;
					}
					else
					{
						if (brackPos.Count > 0)
						{
							if (c == '[' && !isText)
							{
								brackPos.Push(i);
							}
							else if (c == ']' && !isText)
							{
								if (brackPos.Count > 0)
								{
									var from = brackPos.Pop();
									bracketParts.Add(fullAddress.Substring(from + 1, i - from - 1));

									if (brackPos.Count == 0)
									{
										HandleBrackets(first, second, bracketParts);
									}
								}
								else
								{
									//Invalid address!
									return false;
								}
							}
						}
						else if (c == '[' && !isText)
						{
							brackPos.Push(i);
						}
						else if (c == '!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
						{
							// the following is to handle addresses that specifies the
							// same worksheet twice: Sheet1!A1:Sheet1:A3
							// They will be converted to: Sheet1!A1:A3
							if (hasSheet && second != null && second.ToLower().EndsWith(first.ToLower()))
							{
								second = Regex.Replace(second, $"{first}$", string.Empty);
							}
							hasSheet = true;
						}
						else if (c == ',' && !isText)
						{
							isMulti = true;
							SetAddress(ref first, ref second, ref hasSheet, isMulti);
						}
						else
						{
							if (hasSheet)
							{
								second += c;
							}
							else
							{
								first += c;
							}
						}
					}
				}
				if (Table == null)
				{
					SetAddress(ref first, ref second, ref hasSheet, isMulti);
				}
				return true;
			}
			catch
			{
				return false;
			}
		}

		private void HandleBrackets(string first, string second, List<string> bracketParts)
		{
			if (!string.IsNullOrEmpty(first))
			{
				this.Table = new ExcelTableAddress();
				this.Table.Name = first;
				foreach (var s in bracketParts)
				{
					if (s.IndexOf("[") < 0)
					{
						switch (s.ToLower(CultureInfo.InvariantCulture))
						{
							case "#all":
								this.Table.IsAll = true;
								break;
							case "#headers":
								this.Table.IsHeader = true;
								break;
							case "#data":
								this.Table.IsData = true;
								break;
							case "#totals":
								this.Table.IsTotals = true;
								break;
							case "#this row":
								this.Table.IsThisRow = true;
								break;
							default:
								if (string.IsNullOrEmpty(this.Table.ColumnSpan))
									this.Table.ColumnSpan = s;
								else
									this.Table.ColumnSpan += ":" + s;
								break;
						}
					}
				}
			}
		}
		#endregion

		#region Static Methods
		/// <summary>
		/// Validates the given <see cref="Address"/>.
		/// </summary>
		/// <param name="Address">The address to validate.</param>
		/// <returns>An <see cref="AddressType"/> indicating the address status.</returns>
		internal static AddressType IsValid(string Address)
		{
			double d;
			if (Address == "#REF!")
			{
				return AddressType.Invalid;
			}
			else if (double.TryParse(Address, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) //A double, no valid address
			{
				return AddressType.Invalid;
			}
			else if (IsFormula(Address))
			{
				return AddressType.Formula;
			}
			else
			{
				string wb, ws, intAddress;
				if (SplitAddress(Address, out wb, out ws, out intAddress))
				{
					if (intAddress.Contains("[")) //Table reference
					{
						return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
					}
					else if (intAddress.Contains(","))
					{
						intAddress = intAddress.Substring(0, intAddress.IndexOf(','));
					}
					if (IsAddress(intAddress))
					{
						return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
					}
					else
					{
						return string.IsNullOrEmpty(wb) ? AddressType.InternalName : AddressType.ExternalName;
					}
				}
				else
				{
					return AddressType.Invalid;
				}

				//if(string.IsNullOrEmpty(wb));

			}
			//ExcelAddress a = new ExcelAddress(Address);
			//if (Address.IndexOf('!') > 0)
			//{                
			//    string[] split = Address.Split('!');
			//    if (split.Length == 2)
			//    {
			//        ws = split[0];
			//        Address = split[1];
			//    }
			//    else if (split.Length == 3 && split[1] == "#REF" && split[2] == "")
			//    {
			//        ws = split[0];
			//        Address = "#REF!";
			//        if (ws.StartsWith("[") && ws.IndexOf("]") > 1)
			//        {
			//            return AddressType.ExternalAddress;
			//        }
			//        else
			//        {
			//            return AddressType.InternalAddress;
			//        }
			//    }
			//    else
			//    {
			//        return AddressType.Invalid;
			//    }            
			//}
			//int _fromRow, column, _toRow, _toCol;
			//if (ExcelAddressBase.GetRowColFromAddress(Address, out _fromRow, out column, out _toRow, out _toCol))
			//{
			//    if (_fromRow > 0 && column > 0 && _toRow <= ExcelPackage.MaxRows && _toCol <= ExcelPackage.MaxColumns)
			//    {
			//        if (ws.StartsWith("[") && ws.IndexOf("]") > 1)
			//        {
			//            return AddressType.ExternalAddress;
			//        }
			//        else
			//        {
			//            return AddressType.InternalAddress;
			//        }
			//    }
			//    else
			//    {
			//        return AddressType.Invalid;
			//    }
			//}
			//else
			//{
			//    if(IsValidName(Address))
			//    {
			//        if (ws.StartsWith("[") && ws.IndexOf("]") > 1)
			//        {
			//            return AddressType.ExternalName;
			//        }
			//        else
			//        {
			//            return AddressType.InternalName;
			//        }
			//    }
			//    else
			//    {
			//        return AddressType.Invalid;
			//    }
			//}

		}

		/// <summary>
		/// Splits the given <paramref name="fullAddress"/> into its components.
		/// </summary>
		/// <param name="fullAddress">The address to split.</param>
		/// <param name="wb">The address' workbook.</param>
		/// <param name="ws">The address' worksheet.</param>
		/// <param name="address">The address' cell address.</param>
		/// <param name="defaultWorksheet">The default worksheet to use if none is parsed.</param>
		internal static void SplitAddress(string fullAddress, out string wb, out string ws, out string address, string defaultWorksheet = "")
		{
			wb = GetWorkbookPart(fullAddress);
			int ix = 0;
			ws = GetWorksheetPart(fullAddress, defaultWorksheet, ref ix);
			if (ix < fullAddress.Length)
			{
				if (fullAddress[ix] == '!')
				{
					address = fullAddress.Substring(ix + 1);
				}
				else
				{
					address = fullAddress.Substring(ix);
				}
			}
			else
			{
				address = "";
			}
		}

		private static bool IsAddress(string intAddress)
		{
			if (string.IsNullOrEmpty(intAddress)) return false;
			var cells = intAddress.Split(':');
			int fromRow, toRow, fromCol, toCol;

			if (!GetRowCol(cells[0], out fromRow, out fromCol, false))
			{
				return false;
			}
			if (cells.Length > 1)
			{
				if (!GetRowCol(cells[1], out toRow, out toCol, false))
				{
					return false;
				}
			}
			else
			{
				toRow = fromRow;
				toCol = fromCol;
			}
			if (fromRow <= toRow &&
				fromCol <= toCol &&
				fromCol > -1 &&
				toCol <= ExcelPackage.MaxColumns &&
				fromRow > -1 &&
				toRow <= ExcelPackage.MaxRows)
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		private static bool SplitAddress(string Address, out string wb, out string ws, out string intAddress)
		{
			wb = "";
			ws = "";
			intAddress = "";
			var text = "";
			bool isText = false;
			var brackPos = -1;
			for (int i = 0; i < Address.Length; i++)
			{
				if (Address[i] == '\'')
				{
					isText = !isText;
					if (i > 0 && Address[i - 1] == '\'')
					{
						text += "'";
					}
				}
				else
				{
					if (Address[i] == '!' && !isText)
					{
						if (text.Length > 0 && text[0] == '[')
						{
							wb = text.Substring(1, text.IndexOf("]") - 1);
							ws = text.Substring(text.IndexOf("]") + 1);
						}
						else
						{
							ws = text;
						}
						intAddress = Address.Substring(i + 1);
						return true;
					}
					else
					{
						if (Address[i] == '[' && !isText)
						{
							if (i > 0) //Table reference return full address;
							{
								intAddress = Address;
								return true;
							}
							brackPos = i;
						}
						else if (Address[i] == ']' && !isText)
						{
							if (brackPos > -1)
							{
								wb = text;
								text = "";
							}
							else
							{
								return false;
							}
						}
						else
						{
							text += Address[i];
						}
					}
				}
			}
			intAddress = text;
			return true;
		}

		private static bool IsFormula(string address)
		{
			var isText = false;
			for (int i = 0; i < address.Length; i++)
			{
				var addressChar = address[i];
				if (addressChar == '\'')
				{
					isText = !isText;
				}
				else
				{
					// Table references use [ ] around column names and since table column names can also contain unescaped formula characters,
					// we need to check that this is not a table column reference in order to avoid false positives.  Since function names and
					// formulas cannot contain [ ], we should be safe doing this check.
					if (addressChar == '[' || addressChar == ']')
						return false;
					if (isText == false && FormulaCharacters.Contains(addressChar))
					{
						return true;
					}
				}
			}
			return false;
		}

		private static string GetWorkbookPart(string address)
		{
			var ix = 0;
			if (address[0] == '[')
			{
				ix = address.IndexOf(']') + 1;
				if (ix > 0)
				{
					return address.Substring(1, ix - 2);
				}
			}
			return "";
		}

		private static string GetWorksheetPart(string address, string defaultWorkSheet, ref int endIx)
		{
			if (address == "") return defaultWorkSheet;
			var ix = 0;
			if (address[0] == '[')
			{
				ix = address.IndexOf(']') + 1;
			}
			else if (address[0] == '\'')
			{
				var escapedAddress = GetString(address, 0, out endIx);
				if (string.Empty == escapedAddress)
				{
					endIx = address.IndexOf('\'', 1) + 1;
					return address.Substring(1, endIx - 2);
				}
			}
			else if (address.IndexOf('!') > 0)
			{
				endIx = address.IndexOf('!');
				return address.Substring(0, endIx);
			}
			if (ix > 0 && ix < address.Length)
			{
				if (address[ix] == '\'')
				{
					return GetString(address, ix, out endIx);
				}
				else
				{
					var ixEnd = address.IndexOf('!', ix);
					if (ixEnd > ix)
					{
						return address.Substring(ix, ixEnd - ix);
					}
					else
					{
						return defaultWorkSheet;
					}
				}
			}
			else
			{
				return defaultWorkSheet;
			}
		}

		private static string GetString(string address, int ix, out int endIx)
		{
			var strIx = address.IndexOf("''");
			var prevStrIx = ix;
			while (strIx > -1)
			{
				prevStrIx = strIx;
				strIx = address.IndexOf("''");
			}
			endIx = address.IndexOf("'");
			return address.Substring(ix, endIx - ix).Replace("''", "'");
		}
		#endregion

		#region Object Overrides
		/// <summary>
		/// Returns the address text
		/// </summary>
		/// <returns>The address.</returns>
		public override string ToString()
		{
			return _address;
		}
		#endregion

		#region Nested Classes
		private class ExcelTableAddress
		{
			public string Name { get; set; }
			public string ColumnSpan { get; set; }
			public bool IsAll { get; set; }
			public bool IsHeader { get; set; }
			public bool IsData { get; set; }
			public bool IsTotals { get; set; }
			public bool IsThisRow { get; set; }
		}
		#endregion

		#region Enums
		internal enum eAddressCollition
		{
			No,
			Partly,
			Inside,
			Equal
		}

		internal enum AddressType
		{
			Invalid,
			InternalAddress,
			ExternalAddress,
			InternalName,
			ExternalName,
			Formula
		}
		#endregion
	}

	/// <summary>
	/// Range address with the address property readonly.
	/// NOTE:: This class is dumb and should be deleted whenever a motivated soul reads this.
	/// </summary>
	public class ExcelAddress : ExcelAddressBase
	{
		#region Properties
		/// <summary>
		/// The address for the range
		/// </summary>
		/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
		public new string Address
		{
			get
			{
				if (string.IsNullOrEmpty(_address) && _fromRow > 0)
				{
					_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
				}
				return _address;
			}
			set
			{
				this.Addresses?.Clear();
				SetAddress(value);
				ChangeAddress();
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of <see cref="ExcelAddress"/>.
		/// </summary>
		public ExcelAddress() { }

		/// <summary>
		/// Creates an instance of <see cref="ExcelAddress"/>.
		/// </summary>
		/// <param name="fromRow">start row</param>
		/// <param name="fromCol">start column</param>
		/// <param name="toRow">End row</param>
		/// <param name="toColumn">End column</param>
		public ExcelAddress(int fromRow, int fromCol, int toRow, int toColumn)
			: base(fromRow, fromCol, toRow, toColumn)
		{
			_ws = "";
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelAddress"/>.
		/// </summary>
		/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
		/// <param name="address">The Excel Address</param>
		public ExcelAddress(string address) : base(address) { }

		/// <summary>
		/// Creates an instance of <see cref="ExcelAddress"/>.
		/// </summary>
		/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
		/// <param name="address">The Excel Address</param>
		/// <param name="package">Reference to the package to find information about tables and names</param>
		/// <param name="referenceAddress">The address</param>
		public ExcelAddress(string address, ExcelPackage package, ExcelAddressBase referenceAddress) :
			base(address, package, referenceAddress)
		{ }

		/// <summary>
		/// Creates an instance of <see cref="ExcelAddress"/>.
		/// </summary>
		/// <param name="ws">A worksheet to assign the address.</param>
		/// <param name="address">The address to parse.</param>
		internal ExcelAddress(string ws, string address)
			: base(address)
		{
			if (string.IsNullOrEmpty(_ws)) _ws = ws;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelAddress"/>.
		/// </summary>
		/// <param name="ws">A worksheet name ot assign the addres..</param>
		/// <param name="address">The address to parse.</param>
		/// <param name="isName">Indicates whether or not this is a named range.</param>
		internal ExcelAddress(string ws, string address, bool isName)
			: base(address, isName)
		{
			if (string.IsNullOrEmpty(_ws)) _ws = ws;
		}
		#endregion
	}

	/// <summary>
	/// Represents a formula address.
	/// NOTE:: This class is dumb and should be deleted whenever a motivated soul reads this.
	/// </summary>
	public class ExcelFormulaAddress : ExcelAddressBase
	{
		#region Properties
		/// <summary>
		/// The address for the range
		/// </summary>
		/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
		public new string Address
		{
			get
			{
				if (string.IsNullOrEmpty(_address) && _fromRow > 0)
				{
					_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _toRowFixed, _fromColFixed, _toColFixed);
				}
				return _address;
			}
			set
			{
				base.Addresses?.Clear();
				SetAddress(value);
				ChangeAddress();
				SetFixed();
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelFormulaAddress"/>.
		/// </summary>
		/// <param name="address">The address for this instance.</param>
		public ExcelFormulaAddress(string address)
			: base(address)
		{
			SetFixed();
		}
		#endregion

		#region Internal Methods
		internal string GetOffset(int row, int column)
		{
			int fromRow = _fromRow, fromCol = _fromCol, toRow = _toRow, tocol = _toCol;
			var isMulti = (fromRow != toRow || fromCol != tocol);
			if (!_fromRowFixed)
			{
				fromRow += row;
			}
			if (!_fromColFixed)
			{
				fromCol += column;
			}
			if (isMulti)
			{
				if (!_toRowFixed)
				{
					toRow += row;
				}
				if (!_toColFixed)
				{
					tocol += column;
				}
			}
			else
			{
				toRow = fromRow;
				tocol = fromCol;
			}
			string a = GetAddress(fromRow, fromCol, toRow, tocol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			if (Addresses != null)
			{
				foreach (var sa in Addresses.Cast<ExcelFormulaAddress>())
				{
					a += "," + sa.GetOffset(row, column);
				}
			}
			return a;
		}
		#endregion

		#region Private Methods
		private void SetFixed()
		{
			if (Address.IndexOf("[") >= 0) return;
			var address = FirstAddress;
			if (_fromRow == _toRow && _fromCol == _toCol)
			{
				GetFixed(address, out _fromRowFixed, out _fromColFixed);
			}
			else
			{
				var cells = address.Split(':');
				GetFixed(cells[0], out _fromRowFixed, out _fromColFixed);
				GetFixed(cells[1], out _toRowFixed, out _toColFixed);
			}
		}

		private void GetFixed(string address, out bool rowFixed, out bool colFixed)
		{
			rowFixed = colFixed = false;
			var ix = address.IndexOf('$');
			while (ix > -1)
			{
				ix++;
				if (ix < address.Length)
				{
					if (address[ix] >= '0' && address[ix] <= '9')
					{
						rowFixed = true;
						break;
					}
					else
					{
						colFixed = true;
					}
				}
				ix = address.IndexOf('$', ix);
			}
		}
		#endregion
	}
}