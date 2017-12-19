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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml
{
	/// <summary>
	/// Collection for named ranges
	/// </summary>
	public class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>
	{
		#region Class Variables
		internal ExcelWorksheet _ws;
		internal ExcelWorkbook _wb;
		List<ExcelNamedRange> _list = new List<ExcelNamedRange>();
		Dictionary<string, int> _dic = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
		#endregion

		#region Constructors
		internal ExcelNamedRangeCollection(ExcelWorkbook wb)
		{
			_wb = wb;
			_ws = null;
		}

		internal ExcelNamedRangeCollection(ExcelWorkbook wb, ExcelWorksheet ws)
		{
			_wb = wb;
			_ws = ws;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add a new named range to the collection.
		/// </summary>
		/// <param name="Name">The name</param>
		/// <param name="range">The range</param>
		/// <returns></returns>
		public ExcelNamedRange Add(string Name, ExcelRangeBase range)
		{
			ExcelNamedRange item;
			if (range.IsName)
				item = new ExcelNamedRange(Name, _wb, _ws, _dic.Count);
			else
				item = new ExcelNamedRange(Name, _ws, range.Worksheet, range.Address, _dic.Count);
			item.NameFormula = range.FullAddress;
			AddName(Name, item);
			return item;
		}

		/// <summary>
		/// Adds a new named range to the collection.
		/// </summary>
		/// <param name="name">The name of the named range.</param>
		/// <param name="formula">The formula of the named range.</param>
		/// <param name="isHidden">A value indicating whether or not the hidden attribute was set on the named range (optional).</param>
		/// <param name="comments">The comments set on the named range (optional).</param>
		/// <returns>The newly-added named named range.</returns>
		public ExcelNamedRange Add(string name, string formula, bool isHidden = false, string comments = null)
		{
			var namedRange = new ExcelNamedRange(name, this._wb, this._ws, this._dic.Count);
			namedRange.IsNameHidden = isHidden;
			namedRange.NameComment = comments;
			namedRange.NameFormula = formula;
			this.AddName(name, namedRange);
			if (ExcelAddressBase.IsValidAddress(formula))
				namedRange.SetAddress(formula);
			return namedRange;
		}


		/// <summary>
		/// Remove a defined name from the collection
		/// </summary>
		/// <param name="Name">The name</param>
		public void Remove(string Name)
		{
			if (_dic.ContainsKey(Name))
			{
				var ix = _dic[Name];

				for (int i = ix + 1; i < _list.Count; i++)
				{
					_dic.Remove(_list[i].Name);
					_list[i].Index--;
					_dic.Add(_list[i].Name, _list[i].Index);
				}
				_dic.Remove(Name);
				_list.RemoveAt(ix);
			}
		}
		/// <summary>
		/// Checks collection for the presence of a key
		/// </summary>
		/// <param name="key">key to search for</param>
		/// <returns>true if the key is in the collection</returns>
		public bool ContainsKey(string key)
		{
			return _dic.ContainsKey(key);
		}
		/// <summary>
		/// The current number of items in the collection
		/// </summary>
		public int Count
		{
			get
			{
				return _dic.Count;
			}
		}
		/// <summary>
		/// Name indexer
		/// </summary>
		/// <param name="Name">The name (key) for a Named range</param>
		/// <returns>a reference to the range</returns>
		/// <remarks>
		/// Throws a KeyNotFoundException if the key is not in the collection.
		/// </remarks>
		public ExcelNamedRange this[string Name]
		{
			get
			{
				return _list[_dic[Name]];
			}
		}

		public ExcelNamedRange this[int Index]
		{
			get
			{
				return _list[Index];
			}
		}

		/// <summary>
		/// Add a defined name referencing value
		/// </summary>
		/// <param name="Name"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public ExcelNamedRange AddValue(string Name, object value)
		{
			var item = new ExcelNamedRange(Name, _wb, _ws, _dic.Count);
			item.NameValue = value;
			AddName(Name, item);
			return item;
		}

		/// <summary>
		/// Add a defined name referencing a formula
		/// </summary>
		/// <param name="Name"></param>
		/// <param name="Formula"></param>
		/// <returns></returns>
		public ExcelNamedRange AddFormula(string Name, string Formula)
		{
			var item = new ExcelNamedRange(Name, _wb, _ws, _dic.Count);
			item.NameFormula = Formula;
			AddName(Name, item);
			return item;
		}
		#endregion

		#region Internal Methods
		internal void Insert(int rowFrom, int colFrom, int rows, int cols)
		{
			Insert(rowFrom, colFrom, rows, cols, n => true);
		}

		internal void Insert(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
		{
			var namedRanges = this._list.Where(filter);
			foreach (var namedRange in namedRanges)
			{
				InsertRows(rowFrom, rows, namedRange);
				InsertColumns(colFrom, cols, namedRange);
				// TODO: Need to pass in worksheet that insert is occurring on, should only update formulas referencing this worksheet.
				this.UpdateNameFormulaInsert(rowFrom, colFrom, rows, cols, namedRange);
			}
		}

		internal void Delete(int rowFrom, int colFrom, int rows, int cols)
		{
			Delete(rowFrom, colFrom, rows, cols, n => true);
		}

		internal void Delete(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
		{
			// TODO: Update references in NameFormula.
			var namedRanges = this._list.Where(filter);
			foreach (var namedRange in namedRanges)
			{
				ExcelAddressBase adr;
				if (rows == 0)
				{
					adr = namedRange.DeleteColumn(colFrom, cols);
				}
				else
				{
					adr = namedRange.DeleteRow(rowFrom, rows);
				}
				if (adr == null)
				{
					namedRange.Address = "#REF!";
				}
				else
				{
					namedRange.Address = adr.Address;
				}
			}
		}

		internal void Clear()
		{
			while (this.Count > 0)
			{
				Remove(_list[0].Name);
			}
		}
		#endregion

		#region Private Methods
		private void AddName(string Name, ExcelNamedRange item)
		{
			_dic.Add(Name, _list.Count);
			_list.Add(item);
		}

		private void InsertColumns(int colFrom, int cols, ExcelNamedRange namedRange)
		{
			if (colFrom > 0)
			{
				if (namedRange.Addresses?.Any() == true)
				{
					string worksheetPrefix = namedRange.FullAddress.Contains("!") ? this.ExtractWorksheetPrefix(namedRange.FullAddress) : string.Empty;
					var addressBuilder = new StringBuilder();
					foreach (var address in namedRange.Addresses)
					{
						//try
						//{
						//	if (address._fromColFixed && (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns) == false)
						//	{
						//		if (colFrom <= address.Start.Column)
						//			addressBuilder.Append($"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column + cols, address.End.Row, address.End.Column + cols, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)},");
						//		else if (colFrom <= address.End.Column)
						//			addressBuilder.Append($"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column + cols, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)},");
						//		else
						//			addressBuilder.Append($"{worksheetPrefix}{address.Address},");
						//	}
						//	else
						//		addressBuilder.Append($"{worksheetPrefix}{address.Address},");
						//}
						//catch (ArgumentOutOfRangeException)
						//{
						//	addressBuilder.Append($"{worksheetPrefix}{address.Address},");
						//}
						addressBuilder.Append(this.UpdateAddressInsertColumns(colFrom, cols, address, worksheetPrefix));
						addressBuilder.Append(",");
					}
					addressBuilder.Length--;
					namedRange.Address = addressBuilder.ToString();
				}
				else
				{
					try
					{
						if (namedRange._fromColFixed && (namedRange.Start.Column == 1 && namedRange.End.Column == ExcelPackage.MaxColumns) == false)
						{
							if (colFrom <= namedRange.Start.Column)
							{
								var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column + cols, namedRange.End.Row, namedRange.End.Column + cols, namedRange._fromRowFixed, namedRange._fromColFixed, namedRange._toRowFixed, namedRange._toColFixed);
								namedRange.Address = BuildNewAddress(namedRange, newAddress);
							}
							else if (colFrom <= namedRange.End.Column)
							{
								var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column, namedRange.End.Row, namedRange.End.Column + cols, namedRange._fromRowFixed, namedRange._fromColFixed, namedRange._toRowFixed, namedRange._toColFixed);
								namedRange.Address = BuildNewAddress(namedRange, newAddress);
							}
						}
					}
					catch (ArgumentOutOfRangeException) { /* This means the named range has an invalid address in it, so just ignore the problem. */ }
				}
			}
		}

		private void UpdateNameFormulaInsert(int rowFrom, int colFrom, int rows, int columns, ExcelNamedRange namedRange)
		{
			var tokens = namedRange.myWorkbook.FormulaParser.Lexer.Tokenize(namedRange.NameFormula);
			var formulaBuilder = new StringBuilder();
			foreach (var token in tokens)
			{
				if (token.TokenType == TokenType.ExcelAddress)
				{
					var excelAddress = new ExcelAddress(token.Value);
					// TODO: Do correct equivalency check here.
					// Will this check actually matter?
					if (excelAddress.WorkSheet == this._ws.Name)
					{
						string updatedAddress = this.UpdateAddressInsertRows(rowFrom, rows, excelAddress);
						updatedAddress = this.UpdateAddressInsertColumns(colFrom, columns, new ExcelAddress(updatedAddress));
						formulaBuilder.Append(updatedAddress);
					}
					else
						formulaBuilder.Append(token.Value);
				}
				else
					formulaBuilder.Append(token.Value.ToString());
			}
			namedRange.NameFormula = formulaBuilder.ToString();
		}

		private string UpdateAddressInsertColumns(int colFrom, int columns, ExcelAddress address, string worksheetPrefix = null)
		{
			if (columns < 1)
				return address.Address;
			string updatedAddress = string.Empty;
			try
			{
				if (address._fromColFixed && (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns) == false)
				{
					if (colFrom <= address.Start.Column)
						updatedAddress = $"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column + columns, address.End.Row, address.End.Column + columns, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)}";
					else if (colFrom <= address.End.Column)
						updatedAddress = $"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column + columns, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)}";
					else
						updatedAddress = $"{worksheetPrefix}{address.Address}";
				}
				else
					updatedAddress = $"{worksheetPrefix}{address.Address}";
			}
			catch (ArgumentOutOfRangeException)
			{
				updatedAddress = $"{worksheetPrefix}{address.Address}";
			}
			return updatedAddress;
		}

		private string ExtractWorksheetPrefix(string address)
		{
			var worksheetPart = address.Split('!')[0];
			worksheetPart = worksheetPart.Trim('\'');
			return $"'{worksheetPart}'!";
		}

		private static string BuildNewAddress(ExcelNamedRange namedRange, string newAddress)
		{
			if (namedRange.FullAddress.Contains("!"))
			{
				var worksheet = namedRange.FullAddress.Split('!')[0];
				worksheet = worksheet.Trim('\'');
				newAddress = ExcelCellBase.GetFullAddress(worksheet, newAddress);
			}
			return newAddress;
		}

		private void InsertRows(int rowFrom, int rows, ExcelNamedRange namedRange)
		{
			if (rows < 1)
				return;
			if (namedRange.Addresses?.Any() == true)
			{
				string worksheetPrefix = namedRange.FullAddress.Contains("!") ? this.ExtractWorksheetPrefix(namedRange.FullAddress) : string.Empty;
				var addressBuilder = new StringBuilder();
				foreach (var address in namedRange.Addresses)
				{
					//try
					//{
					//	if (address._fromRowFixed)
					//	{
					//		if (rowFrom <= address.Start.Row)
					//			addressBuilder.Append($"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row + rows, address.Start.Column, address.End.Row + rows, address.End.Column, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)},");
					//		else if (rowFrom <= address.End.Row)
					//			addressBuilder.Append($"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column, address.End.Row + rows, address.End.Column, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)},");
					//		else
					//			addressBuilder.Append($"{worksheetPrefix}{address.Address},");
					//	}
					//	else
					//		addressBuilder.Append($"{worksheetPrefix}{address.Address},");
					//}
					//catch (ArgumentOutOfRangeException)
					//{
					//	addressBuilder.Append($"{worksheetPrefix}{address.Address},");
					//}
					addressBuilder.Append(this.UpdateAddressInsertRows(rowFrom, rows, address, worksheetPrefix));
					addressBuilder.Append(",");
				}
				addressBuilder.Length--;
				namedRange.Address = addressBuilder.ToString();
			}
			else
			{
				try
				{
					if (namedRange._fromRowFixed)
					{
						if (rowFrom <= namedRange.Start.Row)
						{
							var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row + rows, namedRange.Start.Column, namedRange.End.Row + rows, namedRange.End.Column, namedRange._fromRowFixed, namedRange._fromColFixed, namedRange._toRowFixed, namedRange._toColFixed);
							namedRange.Address = BuildNewAddress(namedRange, newAddress);
						}
						else if (rowFrom <= namedRange.End.Row)
						{
							var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column, namedRange.End.Row + rows, namedRange.End.Column, namedRange._fromRowFixed, namedRange._fromColFixed, namedRange._toRowFixed, namedRange._toColFixed);
							namedRange.Address = BuildNewAddress(namedRange, newAddress);
						}
					}
				}
				catch (ArgumentOutOfRangeException) { /* This means the named range has an invalid address in it, so just ignore the problem. */ }
			}
		}

		private string UpdateAddressInsertRows(int rowFrom, int rows, ExcelAddress address, string worksheetPrefix = null)
		{
			if (rows < 1)
				return address.Address;
			string updatedAddress = string.Empty;
			try
			{
				if (address._fromRowFixed)
				{
					if (rowFrom <= address.Start.Row)
						updatedAddress = $"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row + rows, address.Start.Column, address.End.Row + rows, address.End.Column, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)}";
					else if (rowFrom <= address.End.Row)
						updatedAddress = $"{worksheetPrefix}{ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column, address.End.Row + rows, address.End.Column, address._fromRowFixed, address._fromColFixed, address._toRowFixed, address._toColFixed)}";
					else
						updatedAddress = $"{worksheetPrefix}{address.Address}";
				}
				else
					updatedAddress = $"{worksheetPrefix}{address.Address}";
			}
			catch (ArgumentOutOfRangeException)
			{
				updatedAddress = $"{worksheetPrefix}{address.Address}";
			}
			return updatedAddress;
		}
		#endregion

		#region "IEnumerable"
		#region IEnumerable<ExcelNamedRange> Members
		/// <summary>
		/// Implement interface method IEnumerator&lt;ExcelNamedRange&gt; GetEnumerator()
		/// </summary>
		/// <returns></returns>
		public IEnumerator<ExcelNamedRange> GetEnumerator()
		{
			return _list.GetEnumerator();
		}
		#endregion
		#region IEnumerable Members
		/// <summary>
		/// Implement interface method IEnumeratable GetEnumerator()
		/// </summary>
		/// <returns></returns>
		IEnumerator IEnumerable.GetEnumerator()
		{
			return _list.GetEnumerator();
		}
		#endregion
		#endregion
	}
}
