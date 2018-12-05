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
 * Jan Källman		    Initial Release		        2011-11-02
 * Jan Källman          Total rewrite               2010-03-01
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.X14DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Slicers;
using OfficeOpenXml.Drawing.Sparkline;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using static OfficeOpenXml.ExcelErrorValue;

namespace OfficeOpenXml
{
	#region Enumerators
	/// <summary>
	/// Worksheet hidden enumeration
	/// </summary>
	public enum eWorkSheetHidden
	{
		/// <summary>
		/// The worksheet is visible
		/// </summary>
		Visible,
		/// <summary>
		/// The worksheet is hidden but can be shown by the user via the user interface
		/// </summary>
		Hidden,
		/// <summary>
		/// The worksheet is hidden and cannot be shown by the user via the user interface
		/// </summary>
		VeryHidden
	}

	[Flags]
	internal enum CellFlags
	{
		//Merged = 0x1,
		RichText = 0x2,
		SharedFormula = 0x4,
		ArrayFormula = 0x8
	}
	#endregion

	#region Data Structures
	/// <summary>
	/// For Cell value structure (for memory optimization of huge sheet)
	/// </summary>
	public struct ExcelCoreValue
	{
		internal object _value;
		internal int _styleId;
	}
	#endregion

	/// <summary>
	/// Represents an Excel worksheet and provides access to its properties and methods
	/// </summary>
	public class ExcelWorksheet : XmlHelper, IEqualityComparer<ExcelWorksheet>, IDisposable
	{
		#region Internal Nested Classes
		/// <summary>
		/// Represents the Formulas component of a worksheet.
		/// </summary>
		internal class Formulas
		{
			#region Properties
			private ISourceCodeTokenizer Tokenizer { get; set; }
			internal int Index { get; set; }
			internal string Address { get; set; }
			internal bool IsArray { get; set; }
			public string Formula { get; set; }
			public int StartRow { get; set; }
			public int StartCol { get; set; }

			private IEnumerable<Token> Tokens { get; set; }
			#endregion

			#region Constructors
			/// <summary>
			/// Initialize a new <see cref="Formulas"/> object.
			/// </summary>
			/// <param name="tokenizer">The tokenizer to initialize the Formulas object from.</param>
			public Formulas(ISourceCodeTokenizer tokenizer)
			{
				this.Tokenizer = tokenizer;
			}
			#endregion

			#region Public Methods
			/// <summary>
			/// Clone this object.
			/// </summary>
			/// <returns>A deep copy of this <see cref="Formulas"/> object.</returns>
			public Formulas Clone()
			{
				var formulas = new Formulas(this.Tokenizer);
				formulas.Index = this.Index;
				formulas.Address = (string)this.Address.Clone();
				formulas.IsArray = this.IsArray;
				formulas.Formula = (string)this.Formula.Clone();
				formulas.StartRow = this.StartRow;
				formulas.StartCol = this.StartCol;
				return formulas;
			}
			#endregion

			#region Internal Methods
			/// <summary>
			/// Gets a formula from the specified worksheet position.
			/// </summary>
			/// <param name="row">The row of the cell to read.</param>
			/// <param name="column">The column of the sheet to read.</param>
			/// <param name="worksheet">The sheet to read from.</param>
			/// <returns>The formula at the specified position.</returns>
			internal string GetFormula(int row, int column, string worksheet)
			{
				if (this.StartRow == row && this.StartCol == column)
				{
					return this.Formula;
				}

				if (this.Tokens == null)
				{
					this.Tokens = this.Tokenizer.Tokenize(this.Formula, worksheet);
				}

				string f = "";
				foreach (var token in Tokens)
				{
					if (token.TokenType == TokenType.ExcelAddress)
					{
						var a = new ExcelFormulaAddress(token.Value);
						f += !string.IsNullOrEmpty(a._wb) || !string.IsNullOrEmpty(a._ws)
							? token.Value
							: a.GetOffset(row - StartRow, column - StartCol);
					}
					else
					{
						f += token.Value;
					}
				}
				return f;
			}
			#endregion
		}

		/// <summary>
		/// Collection containing merged cell addresses
		/// </summary>
		public class MergeCellsCollection : IEnumerable<string>
		{
			#region Properties
			/// <summary>
			/// Gets the number of cells in this collection.
			/// </summary>
			public int Count
			{
				get
				{
					return this.List.Count;
				}
			}
			internal ICellStore<int> Cells { get; set; } = CellStore.Build<int>();
			internal List<string> List { get; } = new List<string>();
			#endregion

			#region Constructors
			internal MergeCellsCollection() { }
			#endregion

			#region Public Operators
			public string this[int row, int column]
			{
				get
				{
					int ix = -1;
					if (this.Cells.Exists(row, column, out ix) && ix >= 0 && ix < List.Count)  //Fixes issue 15075
					{
						return List[ix];
					}
					else
					{
						return null;
					}
				}
			}
			public string this[int index]
			{
				get
				{
					return List[index];
				}
			}
			#endregion

			#region Internal Methods
			internal void Add(ExcelAddress address, bool doValidate)
			{
				int ix = 0;

				//Validate
				if (doValidate && Validate(address) == false)
				{
					throw (new ArgumentException("Can't merge and already merged range"));
				}
				lock (this)
				{
					ix = this.List.Count;
					this.List.Add(address.Address);
					this.SetIndex(address, ix);
				}
			}

			internal void SetIndex(ExcelAddress address, int ix)
			{
				if (address._fromRow == 1 && address._toRow == ExcelPackage.MaxRows) //Entire row
				{
					for (int col = address._fromCol; col <= address._toCol; col++)
					{
						this.Cells.SetValue(0, col, ix);
					}
				}
				else if (address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns) //Entire row
				{
					for (int row = address._fromRow; row <= address._toRow; row++)
					{
						this.Cells.SetValue(row, 0, ix);
					}
				}
				else
				{
					for (int col = address._fromCol; col <= address._toCol; col++)
					{
						for (int row = address._fromRow; row <= address._toRow; row++)
						{
							this.Cells.SetValue(row, col, ix);
						}
					}
				}
			}
			internal void Remove(string Item)
			{
				this.List.Remove(Item);
			}

			internal void Clear(ExcelAddress Destination)
			{
				var cse = this.Cells.GetEnumerator(Destination._fromRow, Destination._fromCol, Destination._toRow, Destination._toCol);
				var used = new HashSet<int>();
				while (cse.MoveNext())
				{
					var v = cse.Value;
					if (!used.Contains(v) && this.List[v] != null)
					{
						var adr = new ExcelAddress(this.List[v]);
						if (!(Destination.Collide(adr) == ExcelAddress.eAddressCollition.Inside || Destination.Collide(adr) == ExcelAddress.eAddressCollition.Equal))
						{
							throw (new InvalidOperationException(string.Format("Can't delete/overwrite merged cells. A range is partly merged with the another merged range. {0}", adr._address)));
						}
						used.Add(v);
					}
				}

				this.Cells.Clear(Destination._fromRow, Destination._fromCol, Destination._toRow - Destination._fromRow + 1, Destination._toCol - Destination._fromCol + 1);
				foreach (var i in used)
				{
					this.List[i] = null;
				}
			}
			#endregion

			#region Private Methods
			private bool Validate(ExcelAddress address)
			{
				int ix = 0;
				if (this.Cells.Exists(address._fromRow, address._fromCol, out ix))
				{
					if (ix >= 0 && ix < this.List.Count && this.List[ix] != null && address.Address == this.List[ix])
					{
						return true;
					}
					else
					{
						return false;
					}
				}

				var cse = this.Cells.GetEnumerator(address._fromRow, address._fromCol, address._toRow, address._toCol);
				//cells
				while (cse.MoveNext())
				{
					return false;
				}
				//Entire column
				cse = this.Cells.GetEnumerator(0, address._fromCol, 0, address._toCol);
				while (cse.MoveNext())
				{
					return false;
				}
				//Entire row
				cse = this.Cells.GetEnumerator(address._fromRow, 0, address._toRow, 0);
				while (cse.MoveNext())
				{
					return false;
				}
				return true;
			}
			#endregion

			#region IEnumerable<string> Members

			public IEnumerator<string> GetEnumerator()
			{
				return this.List.GetEnumerator();
			}

			#endregion

			#region IEnumerable Members

			System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
			{
				return this.List.GetEnumerator();
			}

			#endregion
		}
		#endregion

		#region Constants
		const string outLineSummaryBelowPath = "d:sheetPr/d:outlinePr/@summaryBelow";
		const string outLineSummaryRightPath = "d:sheetPr/d:outlinePr/@summaryRight";
		const string outLineApplyStylePath = "d:sheetPr/d:outlinePr/@applyStyles";
		const string tabColorPath = "d:sheetPr/d:tabColor/@rgb";
		const string codeModuleNamePath = "d:sheetPr/@codeName";
		const int BLOCKSIZE = 8192;
		#endregion

		#region Class Variables
		internal ICellStore<ExcelCoreValue> _values;
		internal ICellStore<object> _formulas;
		internal IFlagStore _flags;
		internal ICellStore<List<Token>> _formulaTokens;
		internal ICellStore<Uri> _hyperLinks;
		internal ICellStore<int> _commentsStore;
		internal Dictionary<int, Formulas> _sharedFormulas = new Dictionary<int, Formulas>();
		private ExcelSlicers _slicers;
		private Uri _worksheetUri;
		private string _name;
		private int _sheetID;
		private int _positionID;
		private string _relationshipID;
		private XmlDocument _worksheetXml;
		private ExcelWorksheetView SheetView;
		private ExcelNamedRangeCollection _names;
		private ExcelSparklineGroups _sparklineGroups;
		private double _defaultRowHeight = double.NaN;
		private ExcelCommentCollection _comments = null;
		private ExcelHeaderFooter _headerFooter;
		private MergeCellsCollection _mergedCells = new MergeCellsCollection();
		private Dictionary<int, int> columnStyles = null;
		private ExcelSheetProtection _protection = null;
		private ExcelProtectedRangeCollection _protectedRanges;
		private ExcelDrawings _drawings = null;
		private ExcelTableCollection _tables = null;
		private ExcelPivotTableCollection _pivotTables = null;
		private ExcelConditionalFormattingCollection _conditionalFormatting = null;
		private X14ConditionalFormattingCollection _x14ConditionalFormatting = null;
		private ExcelDataValidationCollection _dataValidation = null;
		private ExcelX14DataValidationCollection _x14dataValidation = null;
		private ExcelBackgroundImage _backgroundImage = null;
		private bool? _customHeight;
		#endregion

		#region Internal Properties
		/// <summary>
		///  The ExcelPackage this worksheet exists in.
		/// </summary>
		internal ExcelPackage Package { get; set; }

		/// <summary>
		/// The Uri to the worksheet within the package
		/// </summary>
		internal Uri WorksheetUri { get { return (this._worksheetUri); } }

		/// <summary>
		/// The Zip.ZipPackagePart for the worksheet within the package
		/// </summary>
		internal Packaging.ZipPackagePart Part { get { return (this.Package.Package.GetPart(WorksheetUri)); } }

		/// <summary>
		/// The ID for the worksheet's relationship with the workbook in the package
		/// </summary>
		internal string RelationshipID { get { return (this._relationshipID); } }
		#endregion

		#region Public Properties
		/// <summary>
		/// The unique identifier for the worksheet.
		/// </summary>
		public int SheetID { get { return (this._sheetID); } }

		/// <summary>
		/// The position of the worksheet.
		/// </summary>
		public int PositionID { get { return (this._positionID); } internal set { this._positionID = value; } }

		/// <summary>
		/// Returns a ExcelWorksheetView object that allows you to set the view state properties of the worksheet
		/// </summary>
		public ExcelWorksheetView View
		{
			get
			{
				if (this.SheetView == null)
				{
					XmlNode node = this.TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
					if (node == null)
					{
						this.CreateNode("d:sheetViews/d:sheetView");     //this one shouls always exist. but check anyway
						node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
					}
					this.SheetView = new ExcelWorksheetView(this.NameSpaceManager, node, this);
				}
				return (this.SheetView);
			}
		}

		/// <summary>
		/// The worksheet's display name as it appears on the tab
		/// </summary>
		public string Name
		{
			get { return (this._name); }
			set
			{
				if (value == this._name) return;
				value = this.Package.Workbook.Worksheets.ValidateAndFixSheetName(value);
				foreach (var ws in this.Workbook.Worksheets)
				{
					if (ws.PositionID != this.PositionID && ws.Name.Equals(value, StringComparison.InvariantCultureIgnoreCase))
					{
						throw (new ArgumentException("Worksheet name must be unique"));
					}
				}
				this.Package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@name", _sheetID), value);
				this.ChangeNames(value);

				this._name = value;
			}
		}

		/// <summary>
		/// The slicers that exist on this worksheet.
		/// </summary>
		public ExcelSlicers Slicers
		{
			get
			{
				if (this._slicers == null)
				{
					this._slicers = new ExcelSlicers(this);
				}
				return this._slicers;
			}
		}

		/// <summary>
		/// The index in the worksheets collection
		/// </summary>
		public int Index { get { return (this._positionID); } }

		/// <summary>
		/// Address for autofilter
		/// <seealso cref="ExcelRangeBase.AutoFilter" />
		/// </summary>
		public ExcelAddress AutoFilterAddress
		{
			get
			{
				this.CheckSheetType();
				string address = this.GetXmlNodeString("d:autoFilter/@ref");
				if (address == "")
				{
					return null;
				}
				else
				{
					return new ExcelAddress(address);
				}
			}
			internal set
			{
				this.CheckSheetType();
				if (value == null)
				{
					this.DeleteAllNode("d:autoFilter/@ref");
				}
				else
				{
					this.SetXmlNodeString("d:autoFilter/@ref", value.Address);
				}
			}
		}

		/// <summary>
		/// Gets the number of distinct autoFilter ranges present on the worksheet.
		/// </summary>
		public bool HasAutoFilters
		{
			get
			{
				var autoFilters = this.WorksheetXml.SelectNodes("//d:autoFilter", this.NameSpaceManager);
				var autoFilteredRanges = this.Names.ContainsKey("_xlnm._FilterDatabase") ? 1 : 0;
				return (autoFilters.Count + autoFilteredRanges) > 0;
			}
		}

		/// <summary>
		/// Provides access to named ranges
		/// </summary>
		public ExcelNamedRangeCollection Names
		{
			get
			{
				CheckSheetType();
				return _names;
			}
		}

		/// <summary>
		/// Gets the <see cref="ExcelSparklineGroups"/> that exist on the worksheet.
		/// </summary>
		public ExcelSparklineGroups SparklineGroups
		{
			get
			{
				CheckSheetType();
				if (_sparklineGroups == null)
				{
					// Add required namespaces for Sparkline support.
					if (!NameSpaceManager.HasNamespace("x14"))
						NameSpaceManager.AddNamespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
					if (!NameSpaceManager.HasNamespace("xm"))
						NameSpaceManager.AddNamespace("xm", "http://schemas.microsoft.com/office/excel/2006/main");

					var sparklineGroupsNode = TopNode.SelectSingleNode("d:extLst/d:ext/x14:sparklineGroups", NameSpaceManager);
					if (sparklineGroupsNode != null)
						_sparklineGroups = new ExcelSparklineGroups(this, NameSpaceManager, sparklineGroupsNode);
					else
						_sparklineGroups = new ExcelSparklineGroups(this, NameSpaceManager);
				}

				return _sparklineGroups;
			}
		}

		/// <summary>
		/// Indicates if the worksheet is hidden in the workbook
		/// </summary>
		public eWorkSheetHidden Hidden
		{
			get
			{
				string state = this.Package.Workbook.GetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
				if (state == "hidden")
				{
					return eWorkSheetHidden.Hidden;
				}
				else if (state == "veryHidden")
				{
					return eWorkSheetHidden.VeryHidden;
				}
				return eWorkSheetHidden.Visible;
			}
			set
			{
				if (value == eWorkSheetHidden.Visible)
				{
					this.Package.Workbook.DeleteNode(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
				}
				else
				{
					string v;
					v = value.ToString();
					v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1);
					this.Package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID), v);
				}
			}
		}

		/// <summary>
		/// Get/set the default height of all rows in the worksheet
		/// </summary>
		public double DefaultRowHeight
		{
			get
			{
				this.CheckSheetType();
				if (double.IsNaN(this._defaultRowHeight) == false)
					return this._defaultRowHeight;

				this._defaultRowHeight = GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight");
				if (double.IsNaN(this._defaultRowHeight) || this.CustomHeight == false)
				{
					this._defaultRowHeight = this.GetRowHeightFromNormalStyle();
				}
				return this._defaultRowHeight;
			}
			set
			{
				this.CheckSheetType();
				this._defaultRowHeight = value;
				if (double.IsNaN(value))
				{
					this.DeleteNode("d:sheetFormatPr/@defaultRowHeight");
				}
				else
				{
					this.SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", value.ToString(CultureInfo.InvariantCulture));
					//Check if this is the default width for the normal style
					double defHeight = this.GetRowHeightFromNormalStyle();
					this.CustomHeight = true;
				}
			}
		}

		/// <summary>
		/// 'True' if defaultRowHeight value has been manually set, or is different from the default value.
		/// Is automaticlly set to 'True' when assigning the DefaultRowHeight property
		/// </summary>
		public bool CustomHeight
		{
			get
			{
				if (this._customHeight == null)
					this._customHeight = this.GetXmlNodeBool("d:sheetFormatPr/@customHeight");
				return this._customHeight.Value;
			}
			set
			{
				this._customHeight = value;
				this.SetXmlNodeBool("d:sheetFormatPr/@customHeight", value);
			}
		}

		/// <summary>
		/// Get/set the default width of all columns in the worksheet
		/// </summary>
		public double DefaultColWidth
		{
			get
			{
				this.CheckSheetType();
				double ret = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
				if (double.IsNaN(ret))
				{
					var mfw = Convert.ToDouble(Workbook.MaxFontWidth);
					var widthPx = mfw * 7;
					var margin = Math.Truncate(mfw / 4 + 0.999) * 2 + 1;
					if (margin < 5) margin = 5;
					while (Math.Truncate((widthPx - margin) / mfw * 100 + 0.5) / 100 < 8)
					{
						widthPx++;
					}
					widthPx = widthPx % 8 == 0 ? widthPx : 8 - widthPx % 8 + widthPx;
					var width = Math.Truncate((widthPx - margin) / mfw * 100 + 0.5) / 100;
					return Math.Truncate((width * mfw + margin) / mfw * 256) / 256;
				}
				return ret;
			}
			set
			{
				this.CheckSheetType();
				this.SetXmlNodeString("d:sheetFormatPr/@defaultColWidth", value.ToString(CultureInfo.InvariantCulture));

				if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight")))
				{
					this.SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", GetRowHeightFromNormalStyle().ToString(CultureInfo.InvariantCulture));
				}
			}
		}

		/// <summary>
		/// Summary rows below details
		/// </summary>
		public bool OutLineSummaryBelow
		{
			get
			{
				this.CheckSheetType();
				return GetXmlNodeBool(outLineSummaryBelowPath);
			}
			set
			{
				this.CheckSheetType();
				this.SetXmlNodeString(outLineSummaryBelowPath, value ? "1" : "0");
			}
		}

		/// <summary>
		/// Summary rows to right of details
		/// </summary>
		public bool OutLineSummaryRight
		{
			get
			{
				this.CheckSheetType();
				return this.GetXmlNodeBool(outLineSummaryRightPath);
			}
			set
			{
				this.CheckSheetType();
				this.SetXmlNodeString(outLineSummaryRightPath, value ? "1" : "0");
			}
		}

		/// <summary>
		/// Automatic styles
		/// </summary>
		public bool OutLineApplyStyle
		{
			get
			{
				this.CheckSheetType();
				return this.GetXmlNodeBool(outLineApplyStylePath);
			}
			set
			{
				this.CheckSheetType();
				this.SetXmlNodeString(outLineApplyStylePath, value ? "1" : "0");
			}
		}

		/// <summary>
		/// Color of the sheet tab
		/// </summary>
		public Color TabColor
		{
			get
			{
				string col = this.GetXmlNodeString(tabColorPath);
				if (col == "")
				{
					return Color.Empty;
				}
				else
				{
					return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
				}
			}
			set
			{
				this.SetXmlNodeString(tabColorPath, value.ToArgb().ToString("X"));
			}
		}

		/// <summary>
		/// Gets or sets the name of this worksheet's code module.
		/// </summary>
		internal string CodeModuleName
		{
			get
			{
				return this.GetXmlNodeString(codeModuleNamePath);
			}
			set
			{
				this.SetXmlNodeString(codeModuleNamePath, value);
			}
		}

		/// <summary>
		/// Gets this worksheet's VBA code module.
		/// </summary>
		public VBA.ExcelVBAModule CodeModule
		{
			get
			{
				if (this.Package.Workbook.VbaProject != null)
				{
					return this.Package.Workbook.VbaProject.Modules[CodeModuleName];
				}
				else
				{
					return null;
				}
			}
		}

		/// <summary>
		/// The XML document holding the worksheet data.
		/// All column, row, cell, pagebreak, merged cell and hyperlink-data are loaded into memory and removed from the document when loading the document.
		/// </summary>
		public XmlDocument WorksheetXml
		{
			get
			{
				return (_worksheetXml);
			}
		}

		/// <summary>
		/// Collection of comments
		/// </summary>
		public ExcelCommentCollection Comments
		{
			get
			{
				this.CheckSheetType();
				if (this._comments == null)
					this._comments = new ExcelCommentCollection(this.Package, this, this.NameSpaceManager);
				return this._comments;
			}
		}

		/// <summary>
		/// A reference to the header and footer class which allows you to
		/// set the header and footer for all odd, even and first pages of the worksheet
		/// </summary>
		/// <remarks>
		/// To format the text you can use the following format
		/// <list type="table">
		/// <listheader><term>Prefix</term><description>Description</description></listheader>
		/// <item><term>&amp;U</term><description>Underlined</description></item>
		/// <item><term>&amp;E</term><description>Double Underline</description></item>
		/// <item><term>&amp;K:xxxxxx</term><description>Color. ex &amp;K:FF0000 for red</description></item>
		/// <item><term>&amp;"Font,Regular Bold Italic"</term><description>Changes the font. Regular or Bold or Italic or Bold Italic can be used. ex &amp;"Arial,Bold Italic"</description></item>
		/// <item><term>&amp;nn</term><description>Change font size. nn is an integer. ex &amp;24</description></item>
		/// <item><term>&amp;G</term><description>Placeholder for images. Images can not be added by the library, but its possible to use in a template.</description></item>
		/// </list>
		/// </remarks>
		public ExcelHeaderFooter HeaderFooter
		{
			get
			{
				if (_headerFooter == null)
				{
					XmlNode headerFooterNode = TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager);
					if (headerFooterNode == null)
						headerFooterNode = CreateNode("d:headerFooter");
					_headerFooter = new ExcelHeaderFooter(NameSpaceManager, headerFooterNode, this);
				}
				return (_headerFooter);
			}
		}

		/// <summary>
		/// Printer settings
		/// </summary>
		public ExcelPrinterSettings PrinterSettings
		{
			get
			{
				var ps = new ExcelPrinterSettings(this.NameSpaceManager, this.TopNode, this);
				ps.SchemaNodeOrder = this.SchemaNodeOrder;
				return ps;
			}
		}

		/// <summary>
		/// Provides access to a range of cells
		/// </summary>
		public ExcelRange Cells
		{
			get
			{
				CheckSheetType();
				return new ExcelRange(this, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
			}
		}
		/// <summary>
		/// Provides access to the selected range of cells
		/// </summary>
		public ExcelRange SelectedRange
		{
			get
			{
				CheckSheetType();
				return new ExcelRange(this, View.SelectedRange);
			}
		}

		/// <summary>
		/// Addresses to merged ranges
		/// </summary>
		public MergeCellsCollection MergedCells
		{
			get
			{
				CheckSheetType();
				return _mergedCells;
			}
		}

		/// <summary>
		/// Dimension address for the worksheet.
		/// Top left cell to Bottom right.
		/// If the worksheet has no cells, null is returned
		/// </summary>
		public ExcelAddress Dimension
		{
			get
			{
				this.CheckSheetType();
				int fromRow, fromCol, toRow, toCol;
				if (this._values.GetDimension(out fromRow, out fromCol, out toRow, out toCol))
				{
					var addr = new ExcelAddress(fromRow, fromCol, toRow, toCol);
					addr._ws = Name;
					return addr;
				}
				else
				{
					return null;
				}
			}
		}

		/// <summary>
		/// Access to sheet protection properties
		/// </summary>
		public ExcelSheetProtection Protection
		{
			get
			{
				if (this._protection == null)
				{
					this._protection = new ExcelSheetProtection(this.NameSpaceManager, this.TopNode, this);
				}
				return this._protection;
			}
		}

		/// <summary>
		/// Gets the ranges that have been protected on this worksheet.
		/// </summary>
		public ExcelProtectedRangeCollection ProtectedRanges
		{
			get
			{
				if (this._protectedRanges == null)
					this._protectedRanges = new ExcelProtectedRangeCollection(NameSpaceManager, TopNode, this);
				return this._protectedRanges;
			}
		}

		/// <summary>
		/// Collection of drawing-objects like shapes, images and charts
		/// </summary>
		public ExcelDrawings Drawings
		{
			get
			{
				if (this._drawings == null)
				{
					this._drawings = new ExcelDrawings(this.Package, this);
				}
				return this._drawings;
			}
		}

		/// <summary>
		/// Tables defined in the worksheet.
		/// </summary>
		public ExcelTableCollection Tables
		{
			get
			{
				this.CheckSheetType();
				if (this.Workbook.NextTableID == int.MinValue)
					this.Workbook.ReadAllTables();
				if (this._tables == null)
				{
					this._tables = new ExcelTableCollection(this);
				}
				return this._tables;
			}
		}

		/// <summary>
		/// Pivottables defined in the worksheet.
		/// </summary>
		public ExcelPivotTableCollection PivotTables
		{
			get
			{
				this.CheckSheetType();
				if (this._pivotTables == null)
				{
					if (this.Workbook.NextPivotTableID == int.MinValue)
						this.Workbook.ReadAllTables();
					this._pivotTables = new ExcelPivotTableCollection(this);
				}
				return this._pivotTables;
			}
		}

		/// <summary>
		/// ConditionalFormatting defined in the worksheet. Use the Add methods to create ConditionalFormatting and add them to the worksheet. Then
		/// set the properties on the instance returned.
		/// </summary>
		/// <seealso cref="ExcelConditionalFormattingCollection"/>
		public ExcelConditionalFormattingCollection ConditionalFormatting
		{
			get
			{
				CheckSheetType();
				if (this._conditionalFormatting == null)
				{
					this._conditionalFormatting = new ExcelConditionalFormattingCollection(this);
				}
				return this._conditionalFormatting;
			}
		}

		/// <summary>
		/// Gets a collection representing extension-data (x14) conditional formatting rules defined on the worksheet.
		/// Adding new rules is currently unsupported.
		/// </summary>
		/// <seealso cref="ExcelConditionalFormattingCollection"/>
		public X14ConditionalFormattingCollection X14ConditionalFormatting
		{
			get
			{
				CheckSheetType();
				if (this._x14ConditionalFormatting == null)
				{
					this._x14ConditionalFormatting = new X14ConditionalFormattingCollection(this);
				}
				return this._x14ConditionalFormatting;
			}
		}

		/// <summary>
		/// DataValidation defined in the worksheet. Use the Add methods to create DataValidations and add them to the worksheet. Then
		/// set the properties on the instance returned.
		/// </summary>
		/// <seealso cref="ExcelDataValidationCollection"/>
		public ExcelDataValidationCollection DataValidations
		{
			get
			{
				this.CheckSheetType();
				if (this._dataValidation == null)
				{
					this._dataValidation = new ExcelDataValidationCollection(this);
				}
				return this._dataValidation;
			}
		}

		public ExcelX14DataValidationCollection X14DataValidations
		{
			get
			{
				this.CheckSheetType();
				if (this._x14dataValidation == null)
				{
					this._x14dataValidation = new ExcelX14DataValidationCollection(this);
				}
				return this._x14dataValidation;
			}
		}

		/// <summary>
		/// An image displayed as the background of the worksheet.
		/// </summary>
		public ExcelBackgroundImage BackgroundImage
		{
			get
			{
				if (this._backgroundImage == null)
				{
					this._backgroundImage = new ExcelBackgroundImage(this.NameSpaceManager, this.TopNode, this);
				}
				return this._backgroundImage;
			}
		}

		/// <summary>
		/// The workbook object
		/// </summary>
		public ExcelWorkbook Workbook
		{
			get
			{
				return this.Package.Workbook;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// A worksheet
		/// </summary>
		/// <param name="ns">Namespacemanager</param>
		/// <param name="excelPackage">Package</param>
		/// <param name="relID">Relationship ID</param>
		/// <param name="uriWorksheet">URI</param>
		/// <param name="sheetName">Name of the sheet</param>
		/// <param name="sheetID">Sheet id</param>
		/// <param name="positionID">Position</param>
		/// <param name="hide">hide</param>
		public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID,
							  Uri uriWorksheet, string sheetName, int sheetID, int positionID,
							  eWorkSheetHidden hide) :
			base(ns, null)
		{
			SchemaNodeOrder = new string[] { "sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "activeXControls", "webPublishItems", "tableParts", "extLst" };
			this.Package = excelPackage;
			_relationshipID = relID;
			_worksheetUri = uriWorksheet;
			_name = sheetName;
			_sheetID = sheetID;
			_positionID = positionID;
			Hidden = hide;

			_values = CellStore.Build<ExcelCoreValue>();
			_formulas = CellStore.Build<object>();
			_flags = CellStore.BuildFlagStore();
			_commentsStore = CellStore.Build<int>();
			_hyperLinks = CellStore.Build<Uri>();

			_names = new ExcelNamedRangeCollection(Workbook, this);

			this.CreateXml();
			this.TopNode = _worksheetXml.DocumentElement;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Provides access to an individual row within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <returns></returns>
		public ExcelRow Row(int row)
		{
			CheckSheetType();
			if (row < 1 || row > ExcelPackage.MaxRows)
			{
				throw (new ArgumentException("Row number out of bounds"));
			}
			return new ExcelRow(this, row);
		}

		/// <summary>
		/// Provides access to an individual column within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="col">The column number in the worksheet</param>
		/// <returns></returns>
		public ExcelColumn Column(int col)
		{
			this.CheckSheetType();
			if (col < 1 || col > ExcelPackage.MaxColumns)
			{
				throw (new ArgumentException("Column number out of bounds"));
			}
			var column = GetValueInner(0, col) as ExcelColumn;
			if (column != null)
			{

				if (column.ColumnMin != column.ColumnMax)
				{
					int maxCol = column.ColumnMax;
					column.ColumnMax = col;
					ExcelColumn copy = this.CopyColumn(column, col + 1, maxCol);
				}
			}
			else
			{
				int r = 0, c = col;
				if (_values.PrevCell(ref r, ref c))
				{
					column = this.GetValueInner(0, c) as ExcelColumn;
					int maxCol = column.ColumnMax;
					if (maxCol >= col)
					{
						column.ColumnMax = col - 1;
						if (maxCol > col)
						{
							ExcelColumn newC = this.CopyColumn(column, col + 1, maxCol);
						}
						return this.CopyColumn(column, col, col);
					}
				}

				column = new ExcelColumn(this, col);
				this.SetValueInner(0, col, column);
			}
			return column;
		}

		/// <summary>
		/// Returns the name of the worksheet
		/// </summary>
		/// <returns>The name of the worksheet</returns>
		public override string ToString()
		{
			return this.Name;
		}

		/// <summary>
		/// Make the current worksheet active.
		/// </summary>
		public void Select()
		{
			this.View.TabSelected = true;
		}

		/// <summary>
		/// Selects a range in the worksheet. The active cell is the topmost cell.
		/// Make the current worksheet active.
		/// </summary>
		/// <param name="Address">An address range</param>
		public void Select(string Address)
		{
			this.Select(Address, true);
		}

		/// <summary>
		/// Selects a range in the worksheet. The actice cell is the topmost cell.
		/// </summary>
		/// <param name="address">A range of cells</param>
		/// <param name="selectSheet">Make the sheet active</param>
		public void Select(string address, bool selectSheet)
		{
			this.CheckSheetType();
			int fromCol, fromRow, toCol, toRow;
			//Get rows and columns and validate as well
			ExcelCellBase.GetRowColFromAddress(address, out fromRow, out fromCol, out toRow, out toCol);

			if (selectSheet)
			{
				View.TabSelected = true;
			}
			this.View.SelectedRange = address;
			this.View.ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
		}

		/// <summary>
		/// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
		/// Make the current worksheet active.
		/// </summary>
		/// <param name="address">An address range</param>
		public void Select(ExcelAddress address)
		{
			this.CheckSheetType();
			this.Select(address, true);
		}

		/// <summary>
		/// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
		/// </summary>
		/// <param name="address">A range of cells</param>
		/// <param name="selectSheet">Make the sheet active</param>
		public void Select(ExcelAddress address, bool selectSheet)
		{

			this.CheckSheetType();
			if (selectSheet)
			{
				this.View.TabSelected = true;
			}
			string selAddress = ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column) + ":" + ExcelCellBase.GetAddress(address.End.Row, address.End.Column);
			if (address.Addresses != null)
			{
				foreach (var a in address.Addresses)
				{
					selAddress += " " + ExcelCellBase.GetAddress(a.Start.Row, a.Start.Column) + ":" + ExcelCellBase.GetAddress(a.End.Row, a.End.Column);
				}
			}
			this.View.SelectedRange = selAddress;
			this.View.ActiveCell = ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column);
		}

		/// <summary>
		/// Inserts a new row into the spreadsheet.  Existing rows below the position are
		/// shifted down.  All formula are updated to take account of the new row.
		/// </summary>
		/// <param name="rowFrom">The position of the new row</param>
		/// <param name="rows">Number of rows to insert</param>
		public void InsertRow(int rowFrom, int rows)
		{
			this.InsertRow(rowFrom, rows, 0);
		}

		/// <summary>
		/// Inserts a new row into the spreadsheet.  Existing rows below the position are
		/// shifted down.  All formula are updated to take account of the new row.
		/// </summary>
		/// <param name="rowFrom">The position of the new row</param>
		/// <param name="rows">Number of rows to insert.</param>
		/// <param name="copyStylesFromRow">Copy Styles from this row. Applied to all inserted rows</param>
		public void InsertRow(int rowFrom, int rows, int copyStylesFromRow)
		{
			this.CheckSheetType();
			var d = Dimension;

			if (rowFrom < 1)
			{
				throw (new ArgumentOutOfRangeException("rowFrom can't be lesser that 1"));
			}

			//Check that cells aren't shifted outside the boundries
			if (d != null && d.End.Row > rowFrom && d.End.Row + rows > ExcelPackage.MaxRows)
			{
				throw (new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet."));
			}

			lock (this)
			{
				this._values.Insert(rowFrom, 0, rows, 0);
				this._formulas.Insert(rowFrom, 0, rows, 0);
				this._hyperLinks.Insert(rowFrom, 0, rows, 0);
				this._flags.Insert(rowFrom, 0, rows, 0);
				this.Comments.Insert(rowFrom, 0, rows, 0);
				this.Names.Insert(rowFrom, 0, rows, 0, this);

				foreach (var f in this._sharedFormulas.Values)
				{
					if (f.StartRow >= rowFrom) f.StartRow += rows;
					var a = new ExcelAddress(f.Address);
					if (a._fromRow >= rowFrom)
					{
						a._fromRow += rows;
						a._toRow += rows;
					}
					else if (a._toRow >= rowFrom)
					{
						a._toRow += rows;
					}
					f.Address = ExcelAddress.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
					f.Formula = this.Package.FormulaManager.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0, this.Name, this.Name);
				}
				var cse = _formulas.GetEnumerator();
				while (cse.MoveNext())
				{
					if (cse.Value is string)
					{
						cse.Value = this.Package.FormulaManager.UpdateFormulaReferences(cse.Value.ToString(), rows, 0, rowFrom, 0, this.Name, this.Name);
					}
				}
				this.FixMergedCellsRow(rowFrom, rows, false);
				if (copyStylesFromRow > 0)
				{
					var cseS = _values.GetEnumerator(copyStylesFromRow, 0, copyStylesFromRow, ExcelPackage.MaxColumns); //Fixes issue 15068 , 15090
					while (cseS.MoveNext())
					{
						if (cseS.Value._styleId == 0) continue;
						for (var r = 0; r < rows; r++)
						{
							this.SetStyleInner(rowFrom + r, cseS.Column, cseS.Value._styleId);
						}
					}
					var newOutlineLevel = this.Row(copyStylesFromRow + rows).OutlineLevel;
					for (var r = 0; r < rows; r++)
					{
						this.Row(rowFrom + r).OutlineLevel = newOutlineLevel;
					}
				}
				this.UpdateSparkLines(rows, rowFrom, 0, 0);
				foreach (var tbl in this.Tables)
				{
					tbl.Address = tbl.Address.AddRow(rowFrom, rows);
				}
				foreach (var ptbl in this.PivotTables)
				{
					if (rowFrom <= ptbl.Address.End.Row)
						ptbl.Address = ptbl.Address.AddRow(rowFrom, rows);
				}
				foreach (var cacheDefinition in this.Workbook.PivotCacheDefinitions)
				{
					if (cacheDefinition.CacheSource == eSourceType.Worksheet &&
						cacheDefinition.GetSourceRangeAddress()?.Worksheet == this &&
						cacheDefinition.GetSourceRangeAddress().IsName == false &&
						rowFrom <= cacheDefinition.GetSourceRangeAddress().End.Row)
						cacheDefinition.SetSourceRangeAddress(this, cacheDefinition.GetSourceRangeAddress().AddRow(rowFrom, rows).Address);
				}
				this.UpdateCharts(rows, 0, rowFrom, 0);
				// Update cross-sheet references.
				foreach (var sheet in this.Workbook.Worksheets.Where(sheet => sheet != this))
				{
					sheet.Names.Insert(rowFrom, 0, rows, 0, this);
					sheet.UpdateCrossSheetReferences(this.Name, rowFrom, rows, 0, 0);
				}
				this.Workbook.Names.Insert(rowFrom, 0, rows, 0, this);
				this.UpdateDataValidationRanges(rowFrom, rows, 0, 0);
				this.UpdateX14DataValidationRanges(rowFrom, rows, 0, 0);
			}
		}

		/// <summary>
		/// Inserts a new column into the spreadsheet.  Existing columns below the position are
		/// shifted down.  All formula are updated to take account of the new column.
		/// </summary>
		/// <param name="columnFrom">The position of the new column</param>
		/// <param name="columns">Number of columns to insert</param>
		public void InsertColumn(int columnFrom, int columns)
		{
			this.InsertColumn(columnFrom, columns, 0);
		}

		///<summary>
		/// Inserts a new column into the spreadsheet.  Existing column to the left are
		/// shifted.  All formula are updated to take account of the new column.
		/// </summary>
		/// <param name="columnFrom">The position of the new column</param>
		/// <param name="columns">Number of columns to insert.</param>
		/// <param name="copyStylesFromColumn">Copy Styles from this column. Applied to all inserted columns</param>
		public void InsertColumn(int columnFrom, int columns, int copyStylesFromColumn)
		{
			this.CheckSheetType();
			var d = Dimension;

			if (columnFrom < 1)
			{
				throw (new ArgumentOutOfRangeException("columnFrom can't be lesser that 1"));
			}
			//Check that cells aren't shifted outside the boundries
			if (d != null && d.End.Column > columnFrom && d.End.Column + columns > ExcelPackage.MaxColumns)
			{
				throw (new ArgumentOutOfRangeException("Can't insert. Columns will be shifted outside the boundries of the worksheet."));
			}

			lock (this)
			{
				this._values.Insert(0, columnFrom, 0, columns);
				this._formulas.Insert(0, columnFrom, 0, columns);
				this._hyperLinks.Insert(0, columnFrom, 0, columns);
				this._flags.Insert(0, columnFrom, 0, columns);
				this.Comments.Insert(0, columnFrom, 0, columns);
				this.Names.Insert(0, columnFrom, 0, columns, this);

				foreach (var f in _sharedFormulas.Values)
				{
					if (f.StartCol >= columnFrom) f.StartCol += columns;
					var a = new ExcelAddress(f.Address);
					if (a._fromCol >= columnFrom)
					{
						a._fromCol += columns;
						a._toCol += columns;
					}
					else if (a._toCol >= columnFrom)
					{
						a._toCol += columns;
					}
					f.Address = ExcelAddress.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
					f.Formula = this.Package.FormulaManager.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, this.Name, this.Name);
				}

				var cse = _formulas.GetEnumerator();
				while (cse.MoveNext())
				{
					if (cse.Value is string)
					{
						cse.Value = this.Package.FormulaManager.UpdateFormulaReferences(cse.Value.ToString(), 0, columns, 0, columnFrom, this.Name, this.Name);
					}
				}

				this.FixMergedCellsColumn(columnFrom, columns, false);

				var csec = _values.GetEnumerator(0, 1, 0, ExcelPackage.MaxColumns);
				var lst = new List<ExcelColumn>();
				foreach (var val in csec)
				{
					var col = val._value;
					if (col is ExcelColumn)
					{
						lst.Add((ExcelColumn)col);
					}
				}

				for (int i = lst.Count - 1; i >= 0; i--)
				{
					var c = lst[i];
					if (c._columnMin >= columnFrom)
					{
						if (c._columnMin + columns <= ExcelPackage.MaxColumns)
						{
							c._columnMin += columns;
						}
						else
						{
							c._columnMin = ExcelPackage.MaxColumns;
						}

						if (c._columnMax + columns <= ExcelPackage.MaxColumns)
						{
							c._columnMax += columns;
						}
						else
						{
							c._columnMax = ExcelPackage.MaxColumns;
						}
					}
					else if (c._columnMax >= columnFrom)
					{
						var cc = c._columnMax - columnFrom;
						c._columnMax = columnFrom - 1;
						this.CopyColumn(c, columnFrom + columns, columnFrom + columns + cc);
					}
				}

				//Copy style from another column?
				if (copyStylesFromColumn > 0)
				{
					if (copyStylesFromColumn >= columnFrom)
					{
						copyStylesFromColumn += columns;
					}

					//Get styles to a cached list,
					var l = new List<int[]>();
					var sce = _values.GetEnumerator(0, copyStylesFromColumn, ExcelPackage.MaxRows, copyStylesFromColumn);
					lock (sce)
					{
						while (sce.MoveNext())
						{
							if (sce.Value._styleId == 0) continue;
							l.Add(new int[] { sce.Row, sce.Value._styleId });
						}
					}

					//Set the style id's from the list.
					foreach (var sc in l)
					{
						for (var c = 0; c < columns; c++)
						{
							if (sc[0] == 0)
							{
								var col = Column(columnFrom + c);   //Create the column
								col.StyleID = sc[1];
							}
							else
							{
								this.SetStyleInner(sc[0], columnFrom + c, sc[1]);
							}
						}
					}
					var newOutlineLevel = this.Column(copyStylesFromColumn).OutlineLevel;
					for (var c = 0; c < columns; c++)
					{
						this.Column(columnFrom + c).OutlineLevel = newOutlineLevel;
					}
				}
				this.UpdateSparkLines(0, 0, columns, columnFrom);
				//Adjust tables
				foreach (var tbl in this.Tables)
				{
					if (columnFrom > tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
					{
						ExcelWorksheet.InsertTableColumns(columnFrom, columns, tbl);
					}

					tbl.Address = tbl.Address.AddColumn(columnFrom, columns);
				}
				foreach (var ptbl in this.PivotTables)
				{
					if (columnFrom <= ptbl.Address.End.Column)
						ptbl.Address = ptbl.Address.AddColumn(columnFrom, columns);
				}
				foreach (var cacheDefinition in this.Workbook.PivotCacheDefinitions)
				{
					if (cacheDefinition.CacheSource == eSourceType.Worksheet &&
						cacheDefinition.GetSourceRangeAddress()?.Worksheet == this &&
						cacheDefinition.GetSourceRangeAddress().IsName == false &&
						columnFrom <= cacheDefinition.GetSourceRangeAddress().End.Row)
						cacheDefinition.SetSourceRangeAddress(this, cacheDefinition.GetSourceRangeAddress().AddColumn(columnFrom, columns).Address);
				}
				this.UpdateCharts(0, columns, 0, columnFrom);
				// Update cross-sheet references.
				foreach (var sheet in this.Workbook.Worksheets.Where(sheet => sheet != this))
				{
					sheet.Names.Insert(0, columnFrom, 0, columns, this);
					sheet.UpdateCrossSheetReferences(this.Name, 0, 0, columnFrom, columns);
				}
				this.Workbook.Names.Insert(0, columnFrom, 0, columns, this);
				this.UpdateDataValidationRanges(0, 0, columnFrom, columns);
				this.UpdateX14DataValidationRanges(0, 0, columnFrom, columns);
			}
		}

		/// <summary>
		/// Delete the specified row from the worksheet.
		/// </summary>
		/// <param name="row">A row to be deleted</param>
		public void DeleteRow(int row)
		{
			this.DeleteRow(row, 1);
		}

		/// <summary>
		/// Delete the specified row from the worksheet.
		/// </summary>
		/// <param name="rowFrom">The start row</param>
		/// <param name="rows">Number of rows to delete</param>
		public void DeleteRow(int rowFrom, int rows)
		{
			this.CheckSheetType();
			if (rowFrom < 1 || rowFrom + rows > ExcelPackage.MaxRows)
			{
				throw (new ArgumentException("Row out of range. Spans from 1 to " + ExcelPackage.MaxRows.ToString(CultureInfo.InvariantCulture)));
			}
			lock (this)
			{
				this._values.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
				this._formulas.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
				this._flags.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
				this._hyperLinks.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
				this.Comments.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
				this.Names.Delete(rowFrom, 0, rows, 0, this);

				this.AdjustFormulasRow(rowFrom, rows);
				this.FixMergedCellsRow(rowFrom, rows, true);

				foreach (var table in this.Tables)
				{
					table.Address = table.Address.DeleteRow(rowFrom, rows);
				}
				foreach (var pivotTable in this.PivotTables)
				{
					if (rowFrom <= pivotTable.Address.End.Row)
						pivotTable.Address = pivotTable.Address.DeleteRow(rowFrom, rows);
				}
				foreach (var cacheDefinition in this.Workbook.PivotCacheDefinitions)
				{
					if (cacheDefinition.CacheSource == eSourceType.Worksheet &&
						cacheDefinition.GetSourceRangeAddress()?.Worksheet == this &&
						cacheDefinition.GetSourceRangeAddress().IsName == false &&
						rowFrom <= cacheDefinition.GetSourceRangeAddress().End.Row)
					{
						var shiftedAddress = cacheDefinition.GetSourceRangeAddress().DeleteRow(rowFrom, rows)?.Address;
						var setAddress = shiftedAddress ?? ExcelErrorValue.Create(eErrorType.Ref).ToString();
						cacheDefinition.SetSourceRangeAddress(this, setAddress);
					}
				}
				foreach (var sheet in this.Workbook.Worksheets.Where(sheet => sheet != this))
				{
					sheet.Names.Delete(rowFrom, 0, rows, 0, this);
					sheet.UpdateCrossSheetReferences(this.Name, rowFrom, -rows, 0, 0);
				}
				this.Workbook.Names.Delete(rowFrom, 0, rows, 0, this);
				this.UpdateSparkLines(-rows, rowFrom, 0, 0);
				this.UpdateCharts(-rows, 0, rowFrom, 0);
				this.UpdateDataValidationRanges(rowFrom, -rows, 0, 0);
				this.UpdateX14DataValidationRanges(rowFrom, -rows, 0, 0);
				this.UpdateConditionalFormatting(rowFrom, rows, 0, 0);
				this.UpdateX14ConditionalFormatting(rowFrom, rows, 0, 0);
			}
		}

		/// <summary>
		/// Delete the specified column from the worksheet.
		/// </summary>
		/// <param name="column">The column to be deleted</param>
		public void DeleteColumn(int column)
		{
			this.DeleteColumn(column, 1);
		}

		/// <summary>
		/// Delete the specified column from the worksheet.
		/// </summary>
		/// <param name="columnFrom">The start column</param>
		/// <param name="columns">Number of columns to delete</param>
		public void DeleteColumn(int columnFrom, int columns)
		{
			if (columnFrom < 1 || columnFrom + columns > ExcelPackage.MaxColumns)
			{
				throw (new ArgumentException("Column out of range. Spans from 1 to " + ExcelPackage.MaxColumns.ToString(CultureInfo.InvariantCulture)));
			}
			lock (this)
			{
				var col = this.GetValueInner(0, columnFrom) as ExcelColumn;
				if (col == null)
				{
					var r = 0;
					var c = columnFrom;
					if (this._values.PrevCell(ref r, ref c))
					{
						col = this.GetValueInner(0, c) as ExcelColumn;
						if (col._columnMax >= columnFrom)
						{
							col.ColumnMax = columnFrom - 1;
						}
					}
				}

				this._values.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
				this._formulas.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
				this._flags.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
				this._hyperLinks.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
				this.Comments.Delete(0, columnFrom, 0, columns);
				this.Names.Delete(0, columnFrom, 0, columns, this);

				this.AdjustFormulasColumn(columnFrom, columns);
				this.FixMergedCellsColumn(columnFrom, columns, true);

				var csec = _values.GetEnumerator(0, columnFrom, 0, ExcelPackage.MaxColumns);
				foreach (var val in csec)
				{
					var column = val._value;
					if (column is ExcelColumn)
					{
						var c = (ExcelColumn)column;
						if (c._columnMin >= columnFrom)
						{
							c._columnMin -= columns;
							c._columnMax -= columns;
						}
					}
				}

				foreach (var tbl in this.Tables)
				{
					if (columnFrom >= tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
					{
						var node = tbl.Columns[0].TopNode.ParentNode;
						var ix = columnFrom - tbl.Address.Start.Column;
						for (int i = 0; i < columns; i++)
						{
							if (node.ChildNodes.Count > ix)
							{
								node.RemoveChild(node.ChildNodes[ix]);
							}
						}
						tbl._cols = new ExcelTableColumnCollection(tbl);
					}

					tbl.Address = tbl.Address.DeleteColumn(columnFrom, columns);
				}
				foreach (var ptbl in this.PivotTables)
				{
					if (columnFrom <= ptbl.Address.End.Column)
						ptbl.Address = ptbl.Address.DeleteColumn(columnFrom, columns);
				}
				foreach (var cacheDefinition in this.Workbook.PivotCacheDefinitions)
				{
					if (cacheDefinition.CacheSource == eSourceType.Worksheet &&
						cacheDefinition.GetSourceRangeAddress()?.Worksheet == this &&
						cacheDefinition.GetSourceRangeAddress().IsName == false &&
						columnFrom <= cacheDefinition.GetSourceRangeAddress().End.Row)
					{
						var shiftedAddress = cacheDefinition.GetSourceRangeAddress().DeleteColumn(columnFrom, columns)?.Address;
						var setAddress = shiftedAddress ?? ExcelErrorValue.Create(eErrorType.Ref).ToString();
						cacheDefinition.SetSourceRangeAddress(this, setAddress);
					}
				}
				foreach (var sheet in this.Workbook.Worksheets.Where(sheet => sheet != this))
				{
					sheet.Names.Delete(0, columnFrom, 0, columns, this);
					sheet.UpdateCrossSheetReferences(this.Name, 0, 0, columnFrom, -columns);
				}
				this.Workbook.Names.Delete(0, columnFrom, 0, columns, this);
				this.UpdateCharts(0, -columns, 0, columnFrom);
				this.UpdateSparkLines(0, 0, -columns, columnFrom);
				this.UpdateDataValidationRanges(0, 0, columnFrom, -columns);
				this.UpdateX14DataValidationRanges(0, 0, columnFrom, -columns);
				this.UpdateConditionalFormatting(0, 0, columnFrom, columns);
				this.UpdateX14ConditionalFormatting(0, 0, columnFrom, columns);
			}
		}

		/// <summary>
		/// Deletes the specified row from the worksheet.
		/// </summary>
		/// <param name="rowFrom">The number of the start row to be deleted</param>
		/// <param name="rows">Number of rows to delete</param>
		/// <param name="shiftOtherRowsUp">Not used. Rows are always shifted</param>
		public void DeleteRow(int rowFrom, int rows, bool shiftOtherRowsUp)
		{
			this.DeleteRow(rowFrom, rows);
		}

		/// <summary>
		/// Get the cell value from thw worksheet
		/// </summary>
		/// <param name="Row">The row number</param>
		/// <param name="Column">The row number</param>
		/// <returns>The value</returns>
		public object GetValue(int Row, int Column)
		{
			this.CheckSheetType();
			var v = this.GetValueInner(Row, Column);
			if (v != null)
			{
				if (this._flags.GetFlagValue(Row, Column, CellFlags.RichText))
				{
					return (object)this.Cells[Row, Column].RichText.Text;
				}
				else
				{
					return v;
				}
			}
			else
			{
				return null;
			}
		}

		/// <summary>
		/// Get a strongly typed cell value from the worksheet
		/// </summary>
		/// <typeparam name="T">The type</typeparam>
		/// <param name="Row">The row number</param>
		/// <param name="Column">The row number</param>
		/// <returns>The value. If the value can't be converted to the specified type, the default value will be returned</returns>
		public T GetValue<T>(int Row, int Column)
		{
			this.CheckSheetType();
			var v = this.GetValueInner(Row, Column);
			if (v == null)
			{
				return default(T);
			}

			if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
			{
				return (T)(object)this.Cells[Row, Column].RichText.Text;
			}
			else
			{
				return this.GetTypedValue<T>(v);
			}
		}

		//Thanks to Michael Tran for parts of this method
		internal T GetTypedValue<T>(object v)
		{
			if (v == null)
			{
				return default(T);
			}
			Type fromType = v.GetType();
			Type toType = typeof(T);
			Type toType2 = (toType.IsGenericType && toType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
				? Nullable.GetUnderlyingType(toType)
				: null;
			if (fromType == toType || fromType == toType2)
			{
				return (T)v;
			}
			var cnv = TypeDescriptor.GetConverter(fromType);
			if (toType == typeof(DateTime) || toType2 == typeof(DateTime))    //Handle dates
			{
				if (fromType == typeof(TimeSpan))
				{
					return ((T)(object)(new DateTime(((TimeSpan)v).Ticks)));
				}
				else if (fromType == typeof(string))
				{
					DateTime dt;
					if (DateTime.TryParse(v.ToString(), out dt))
					{
						return (T)(object)(dt);
					}
					else
					{
						return default(T);
					}

				}
				else
				{
					if (cnv.CanConvertTo(typeof(double)))
					{
						return (T)(object)(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))));
					}
					else
					{
						return default(T);
					}
				}
			}
			else if (toType == typeof(TimeSpan) || toType2 == typeof(TimeSpan))    //Handle timespan
			{
				if (fromType == typeof(DateTime))
				{
					return ((T)(object)(new TimeSpan(((DateTime)v).Ticks)));
				}
				else if (fromType == typeof(string))
				{
					TimeSpan ts;
					if (TimeSpan.TryParse(v.ToString(), out ts))
					{
						return (T)(object)(ts);
					}
					else
					{
						return default(T);
					}
				}
				else
				{
					if (cnv.CanConvertTo(typeof(double)))
					{

						return (T)(object)(new TimeSpan(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))).Ticks));
					}
					else
					{
						try
						{
							// Issue 14682 -- "GetValue<decimal>() won't convert strings"
							// As suggested, after all special cases, all .NET to do it's
							// preferred conversion rather than simply returning the default
							return (T)Convert.ChangeType(v, typeof(T));
						}
						catch (Exception)
						{
							// This was the previous behaviour -- no conversion is available.
							return default(T);
						}
					}
				}
			}
			else
			{
				if (cnv.CanConvertTo(toType))
				{
					return (T)cnv.ConvertTo(v, typeof(T));
				}
				else
				{
					if (toType2 != null)
					{
						toType = toType2;
						if (cnv.CanConvertTo(toType))
						{
							return (T)cnv.ConvertTo(v, toType); //Fixes issue 15377
						}
					}

					if (fromType == typeof(double) && toType == typeof(decimal))
					{
						return (T)(object)Convert.ToDecimal(v);
					}
					else if (fromType == typeof(decimal) && toType == typeof(double))
					{
						return (T)(object)Convert.ToDouble(v);
					}
					else
					{
						return default(T);
					}
				}
			}
		}

		/// <summary>
		/// Set the value of a cell
		/// </summary>
		/// <param name="Row">The row number</param>
		/// <param name="Column">The column number</param>
		/// <param name="Value">The value</param>
		public void SetValue(int Row, int Column, object Value)
		{
			this.CheckSheetType();
			if (Row < 1 || Column < 1 || Row > ExcelPackage.MaxRows && Column > ExcelPackage.MaxColumns)
			{
				throw new ArgumentOutOfRangeException("Row or Column out of range");
			}
			this.SetValueInner(Row, Column, Value);
		}

		/// <summary>
		/// Set the value of a cell
		/// </summary>
		/// <param name="Address">The Excel address</param>
		/// <param name="Value">The value</param>
		public void SetValue(string Address, object Value)
		{
			CheckSheetType();
			int row, col;
			ExcelAddress.GetRowCol(Address, out row, out col, true);
			if (row < 1 || col < 1 || row > ExcelPackage.MaxRows && col > ExcelPackage.MaxColumns)
			{
				throw new ArgumentOutOfRangeException("Address is invalid or out of range");
			}
			this.SetValueInner(row, col, Value);
		}

		/// <summary>
		/// Get MergeCell Index No
		/// </summary>
		/// <param name="row"></param>
		/// <param name="column"></param>
		/// <returns></returns>
		public int GetMergeCellId(int row, int column)
		{
			for (int i = 0; i < _mergedCells.Count; i++)
			{
				if (!string.IsNullOrEmpty(_mergedCells[i]))
				{
					ExcelRange range = Cells[_mergedCells[i]];

					if (range.Start.Row <= row && row <= range.End.Row)
					{
						if (range.Start.Column <= column && column <= range.End.Column)
						{
							return i + 1;
						}
					}
				}
			}
			return 0;
		}

		/// <summary>
		/// Removes all "autoFilter" nodes from the worksheet.
		/// </summary>
		public void RemoveAutoFilters()
		{
			this.DeleteAllNode("d:autoFilter");
			if (this.Names.ContainsKey("_xlnm._FilterDatabase"))
			{
				this.Names.Remove("_xlnm._FilterDatabase");
			}
			this.TopNode.SelectSingleNode("d:sheetPr", this.NameSpaceManager)?.Attributes?.RemoveNamedItem("filterMode");
		}

		/// <summary>
		/// Dispose of the worksheet and its child objects.
		/// </summary>
		public void Dispose()
		{
			this.DisposeInternal(_values);
			this.DisposeInternal(_formulas);
			this.DisposeInternal(_flags);
			this.DisposeInternal(_hyperLinks);
			this.DisposeInternal(_commentsStore);
			this.DisposeInternal(_formulaTokens);

			this._values = null;
			this._formulas = null;
			this._flags = null;
			this._hyperLinks = null;
			this._commentsStore = null;
			this._formulaTokens = null;

			this.Package = null;
			this._pivotTables = null;
			this._protection = null;
			if (this._sharedFormulas != null) this._sharedFormulas.Clear();
			this._sharedFormulas = null;
			this.SheetView = null;
			this._tables = null;
			this._conditionalFormatting = null;
			this._dataValidation = null;
			this._x14dataValidation = null;
			this._drawings = null;
		}

		/// <summary>
		/// Determine if two <see cref="ExcelWorksheet"/>s represent the same worksheet.
		/// </summary>
		/// <param name="x">One worksheet to compare.</param>
		/// <param name="y">The other worksheet to compare.</param>
		/// <returns></returns>
		public bool Equals(ExcelWorksheet x, ExcelWorksheet y)
		{
			return x.Name == y.Name && x.SheetID == y.SheetID && x.WorksheetXml.OuterXml == y.WorksheetXml.OuterXml;
		}

		/// <summary>
		///
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public int GetHashCode(ExcelWorksheet obj)
		{
			return obj.WorksheetXml.OuterXml.GetHashCode();
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Validates that a worksheet is not a ChartSheet.
		/// </summary>
		internal void CheckSheetType()
		{
			if (this is ExcelChartsheet)
			{
				throw (new NotSupportedException("This property or method is not supported for a Chartsheet"));
			}
		}

		/// <summary>
		/// Save the worksheet to Xml.
		/// </summary>
		internal void Save()
		{
			this.DeletePrinterSettings();

			if (this._worksheetXml != null)
			{

				if (!(this is ExcelChartsheet))
				{
					// save the header & footer (if defined)
					if (this._headerFooter != null)
						this.HeaderFooter.Save();

					var d = Dimension;
					if (d == null)
					{
						this.DeleteAllNode("d:dimension/@ref");
					}
					else
					{
						this.SetXmlNodeString("d:dimension/@ref", d.Address);
					}


					if (Drawings.Count == 0)
					{
						//Remove node if no drawings exists.
						this.DeleteNode("d:drawing");
					}

					this.SaveComments();
					this.HeaderFooter.SaveHeaderFooterImages();
					this.SaveTables();
					this.SavePivotTables();
					this.SparklineGroups.Save();
				}
			}
			if (this.Slicers.Slicers.Count > 0)
				this.Slicers.Save();

			if (Drawings.UriDrawing != null)
			{
				if (Drawings.Count == 0)
				{
					this.Part.DeleteRelationship(Drawings.DrawingRelationship.Id);
					this.Package.Package.DeletePart(Drawings.UriDrawing);
				}
				else
				{
					foreach (ExcelDrawing d in Drawings)
					{
						d.AdjustPositionAndSize();
						if (d is ExcelChart)
						{
							ExcelChart c = (ExcelChart)d;
							c.ChartXml.Save(c.Part.GetStream(FileMode.Create, FileAccess.Write));
						}
					}
					Packaging.ZipPackagePart partPack = Drawings.Part;
					this.Drawings.DrawingXml.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
				}
			}
		}

		/// <summary>
		/// Save this worksheet's xml to the specified stream.
		/// </summary>
		/// <param name="stream">The output stream.</param>
		/// <param name="compressionLevel">The ZIP compression level.</param>
		/// <param name="fileName">The name of the entry.</param>
		internal void SaveHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
		{
			//Init Zip
			stream.CodecBufferSize = 8096;
			stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
			stream.PutNextEntry(fileName);

			this.SaveXml(stream);
		}

		/// <summary>
		/// Set the totalling row function on a table to a particular totalling function type.
		/// </summary>
		/// <param name="tbl">The table to be updated.</param>
		/// <param name="col">The table column to be updated.</param>
		/// <param name="colNum">The column number to use (optional).</param>
		internal void SetTableTotalFunction(ExcelTable tbl, ExcelTableColumn col, int colNum = -1)
		{
			if (tbl.ShowTotal == false) return;
			if (colNum == -1)
			{
				for (int i = 0; i < tbl.Columns.Count; i++)
				{
					if (tbl.Columns[i].Name == col.Name)
					{
						colNum = tbl.Address._fromCol + i;
					}
				}
			}
			if (col.TotalsRowFunction == RowFunctions.Custom)
			{
				this.SetFormula(tbl.Address._toRow, colNum, col.TotalsRowFormula);
			}
			else if (col.TotalsRowFunction != RowFunctions.None)
			{
				switch (col.TotalsRowFunction)
				{
					case RowFunctions.Average:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "101"));
						break;
					case RowFunctions.CountNums:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "102"));
						break;
					case RowFunctions.Count:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "103"));
						break;
					case RowFunctions.Max:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "104"));
						break;
					case RowFunctions.Min:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "105"));
						break;
					case RowFunctions.StdDev:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "107"));
						break;
					case RowFunctions.Var:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "110"));
						break;
					case RowFunctions.Sum:
						this.SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "109"));
						break;
					default:
						throw (new Exception("Unknown RowFunction enum"));
				}
			}
			else
			{
				this.SetValueInner(tbl.Address._toRow, colNum, col.TotalsRowLabel);
			}
		}

		/// <summary>
		/// Set the formula at a specific cell to the specified <paramref name="value"/>.
		/// </summary>
		/// <param name="row">The row of the cell to insert at.</param>
		/// <param name="col">The column of the cell.</param>
		/// <param name="value">The new formula value.</param>
		internal void SetFormula(int row, int col, object value)
		{
			this._formulas.SetValue(row, col, value);
			if (!this.ExistsValueInner(row, col))
				this.SetValueInner(row, col, null);
		}

		/// <summary>
		/// Get the next ID from a shared formula or an Array formula
		/// Sharedforumlas will have an id from 0-x. Array formula ids start from 0x4000001-.
		/// </summary>
		/// <param name="isArray">If the formula is an array formula</param>
		/// <returns></returns>
		internal int GetMaxShareFunctionIndex(bool isArray)
		{
			int i = this._sharedFormulas.Count + 1;
			if (isArray)
				i |= 0x40000000;

			while (this._sharedFormulas.ContainsKey(i))
			{
				i++;
			}
			return i;
		}

		/// <summary>
		/// Set the HeaderFooter relationship ID.
		/// </summary>
		/// <param name="relID">The new relationship id.</param>
		internal void SetHFLegacyDrawingRel(string relID)
		{
			this.SetXmlNodeString("d:legacyDrawingHF/@r:id", relID);
		}

		/// <summary>
		/// Pre-process cells that use 1904-style dates.
		/// </summary>
		internal void UpdateCellsWithDate1904Setting()
		{
			var cse = _values.GetEnumerator();
			var offset = this.Workbook.Date1904 ? -ExcelWorkbook.date1904Offset : ExcelWorkbook.date1904Offset;
			while (cse.MoveNext())
			{
				if (cse.Value._value is DateTime)
				{
					try
					{
						double sdv = ((DateTime)cse.Value._value).ToOADate();
						sdv += offset;

						SetValueInner(cse.Row, cse.Column, DateTime.FromOADate(sdv));
					}
					catch
					{
					}
				}
			}
		}

		/// <summary>
		/// Get the formula from a specific cell.
		/// </summary>
		/// <param name="row">The cell's row.</param>
		/// <param name="col">The cell's column.</param>
		/// <returns>The formula at the specified location.</returns>
		internal string GetFormula(int row, int col)
		{
			var v = this._formulas.GetValue(row, col);
			if (v is int)
			{
				return this._sharedFormulas[(int)v].GetFormula(row, col, Name);
			}
			else if (v != null)
			{
				return v.ToString();
			}
			else
			{
				return "";
			}
		}

		/// <summary>
		/// Get the formula from a specified location in R1C1 format.
		/// </summary>
		/// <param name="row">The row of the cell.</param>
		/// <param name="col">The column of the cell.</param>
		/// <returns>The formula at the specified location, in R1C1 format.</returns>
		internal string GetFormulaR1C1(int row, int col)
		{
			var v = this._formulas.GetValue(row, col);
			if (v is int)
			{
				var sf = this._sharedFormulas[(int)v];
				return ExcelCellBase.TranslateToR1C1(sf.Formula, sf.StartRow, sf.StartCol);
			}
			else if (v != null)
			{
				return ExcelCellBase.TranslateToR1C1(v.ToString(), row, col);
			}
			else
			{
				return "";
			}
		}

		/// <summary>
		/// Set the name of this worksheet's code module.
		/// </summary>
		/// <param name="value">The new name for the code module.</param>
		internal void CodeNameChange(string value)
		{
			this.CodeModuleName = value;
		}

		/// <summary>
		/// Get accessor of sheet value
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <returns>cell value</returns>
		internal object GetValueInner(int row, int col)
		{
			return this._values.GetValue(row, col)._value;
		}

		/// <summary>
		/// Get accessor of sheet styleId
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <returns>cell styleId</returns>
		internal int GetStyleInner(int row, int col)
		{
			return this._values.GetValue(row, col)._styleId;
		}

		/// <summary>
		/// Set accessor of sheet value
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <param name="value">value</param>
		internal void SetValueInner(int row, int col, object value)
		{
			this._values.SetValue(row, col, new ExcelCoreValue() { _value = value, _styleId = this._values.GetValue(row, col)._styleId });
		}

		/// <summary>
		/// Set accessor of sheet styleId
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <param name="styleId">styleId</param>
		internal void SetStyleInner(int row, int col, int styleId)
		{
			this._values.SetValue(row, col, new ExcelCoreValue() { _value = this._values.GetValue(row, col)._value, _styleId = styleId });
		}

		/// <summary>
		/// Bulk(Range) set accessor of sheet value, for value array
		/// </summary>
		/// <param name="fromRow">start row</param>
		/// <param name="fromColumn">start column</param>
		/// <param name="toRow">end row</param>
		/// <param name="toColumn">end column</param>
		/// <param name="values">set values</param>
		internal void SetRangeValueInner(int fromRow, int fromColumn, int toRow, int toColumn, object[,] values)
		{
			var rowBound = values.GetUpperBound(0);
			var colBound = values.GetUpperBound(1);
			for (int row = fromRow; row <= toRow; row++)
			{
				for (int column = fromColumn; column <= toColumn; column++)
				{
					object val = null;
					if (rowBound >= row - fromRow && colBound >= column - fromColumn)
					{
						val = ((object[,])values)[row - fromRow, column - fromColumn];
					}
					this.SetValueInner(row, column, val);
				}
			}
		}

		/// <summary>
		/// Existance check of sheet value
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <returns>is exists</returns>
		internal bool ExistsValueInner(int row, int col)
		{
			return (this._values.GetValue(row, col)._value != null);
		}

		/// <summary>
		/// Existance check of sheet styleId
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <returns>is exists</returns>
		internal bool ExistsStyleInner(int row, int col)
		{
			return (this._values.GetValue(row, col)._styleId != 0);
		}

		/// <summary>
		/// Existance check of sheet value
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <param name="value"></param>
		/// <returns>is exists</returns>
		internal bool ExistsValueInner(int row, int col, ref object value)
		{
			value = this._values.GetValue(row, col)._value;
			return (value != null);
		}

		/// <summary>
		/// Existance check of sheet styleId
		/// </summary>
		/// <param name="row">row</param>
		/// <param name="col">column</param>
		/// <param name="styleId"></param>
		/// <returns>is exists</returns>
		internal bool ExistsStyleInner(int row, int col, ref int styleId)
		{
			styleId = this._values.GetValue(row, col)._styleId;
			return (styleId != 0);
		}

		/// <summary>
		/// Get the ExcelColumn for column (span ColumnMin and ColumnMax)
		/// </summary>
		/// <param name="column"></param>
		/// <returns></returns>
		internal ExcelColumn GetColumn(int column)
		{
			var c = GetValueInner(0, column) as ExcelColumn;
			if (c == null)
			{
				int row = 0, col = column;
				if (_values.PrevCell(ref row, ref col))
				{
					c = GetValueInner(0, col) as ExcelColumn;
					if (c != null && c.ColumnMax >= column)
					{
						return c;
					}
					return null;
				}
			}
			return c;
		}

		/// <summary>
		/// Copy a column to the specified <paramref name="col"/>.
		/// </summary>
		/// <param name="c">The column to copy.</param>
		/// <param name="col">The destination column.</param>
		/// <param name="maxCol">The maximum allowable column value.</param>
		/// <returns></returns>
		internal ExcelColumn CopyColumn(ExcelColumn c, int col, int maxCol)
		{
			ExcelColumn newC = new ExcelColumn(this, col);
			newC.ColumnMax = maxCol < ExcelPackage.MaxColumns ? maxCol : ExcelPackage.MaxColumns;
			if (c.StyleName != "")
				newC.StyleName = c.StyleName;
			else
				newC.StyleID = c.StyleID;

			newC.OutlineLevel = c.OutlineLevel;
			newC.Phonetic = c.Phonetic;
			newC.BestFit = c.BestFit;
			this.SetValueInner(0, col, newC);
			newC._width = c._width;
			newC._hidden = c._hidden;
			return newC;
		}

		/// <summary>
		/// Returns the style ID given a style name.
		/// The style ID will be created if not found, but only if the style name exists!
		/// </summary>
		/// <param name="StyleName"></param>
		/// <returns></returns>
		internal int GetStyleID(string StyleName)
		{
			ExcelNamedStyleXml namedStyle = null;
			Workbook.Styles.NamedStyles.FindByID(StyleName, ref namedStyle);
			if (namedStyle.XfId == int.MinValue)
			{
				namedStyle.XfId = Workbook.Styles.CellXfs.FindIndexByID(namedStyle.Style.Id);
			}
			return namedStyle.XfId;
		}

		/// <summary>
		/// Clear the Data Validation indicator.
		/// </summary>
		internal void ClearValidations()
		{
			this._dataValidation = null;
		}

		/// <summary>
		/// Clear the Data Validation indicator.
		/// </summary>
		internal void ClearX14Validations()
		{
			this._x14dataValidation = null;
		}

		internal void AdjustFormulasRow(int rowFrom, int rows)
		{
			var delSF = new List<int>();
			foreach (var sf in _sharedFormulas.Values)
			{
				var a = new ExcelAddress(sf.Address).DeleteRow(rowFrom, rows);
				if (a == null)
				{
					delSF.Add(sf.Index);
				}
				else
				{
					sf.Address = a.Address;
					if (sf.StartRow > rowFrom)
					{
						var r = Math.Min(sf.StartRow - rowFrom, rows);
						sf.Formula = this.Package.FormulaManager.UpdateFormulaReferences(sf.Formula, -r, 0, rowFrom, 0, this.Name, this.Name);
						sf.StartRow -= r;
					}
				}
			}
			foreach (var ix in delSF)
			{
				this._sharedFormulas.Remove(ix);
			}
			delSF = null;
			var cse = _formulas.GetEnumerator(1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
			while (cse.MoveNext())
			{
				if (cse.Value is string)
				{
					cse.Value = this.Package.FormulaManager.UpdateFormulaReferences(cse.Value.ToString(), -rows, 0, rowFrom, 0, this.Name, this.Name);
				}
			}
		}

		internal void AdjustFormulasColumn(int columnFrom, int columns)
		{
			var delSF = new List<int>();
			foreach (var sf in this._sharedFormulas.Values)
			{
				var a = new ExcelAddress(sf.Address).DeleteColumn(columnFrom, columns);
				if (a == null)
				{
					delSF.Add(sf.Index);
				}
				else
				{
					sf.Address = a.Address;
					if (sf.StartCol > columnFrom)
					{
						var c = Math.Min(sf.StartCol - columnFrom, columns);
						sf.Formula = this.Package.FormulaManager.UpdateFormulaReferences(sf.Formula, 0, -c, 0, 1, this.Name, this.Name);
						sf.StartCol -= c;
					}
				}
			}
			foreach (var ix in delSF)
			{
				this._sharedFormulas.Remove(ix);
			}
			delSF = null;
			var cse = _formulas.GetEnumerator(1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
			while (cse.MoveNext())
			{
				if (cse.Value is string)
				{
					cse.Value = this.Package.FormulaManager.UpdateFormulaReferences(cse.Value.ToString(), 0, -columns, 0, columnFrom, this.Name, this.Name);
				}
			}
		}
		#endregion

		#region Private Methods
		private void UpdateCharts(int rows, int columns, int rowFrom, int colFrom)
		{
			// Only two-cell-anchor drawings can be deleted by DeleteRow/Column operations. 
			var deletedDrawings = this.Drawings.Where(drawing => drawing.EditAs == eEditAs.TwoCell && this.RangeIsBeingDeleted(drawing.From.Row, drawing.To.Row, drawing.From.Column, drawing.To.Column, rowFrom, rows, colFrom, columns)).ToArray();
			foreach (var drawing in deletedDrawings)
			{
				this.Drawings.Remove(drawing);
			}

			HashSet<ExcelChart> uniqueChartTypes = new HashSet<ExcelChart>();
			foreach (var sheet in this.Workbook.Worksheets)
			{
				foreach (ExcelDrawing drawing in sheet.Drawings)
				{
					bool isUnique = false;
					if (sheet == this)
					{
						drawing.AdjustPositionAndSize();
						int newFromRow = drawing.From.Row;
						int newFromColumn = drawing.From.Column;
						int newToRow = drawing.To.Row;
						int newToColumn = drawing.To.Column;
						int newFromRowOffset = drawing.From.RowOff;
						int newFromColumnOffset = drawing.From.ColumnOff;
						if (drawing.EditAs == eEditAs.TwoCell)
						{
							if (drawing.To.Row > rowFrom)
								newToRow += rows;
							if (drawing.To.Column > colFrom)
								newToColumn += columns;
						}
						if (drawing.EditAs == eEditAs.OneCell || drawing.EditAs == eEditAs.TwoCell)
						{
							if (drawing.From.Row > rowFrom)
							{
								newFromRow += rows;
								if (rows < 0 && rowFrom > newFromRow)
								{
									newFromRow = rowFrom;
									newFromRowOffset = 0;
								}
							}
							if (drawing.From.Column > colFrom)
							{
								newFromColumn += columns;
								if (columns < 0 && colFrom > newFromColumn)
								{
									newFromColumn = colFrom;
									newFromColumnOffset = 0;
								}
							}
						}
						newFromColumn = newFromColumn < 1 ? 1 : newFromColumn;
						newFromRow = newFromRow < 1 ? 1 : newFromRow;
						drawing.SetPosition(newFromRow, newFromRowOffset, newFromColumn, newFromColumnOffset, newToRow, drawing.To.RowOff, newToColumn, drawing.To.ColumnOff);
					}
					// The chart Plot Area contains one copy of a chart for each series in that chart.
					// A chart Plot Area can also have multiple distinct charts (such as when a bar chart and a line chart are plotted in the same area).
					// This captures the behavior of a "Combo Chart".
					if (drawing is ExcelChart chartBase)
					{
						foreach (var chart in chartBase.PlotArea.ChartTypes)
						{
							isUnique = uniqueChartTypes.Add(chart);
							if (isUnique)
							{
								foreach (ExcelChartSerie serie in chart.Series)
								{
									if (!string.IsNullOrEmpty(serie.Series))
									{
										var excelAddress = new ExcelAddress(serie.Series);
										serie.Series = this.Package.FormulaManager.UpdateFormulaReferences(serie.Series, rows, columns, rowFrom, colFrom, excelAddress.WorkSheet, this.Name);
									}
									if (!string.IsNullOrEmpty(serie.XSeries))
									{
										var excelAddress = new ExcelAddress(serie.XSeries);
										serie.XSeries = this.Package.FormulaManager.UpdateFormulaReferences(serie.XSeries, rows, columns, rowFrom, colFrom, excelAddress.WorkSheet, this.Name);
									}
									if (serie is ExcelBubbleChartSerie bubbleSerie && !string.IsNullOrEmpty(bubbleSerie.BubbleSize))
									{
										var excelAddress = new ExcelAddress(bubbleSerie.BubbleSize);
										bubbleSerie.BubbleSize = this.Package.FormulaManager.UpdateFormulaReferences(bubbleSerie.BubbleSize, rows, columns, rowFrom, colFrom, excelAddress.WorkSheet, this.Name);
									}
								}
							}
						}
					}
				}
			}
		}

		private static void InsertTableColumns(int columnFrom, int columns, ExcelTable tbl)
		{
			var node = tbl.Columns[0].TopNode.ParentNode;
			var ix = columnFrom - tbl.Address.Start.Column - 1;
			var insPos = node.ChildNodes[ix];
			ix += 2;
			for (int i = 0; i < columns; i++)
			{
				var name =
					tbl.Columns.GetUniqueName(string.Format("Column{0}",
						(ix++).ToString(CultureInfo.InvariantCulture)));
				XmlElement tableColumn =
					(XmlElement)tbl.TableXml.CreateNode(XmlNodeType.Element, "tableColumn", ExcelPackage.schemaMain);
				tableColumn.SetAttribute("id", (tbl.Columns.Count + i + 1).ToString(CultureInfo.InvariantCulture));
				tableColumn.SetAttribute("name", name);
				insPos = node.InsertAfter(tableColumn, insPos);
			} //Create tbl Column
			tbl._cols = new ExcelTableColumnCollection(tbl);
		}

		/// <summary>
		/// Adds a value to the row of merged cells to fix for inserts or deletes
		/// </summary>
		/// <param name="row"></param>
		/// <param name="rows"></param>
		/// <param name="delete"></param>
		private void FixMergedCellsRow(int row, int rows, bool delete)
		{
			if (delete)
			{
				this.MergedCells.Cells.Delete(row, 0, rows, 0);
			}
			else
			{
				this.MergedCells.Cells.Insert(row, 0, rows, 0);
			}

			List<int> removeIndex = new List<int>();
			for (int i = 0; i < _mergedCells.Count; i++)
			{
				if (!string.IsNullOrEmpty(_mergedCells[i]))
				{
					ExcelAddress addr = new ExcelAddress(_mergedCells[i]), newAddr;
					if (delete)
					{
						newAddr = addr.DeleteRow(row, rows);
						if (newAddr == null)
						{
							removeIndex.Add(i);
							continue;
						}
					}
					else
					{
						newAddr = addr.AddRow(row, rows);
						if (newAddr.Address != addr.Address)
						{
							this.MergedCells.SetIndex(newAddr, i);
						}
					}

					if (newAddr.Address != addr.Address)
					{
						this.MergedCells.List[i] = newAddr._address;
					}
				}
			}
			for (int i = removeIndex.Count - 1; i >= 0; i--)
			{
				this.MergedCells.List.RemoveAt(removeIndex[i]);
			}
		}

		/// <summary>
		/// Adds a value to the row of merged cells to fix for inserts or deletes
		/// </summary>
		/// <param name="column"></param>
		/// <param name="columns"></param>
		/// <param name="delete"></param>
		private void FixMergedCellsColumn(int column, int columns, bool delete)
		{
			if (delete)
			{
				this.MergedCells.Cells.Delete(0, column, 0, columns);
			}
			else
			{
				this.MergedCells.Cells.Insert(0, column, 0, columns);
			}
			List<int> removeIndex = new List<int>();
			for (int i = 0; i < _mergedCells.Count; i++)
			{
				if (!string.IsNullOrEmpty(_mergedCells[i]))
				{
					ExcelAddress addr = new ExcelAddress(_mergedCells[i]), newAddr;
					if (delete)
					{
						newAddr = addr.DeleteColumn(column, columns);
						if (newAddr == null)
						{
							removeIndex.Add(i);
							continue;
						}
					}
					else
					{
						newAddr = addr.AddColumn(column, columns);
						if (newAddr.Address != addr.Address)
						{
							this.MergedCells.SetIndex(newAddr, i);
						}
					}

					if (newAddr.Address != addr.Address)
					{
						this.MergedCells.List[i] = newAddr._address;
					}
				}
			}
			for (int i = removeIndex.Count - 1; i >= 0; i--)
			{
				this.MergedCells.List.RemoveAt(removeIndex[i]);
			}
		}

		private void FixSharedFormulasRows(int position, int rows)
		{
			List<Formulas> added = new List<Formulas>();
			List<Formulas> deleted = new List<Formulas>();

			foreach (int id in this._sharedFormulas.Keys)
			{
				var f = this._sharedFormulas[id];
				int fromCol, fromRow, toCol, toRow;

				ExcelCellBase.GetRowColFromAddress(f.Address, out fromRow, out fromCol, out toRow, out toCol);
				if (position >= fromRow && position + (Math.Abs(rows)) <= toRow) //Insert/delete is whithin the share formula address
				{
					if (rows > 0) //Insert
					{
						f.Address = ExcelCellBase.GetAddress(fromRow, fromCol) + ":" + ExcelCellBase.GetAddress(position - 1, toCol);
						if (toRow != fromRow)
						{
							Formulas newF = new Formulas(SourceCodeTokenizer.Default);
							newF.StartCol = f.StartCol;
							newF.StartRow = position + rows;
							newF.Address = ExcelCellBase.GetAddress(position + rows, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
							newF.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), position, f.StartCol);
							added.Add(newF);
						}
					}
					else
					{
						if (fromRow - rows < toRow)
						{
							f.Address = ExcelCellBase.GetAddress(fromRow, fromCol, toRow + rows, toCol);
						}
						else
						{
							f.Address = ExcelCellBase.GetAddress(fromRow, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
						}
					}
				}
				else if (position <= toRow)
				{
					if (rows > 0) //Insert before shift down
					{
						f.StartRow += rows;
						f.Address = ExcelCellBase.GetAddress(fromRow + rows, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
					}
					else
					{
						if (position <= fromRow && position + Math.Abs(rows) > toRow)  //Delete the formula
						{
							deleted.Add(f);
						}
						else
						{
							toRow = toRow + rows < position - 1 ? position - 1 : toRow + rows;
							if (position <= fromRow)
							{
								fromRow = fromRow + rows < position ? position : fromRow + rows;
							}

							f.Address = ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol);
							this.Cells[f.Address].SetSharedFormulaID(f.Index);
						}
					}
				}
			}

			this.AddFormulas(added, position, rows);

			//Remove formulas
			foreach (Formulas f in deleted)
			{
				this._sharedFormulas.Remove(f.Index);
			}

			//Fix Formulas
			added = new List<Formulas>();
			foreach (int id in this._sharedFormulas.Keys)
			{
				var f = this._sharedFormulas[id];
				this.UpdateSharedFormulaRow(ref f, position, rows, ref added);
			}
			this.AddFormulas(added, position, rows);
		}

		private void AddFormulas(List<Formulas> added, int position, int rows)
		{
			//Add new formulas
			foreach (Formulas f in added)
			{
				f.Index = GetMaxShareFunctionIndex(false);
				this._sharedFormulas.Add(f.Index, f);
				this.Cells[f.Address].SetSharedFormulaID(f.Index);
			}
		}

		private void UpdateSharedFormulaRow(ref Formulas formula, int startRow, int rows, ref List<Formulas> newFormulas)
		{
			int fromRow, fromCol, toRow, toCol;
			int newFormulasCount = newFormulas.Count;
			ExcelCellBase.GetRowColFromAddress(formula.Address, out fromRow, out fromCol, out toRow, out toCol);
			//int refSplits = Regex.Split(formula.Formula, "#REF!").GetUpperBound(0);
			string formualR1C1;
			if (rows > 0 || fromRow <= startRow)
			{
				formualR1C1 = ExcelRangeBase.TranslateToR1C1(formula.Formula, formula.StartRow, formula.StartCol);
				formula.Formula = ExcelRangeBase.TranslateFromR1C1(formualR1C1, fromRow, formula.StartCol);
			}
			else
			{
				formualR1C1 = ExcelRangeBase.TranslateToR1C1(formula.Formula, formula.StartRow - rows, formula.StartCol);
				formula.Formula = ExcelRangeBase.TranslateFromR1C1(formualR1C1, formula.StartRow, formula.StartCol);
			}
			string prevFormualR1C1 = formualR1C1;
			for (int row = fromRow; row <= toRow; row++)
			{
				for (int col = fromCol; col <= toCol; col++)
				{
					string newFormula;
					string currentFormulaR1C1;
					if (rows > 0 || row < startRow)
					{
						newFormula = this.Package.FormulaManager.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row, col), rows, 0, startRow, 0, this.Name, this.Name);
						currentFormulaR1C1 = ExcelRangeBase.TranslateToR1C1(newFormula, row, col);
					}
					else
					{
						newFormula = this.Package.FormulaManager.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row - rows, col), rows, 0, startRow, 0, this.Name, this.Name);
						currentFormulaR1C1 = ExcelRangeBase.TranslateToR1C1(newFormula, row, col);
					}
					if (currentFormulaR1C1 != prevFormualR1C1)
					{
						if (row == fromRow && col == fromCol)
						{
							formula.Formula = newFormula;
						}
						else
						{
							if (newFormulas.Count == newFormulasCount)
							{
								formula.Address = ExcelCellBase.GetAddress(formula.StartRow, formula.StartCol, row - 1, col);
							}
							else
							{
								newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[newFormulas.Count - 1].StartCol, row - 1, col);
							}
							var refFormula = new Formulas(SourceCodeTokenizer.Default);
							refFormula.Formula = newFormula;
							refFormula.StartRow = row;
							refFormula.StartCol = col;
							newFormulas.Add(refFormula);
							prevFormualR1C1 = currentFormulaR1C1;
						}
					}
				}
			}
			if (rows < 0 && formula.StartRow > startRow)
			{
				if (formula.StartRow + rows < startRow)
				{
					formula.StartRow = startRow;
				}
				else
				{
					formula.StartRow += rows;
				}
			}
			if (newFormulas.Count > newFormulasCount)
			{
				newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[newFormulas.Count - 1].StartCol, toRow, toCol);
			}
		}

		private void ChangeNames(string value)
		{
			//Renames name in this Worksheet;
			foreach (var namedRange in this.Workbook.Names)
			{
				namedRange.NameFormula = this.Package.FormulaManager.UpdateFormulaSheetReferences(namedRange.NameFormula, this._name, value);
			}
			this.ChangeSparklineSheetNames(value);
			foreach (var ws in this.Workbook.Worksheets)
			{
				if (!(ws is ExcelChartsheet))
				{
					foreach (var namedRange in ws.Names)
					{
						namedRange.NameFormula = this.Package.FormulaManager.UpdateFormulaSheetReferences(namedRange.NameFormula, this._name, value);
					}
					ws.UpdateCrossSheetReferenceNames(_name, value);
				}
				HashSet<ExcelChart> charts = new HashSet<ExcelChart>();
				foreach (ExcelChart chartBase in ws.Drawings.Where(drawing => drawing is ExcelChart))
				{
					foreach (var chart in chartBase.PlotArea.ChartTypes)
					{
						bool isUnique = charts.Add(chart);
						if (isUnique)
						{
							foreach (ExcelChartSerie serie in chart.Series)
							{
								if (!string.IsNullOrEmpty(serie.HeaderAddress?.Address))
								{
									string updatedHeaderAddress = this.Package.FormulaManager.UpdateFormulaSheetReferences(serie.HeaderAddress.Address, this._name, value);
									serie.HeaderAddress = new ExcelAddress(updatedHeaderAddress);
								}
								if (!string.IsNullOrEmpty(serie.Series))
									serie.Series = this.Package.FormulaManager.UpdateFormulaSheetReferences(serie.Series, this._name, value);
								if (!string.IsNullOrEmpty(serie.XSeries))
									serie.XSeries = this.Package.FormulaManager.UpdateFormulaSheetReferences(serie.XSeries, this._name, value);
								if (serie is ExcelBubbleChartSerie bubbleSerie && !string.IsNullOrEmpty(bubbleSerie.BubbleSize))
									bubbleSerie.BubbleSize = this.Package.FormulaManager.UpdateFormulaSheetReferences(bubbleSerie.BubbleSize, this._name, value);
							}
						}
					}
				}
			}
		}

		private double GetRowHeightFromNormalStyle()
		{
			var ix = Workbook.Styles.NamedStyles.FindIndexByID("Normal");
			if (ix >= 0)
			{
				var f = Workbook.Styles.NamedStyles[ix].Style.Font;
				return ExcelFontXml.GetFontHeight(f.Name, f.Size) * 0.75;
			}
			else
			{
				return 15;   //Default Calibri 11
			}
		}

		private void CreateXml()
		{
			this._worksheetXml = new XmlDocument();
			this._worksheetXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
			Packaging.ZipPackagePart packPart = this.Package.Package.GetPart(WorksheetUri);
			string xml = "";

			// First Columns, rows, cells, mergecells, hyperlinks and pagebreakes are loaded from a xmlstream to optimize speed...
			bool doAdjust = this.Package.DoAdjustDrawings;
			this.Package.DoAdjustDrawings = false;
			Stream stream = packPart.GetStream();

			XmlTextReader xr = new XmlTextReader(stream);
			xr.DtdProcessing = DtdProcessing.Prohibit;
			xr.WhitespaceHandling = WhitespaceHandling.None;
			this.LoadColumns(xr);    //columnXml
			long start = stream.Position;
			this.LoadCells(xr);
			var nextElementLength = GetAttributeLength(xr);
			long end = stream.Position - nextElementLength;
			this.LoadMergeCells(xr);
			this.LoadHyperLinks(xr);
			this.LoadRowPageBreakes(xr);
			this.LoadColPageBreaks(xr);
			//...then the rest of the Xml is extracted and loaded into the WorksheetXml document.
			stream.Seek(0, SeekOrigin.Begin);
			Encoding encoding;
			xml = this.GetWorkSheetXml(stream, start, end, out encoding);

			// now release stream buffer (already converted whole Xml into XmlDocument Object and String)
			stream.Close();
			stream.Dispose();
			packPart.Stream = new MemoryStream();

			//first char is invalid sometimes??
			if (xml[0] != '<')
				XmlHelper.LoadXmlSafe(_worksheetXml, xml.Substring(1, xml.Length - 1), encoding);
			else
				XmlHelper.LoadXmlSafe(_worksheetXml, xml, encoding);

			this.Package.DoAdjustDrawings = doAdjust;
			this.ClearNodes();
		}

		/// <summary>
		/// Get the lenth of the attributes
		/// Conditional formatting attributes can be extremly long som get length of the attributes to finetune position.
		/// </summary>
		/// <param name="xr"></param>
		/// <returns></returns>
		private int GetAttributeLength(XmlTextReader xr)
		{
			if (xr.NodeType != XmlNodeType.Element) return 0;
			var length = 0;

			for (int i = 0; i < xr.AttributeCount; i++)
			{
				var a = xr.GetAttribute(i);
				length += string.IsNullOrEmpty(a) ? 0 : a.Length;
			}
			return length;
		}

		private void LoadRowPageBreakes(XmlTextReader xr)
		{
			if (!this.ReadUntil(xr, "rowBreaks", "colBreaks")) return;
			while (xr.Read())
			{
				if (xr.LocalName == "brk")
				{
					if (xr.NodeType == XmlNodeType.Element)
					{
						int id;
						if (int.TryParse(xr.GetAttribute("id"), out id))
						{
							this.Row(id).PageBreak = true;
						}
					}
				}
				else
				{
					break;
				}
			}
		}

		private void LoadColPageBreaks(XmlTextReader xr)
		{
			if (!this.ReadUntil(xr, "colBreaks")) return;
			while (xr.Read())
			{
				if (xr.LocalName == "brk")
				{
					if (xr.NodeType == XmlNodeType.Element)
					{
						int id;
						if (int.TryParse(xr.GetAttribute("id"), out id))
						{
							this.Column(id).PageBreak = true;
						}
					}
				}
				else
				{
					break;
				}
			}
		}

		private void DisposeInternal(IDisposable candidateDisposable)
		{
			if (candidateDisposable != null)
			{
				candidateDisposable.Dispose();
			}
		}

		private void ClearNodes()
		{
			this.WorksheetXml.SelectSingleNode("//d:cols", NameSpaceManager)?.RemoveAll();
			this.WorksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager)?.RemoveAll();
			this.WorksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager)?.RemoveAll();
			this.WorksheetXml.SelectSingleNode("//d:rowBreaks", NameSpaceManager)?.RemoveAll();
			this.WorksheetXml.SelectSingleNode("//d:colBreaks", NameSpaceManager)?.RemoveAll();
		}

		/// <summary>
		/// Extracts the workbook XML without the sheetData-element (containing all cell data).
		/// Xml-Cell data can be extreemly large (GB), so we find the sheetdata element in the streem (position start) and
		/// then tries to find the &lt;/sheetData&gt; element from the end-parameter.
		/// This approach is to avoid out of memory exceptions reading large packages
		/// </summary>
		/// <param name="stream">the worksheet stream</param>
		/// <param name="start">Position from previous reading where we found the sheetData element</param>
		/// <param name="end">End position, where &lt;/sheetData&gt; or &lt;sheetData/&gt; is found</param>
		/// <param name="encoding">Encoding</param>
		/// <returns>The worksheet xml, with an empty sheetdata. (Sheetdata is in memory in the worksheet)</returns>
		private string GetWorkSheetXml(Stream stream, long start, long end, out Encoding encoding)
		{
			StreamReader sr = new StreamReader(stream);
			int length = 0;
			char[] block;
			int pos;
			StringBuilder sb = new StringBuilder();
			Match startmMatch, endMatch;
			do
			{
				int size = stream.Length < BLOCKSIZE ? (int)stream.Length : BLOCKSIZE;
				block = new char[size];
				pos = sr.ReadBlock(block, 0, size);
				sb.Append(block, 0, pos);
				length += size;
			}
			while (length < start + 20 && length < end);    //the  start-pos contains the stream position of the sheetData element. Add 20 (with some safty for whitespace, streampointer diff etc, just so be sure).
			startmMatch = Regex.Match(sb.ToString(), string.Format("(<[^>]*{0}[^>]*>)", "sheetData"));
			if (!startmMatch.Success) //Not found
			{
				encoding = sr.CurrentEncoding;
				return sb.ToString();
			}
			else
			{
				string s = sb.ToString();
				string xml = s.Substring(0, startmMatch.Index);
				if (Utils.ConvertUtil._invariantCompareInfo.IsSuffix(startmMatch.Value, "/>"))        //Empty sheetdata
				{
					xml += s.Substring(startmMatch.Index, s.Length - startmMatch.Index);
				}
				else
				{
					if (sr.Peek() != -1)        //Now find the end tag </sheetdata> so we can add the end of the xml document
					{
						/**** Fixes issue 14788. Fix by Philip Garrett ****/
						long endSeekStart = end;

						while (endSeekStart >= 0)
						{
							endSeekStart = Math.Max(endSeekStart - BLOCKSIZE, 0);
							int size = (int)(end - endSeekStart);
							stream.Seek(endSeekStart, SeekOrigin.Begin);
							block = new char[size];
							sr = new StreamReader(stream);
							pos = sr.ReadBlock(block, 0, size);
							sb = new StringBuilder();
							sb.Append(block, 0, pos);
							s = sb.ToString();
							endMatch = Regex.Match(s, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
							if (endMatch.Success)
							{
								break;
							}
						}
					}
					endMatch = Regex.Match(s, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
					xml += "<sheetData/>" + s.Substring(endMatch.Index + endMatch.Length, s.Length - (endMatch.Index + endMatch.Length));
				}
				if (sr.Peek() > -1)
				{
					xml += sr.ReadToEnd();
				}

				encoding = sr.CurrentEncoding;
				return xml;
			}
		}

		private void GetBlockPos(string xml, string tag, ref int start, ref int end)
		{
			Match startmMatch, endMatch;
			startmMatch = Regex.Match(xml.Substring(start), string.Format("(<[^>]*{0}[^>]*>)", tag)); //"<[a-zA-Z:]*" + tag + "[?]*>");

			if (!startmMatch.Success) //Not found
			{
				start = -1;
				end = -1;
				return;
			}
			var startPos = startmMatch.Index + start;
			if (startmMatch.Value.Substring(startmMatch.Value.Length - 2, 1) == "/")
			{
				end = startPos + startmMatch.Length;
			}
			else
			{
				endMatch = Regex.Match(xml.Substring(start), string.Format("(</[^>]*{0}[^>]*>)", tag));
				if (endMatch.Success)
				{
					end = endMatch.Index + endMatch.Length + start;
				}
			}
			start = startPos;
		}

		private bool ReadUntil(XmlTextReader xr, params string[] tagName)
		{
			if (xr.EOF) return false;
			while (!Array.Exists(tagName, tag => Utils.ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag)))
			{
				xr.Read();
				if (xr.EOF) return false;
			}
			return (Utils.ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]));
		}

		private void LoadColumns(XmlTextReader xr)
		{
			var colList = new List<IRangeID>();
			if (this.ReadUntil(xr, "cols", "sheetData"))
			{
				while (xr.Read())
				{
					if (xr.NodeType == XmlNodeType.Whitespace) continue;
					if (xr.LocalName != "col") break;
					if (xr.NodeType == XmlNodeType.Element)
					{
						int min = int.Parse(xr.GetAttribute("min"));

						ExcelColumn col = new ExcelColumn(this, min);

						col.ColumnMax = int.Parse(xr.GetAttribute("max"));
						col.Width = xr.GetAttribute("width") == null ? 0 : double.Parse(xr.GetAttribute("width"), CultureInfo.InvariantCulture);
						col.BestFit = xr.GetAttribute("bestFit") != null && xr.GetAttribute("bestFit") == "1" ? true : false;
						col.Collapsed = xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1" ? true : false;
						col.Phonetic = xr.GetAttribute("phonetic") != null && xr.GetAttribute("phonetic") == "1" ? true : false;
						col.OutlineLevel = (short)(xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture));
						col.Hidden = xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false;
						this.SetValueInner(0, min, col);

						int style;
						if (!(xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), out style)))
						{
							this.SetStyleInner(0, min, style);
						}
					}
				}
			}
		}

		/// <summary>
		/// Read until the node is found. If not found the xmlreader is reseted.
		/// </summary>
		/// <param name="xr">The reader</param>
		/// <param name="nodeText">Text to search for</param>
		/// <param name="altNode">Alternative text to search for</param>
		/// <returns></returns>
		private static bool ReadXmlReaderUntil(XmlTextReader xr, string nodeText, string altNode)
		{
			do
			{
				if (xr.LocalName == nodeText || xr.LocalName == altNode) return true;
			}
			while (xr.Read());
			xr.Close();
			return false;
		}

		/// <summary>
		/// Load Hyperlinks
		/// </summary>
		/// <param name="xr">The reader</param>
		private void LoadHyperLinks(XmlTextReader xr)
		{
			if (!this.ReadUntil(xr, "hyperlinks", "rowBreaks", "colBreaks")) return;
			while (xr.Read())
			{
				if (xr.LocalName == "hyperlink")
				{
					int fromRow, fromCol, toRow, toCol;
					ExcelCellBase.GetRowColFromAddress(xr.GetAttribute("ref"), out fromRow, out fromCol, out toRow, out toCol);
					ExcelHyperLink hl = null;
					if (xr.GetAttribute("id", ExcelPackage.schemaRelationships) != null)
					{
						var rId = xr.GetAttribute("id", ExcelPackage.schemaRelationships);
						var uri = Part.GetRelationship(rId).TargetUri;
						if (uri.IsAbsoluteUri)
						{
							try
							{
								hl = new ExcelHyperLink(uri.AbsoluteUri);
							}
							catch
							{
								hl = new ExcelHyperLink(uri.OriginalString, UriKind.Absolute);
							}
						}
						else
						{
							hl = new ExcelHyperLink(uri.OriginalString, UriKind.Relative);
						}
						hl.RId = rId;
						this.Part.DeleteRelationship(rId); //Delete the relationship, it is recreated when we save the package.
					}
					else if (xr.GetAttribute("location") != null)
					{
						hl = new ExcelHyperLink(xr.GetAttribute("location"), xr.GetAttribute("display"));
						hl.RowSpann = toRow - fromRow;
						hl.ColSpann = toCol - fromCol;
					}

					string tt = xr.GetAttribute("tooltip");
					if (!string.IsNullOrEmpty(tt))
					{
						hl.ToolTip = tt;
					}
					this._hyperLinks.SetValue(fromRow, fromCol, hl);
				}
				else
				{
					break;
				}
			}
		}

		/// <summary>
		/// Load cells
		/// </summary>
		/// <param name="xr">The reader</param>
		private void LoadCells(XmlTextReader xr)
		{
			this.ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
			ExcelAddress address = null;
			string type = "";
			int style = 0;
			int row = 0;
			int col = 0;
			xr.Read();

			while (!xr.EOF)
			{
				while (xr.NodeType == XmlNodeType.EndElement)
				{
					xr.Read();
					continue;
				}
				if (xr.LocalName == "row")
				{
					var r = xr.GetAttribute("r");
					if (r == null)
					{
						row++;
					}
					else
					{
						row = Convert.ToInt32(r);
					}

					if (this.DoAddRow(xr))
					{
						this.SetValueInner(row, 0, AddRow(xr, row));
						if (xr.GetAttribute("s") != null)
						{
							this.SetStyleInner(row, 0, int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture));
						}
					}
					xr.Read();
				}
				else if (xr.LocalName == "c")
				{
					var r = xr.GetAttribute("r");
					if (r == null)
					{
						//Handle cells with no reference
						col++;
						address = new ExcelAddress(row, col, row, col);
					}
					else
					{
						address = new ExcelAddress(r);
						col = address._fromCol;
					}

					//Datetype
					if (xr.GetAttribute("t") != null)
					{
						type = xr.GetAttribute("t");
					}
					else
					{
						type = "";
					}
					//Style
					if (xr.GetAttribute("s") != null)
					{
						style = int.Parse(xr.GetAttribute("s"));
						SetStyleInner(address._fromRow, address._fromCol, style);
						//SetValueInner(address._fromRow, address._fromCol, null); //TODO:Better Performance ??
					}
					else
					{
						style = 0;
					}
					xr.Read();
				}
				else if (xr.LocalName == "v")
				{
					this.SetValueFromXml(xr, type, style, address._fromRow, address._fromCol);
					xr.Read();
				}
				else if (xr.LocalName == "f")
				{
					string t = xr.GetAttribute("t");
					if (t == null)
					{
						this._formulas.SetValue(address._fromRow, address._fromCol, xr.ReadElementContentAsString());
						this.SetValueInner(address._fromRow, address._fromCol, null);
					}
					else if (t == "shared")
					{

						string si = xr.GetAttribute("si");
						if (si != null)
						{
							var sfIndex = int.Parse(si);
							_formulas.SetValue(address._fromRow, address._fromCol, sfIndex);
							this.SetValueInner(address._fromRow, address._fromCol, null);
							string fAddress = xr.GetAttribute("ref");
							string formula = ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString());
							if (formula != "")
							{
								this._sharedFormulas.Add(sfIndex, new Formulas(SourceCodeTokenizer.Default) { Index = sfIndex, Formula = formula, Address = fAddress, StartRow = address._fromRow, StartCol = address._fromCol });
							}
						}
						else
						{
							xr.Read();  //Something is wrong in the sheet, read next
						}
					}
					else if (t == "array") //TODO: Array functions are not support yet. Read the formula for the start cell only.
					{
						string aAddress = xr.GetAttribute("ref");
						string formula = xr.ReadElementContentAsString();
						var afIndex = GetMaxShareFunctionIndex(true);
						this._formulas.SetValue(address._fromRow, address._fromCol, afIndex);
						this.SetValueInner(address._fromRow, address._fromCol, null);
						this._sharedFormulas.Add(afIndex, new Formulas(SourceCodeTokenizer.Default) { Index = afIndex, Formula = formula, Address = aAddress, StartRow = address._fromRow, StartCol = address._fromCol, IsArray = true });
					}
					else // ??? some other type
					{
						xr.Read();  //Something is wrong in the sheet, read next
					}

				}
				else if (xr.LocalName == "is")   //Inline string
				{
					xr.Read();
					if (xr.LocalName == "t")
					{
						this.SetValueInner(address._fromRow, address._fromCol, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
					}
					else
					{
						if (xr.LocalName == "r")
						{
							var rXml = xr.ReadOuterXml();
							while (xr.LocalName == "r")
							{
								rXml += xr.ReadOuterXml();
							}
							this.SetValueInner(address._fromRow, address._fromCol, rXml);
						}
						else
						{
							this.SetValueInner(address._fromRow, address._fromCol, xr.ReadOuterXml());
						}
						this._flags.SetFlagValue(address._fromRow, address._fromCol, true, CellFlags.RichText);
					}
				}
				else
				{
					break;
				}
			}
		}

		private bool DoAddRow(XmlTextReader xr)
		{
			var c = xr.GetAttribute("r") == null ? 0 : 1;
			if (xr.GetAttribute("spans") != null)
			{
				c++;
			}
			return xr.AttributeCount > c;
		}

		/// <summary>
		/// Load merged cells
		/// </summary>
		/// <param name="xr"></param>
		private void LoadMergeCells(XmlTextReader xr)
		{
			if (this.ReadUntil(xr, "mergeCells", "hyperlinks", "rowBreaks", "colBreaks") && !xr.EOF)
			{
				while (xr.Read())
				{
					if (xr.LocalName != "mergeCell") break;
					if (xr.NodeType == XmlNodeType.Element)
					{
						string address = xr.GetAttribute("ref");
						this.MergedCells.Add(new ExcelAddress(address), false);
					}
				}
			}
		}

		/// <summary>
		/// Update merged cells
		/// </summary>
		/// <param name="sw">The writer</param>
		private void UpdateMergedCells(StreamWriter sw)
		{
			sw.Write("<mergeCells>");
			foreach (string address in _mergedCells)
			{
				sw.Write("<mergeCell ref=\"{0}\" />", address);
			}
			sw.Write("</mergeCells>");
		}

		/// <summary>
		/// Reads a row from the XML reader
		/// </summary>
		/// <param name="xr">The reader</param>
		/// <param name="row">The row number</param>
		/// <returns></returns>
		private RowInternal AddRow(XmlTextReader xr, int row)
		{
			return new RowInternal()
			{
				Collapsed = (xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1" ? true : false),
				OutlineLevel = (xr.GetAttribute("outlineLevel") == null ? (short)0 : short.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture)),
				Height = (xr.GetAttribute("ht") == null ? -1 : double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture)),
				Hidden = (xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false),
				Phonetic = xr.GetAttribute("ph") != null && xr.GetAttribute("ph") == "1" ? true : false,
				CustomHeight = xr.GetAttribute("customHeight") == null ? false : xr.GetAttribute("customHeight") == "1"
			};
		}

		private void SetValueFromXml(XmlTextReader xr, string type, int styleID, int row, int col)
		{
			if (type == "s")
			{
				int ix = xr.ReadElementContentAsInt();
				this.SetValueInner(row, col, this.Package.Workbook.SharedStringsList[ix].Text);
				if (this.Package.Workbook.SharedStringsList[ix].isRichText)
				{
					this._flags.SetFlagValue(row, col, true, CellFlags.RichText);
				}
			}
			else if (type == "str")
			{
				this.SetValueInner(row, col, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
			}
			else if (type == "b")
			{
				this.SetValueInner(row, col, (xr.ReadElementContentAsString() != "0"));
			}
			else if (type == "e")
			{
				this.SetValueInner(row, col, GetErrorType(xr.ReadElementContentAsString()));
			}
			else
			{
				string v = xr.ReadElementContentAsString();
				var nf = Workbook.Styles.CellXfs[styleID].NumberFormatId;
				if ((nf >= 14 && nf <= 22) || (nf >= 45 && nf <= 47))
				{
					double res;
					if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out res))
					{
						if (Workbook.Date1904)
						{
							res += ExcelWorkbook.date1904Offset;
						}
						if (res >= -657435.0 && res < 2958465.9999999)
						{
							this.SetValueInner(row, col, DateTime.FromOADate(res));
						}
						else
						{
							this.SetValueInner(row, col, res);
						}
					}
					else
					{
						this.SetValueInner(row, col, v);
					}
				}
				else
				{
					double d;
					if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
					{
						this.SetValueInner(row, col, d);
					}
					else
					{
						this.SetValueInner(row, col, double.NaN);
					}
				}
			}
		}

		private object GetErrorType(string v)
		{
			return ExcelErrorValue.Parse(ConvertUtil._invariantTextInfo.ToUpper(v));
		}

		private void UpdateCrossSheetReferences(string sheetWhoseReferencesShouldBeUpdated, int rowFrom, int rows, int columnFrom, int columns)
		{
			lock (this)
			{
				foreach (var f in _sharedFormulas.Values)
				{
					f.Formula = this.Package.FormulaManager.UpdateFormulaReferences(f.Formula, rows, columns, rowFrom, columnFrom, this.Name, sheetWhoseReferencesShouldBeUpdated);
				}
				var cse = _formulas.GetEnumerator();
				while (cse.MoveNext())
				{
					if (cse.Value is string)
					{
						cse.Value = this.Package.FormulaManager.UpdateFormulaReferences(cse.Value.ToString(), rows, columns, rowFrom, columnFrom, this.Name, sheetWhoseReferencesShouldBeUpdated);
					}
				}
			}
		}

		private void UpdateCrossSheetReferenceNames(string oldName, string newName)
		{
			if (string.IsNullOrEmpty(oldName))
				throw new ArgumentNullException(nameof(oldName));
			if (string.IsNullOrEmpty(newName))
				throw new ArgumentNullException(nameof(newName));
			lock (this)
			{
				foreach (var f in _sharedFormulas.Values)
				{
					f.Formula = this.Package.FormulaManager.UpdateFormulaSheetReferences(f.Formula, oldName, newName);
				}
				var cse = _formulas.GetEnumerator();
				while (cse.MoveNext())
				{
					if (cse.Value is string)
					{
						cse.Value = this.Package.FormulaManager.UpdateFormulaSheetReferences(cse.Value.ToString(), oldName, newName);
					}
				}
			}
		}

		/// <summary>
		/// Delete the printersettings relationship and part.
		/// </summary>
		private void DeletePrinterSettings()
		{
			//Delete the relationship from the pageSetup tag
			XmlAttribute attr = (XmlAttribute)WorksheetXml.SelectSingleNode("//d:pageSetup/@r:id", NameSpaceManager);
			if (attr != null)
			{
				string relID = attr.Value;
				//First delete the attribute from the XML
				attr.OwnerElement.Attributes.Remove(attr);
				if (Part.RelationshipExists(relID))
				{
					var rel = Part.GetRelationship(relID);
					Uri printerSettingsUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
					Part.DeleteRelationship(rel.Id);

					//Delete the part from the package
					if (this.Package.Package.PartExists(printerSettingsUri))
					{
						this.Package.Package.DeletePart(printerSettingsUri);
					}
				}
			}
		}

		private void SaveComments()
		{
			if (_comments != null)
			{
				if (_comments.Count == 0)
				{
					if (_comments.Uri != null)
					{
						this.Part.DeleteRelationship(_comments.RelId);
						this.Package.Package.DeletePart(_comments.Uri);
					}
				}
				else
				{
					if (_comments.Uri == null)
						_comments.Uri = new Uri(string.Format(@"/xl/comments{0}.xml", this.SheetID), UriKind.Relative);
					if (_comments.Part == null)
					{
						_comments.Part = this.Package.Package.CreatePart(_comments.Uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", this.Package.Compression);
						this.Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _comments.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");
					}
					_comments.CommentXml.Save(_comments.Part.GetStream(FileMode.Create));
					ExcelVmlDrawingCommentHelper.AddCommentDrawings(this, _comments);
				}
			}
		}

		/// <summary>
		/// Save all table data
		/// </summary>
		private void SaveTables()
		{
			if (this.Tables.Count == 0)
			{
				XmlNode tbls = TopNode.SelectSingleNode("d:tableParts", NameSpaceManager);
				if (tbls != null)
					tbls.ParentNode.RemoveChild(tbls);
			}
			foreach (var tbl in Tables)
			{
				if (tbl.ShowHeader || tbl.ShowTotal)
				{
					int colNum = tbl.Address._fromCol;
					var colVal = new HashSet<string>();
					foreach (var col in tbl.Columns)
					{
						string n = col.Name.ToLower(CultureInfo.InvariantCulture);
						if (tbl.ShowHeader)
						{
							n = tbl.WorkSheet.GetValue<string>(tbl.Address._fromRow,
								tbl.Address._fromCol + col.Position);
							if (string.IsNullOrEmpty(n))
							{
								n = col.Name.ToLower(CultureInfo.InvariantCulture);
								SetValueInner(tbl.Address._fromRow, colNum, ConvertUtil.ExcelDecodeString(col.Name));
							}
							else
							{
								col.Name = n;
							}
						}
						else
						{
							n = col.Name.ToLower(CultureInfo.InvariantCulture);
						}

						if (colVal.Contains(n))
						{
							throw (new InvalidDataException(string.Format("Table {0} Column {1} does not have a unique name.", tbl.Name, col.Name)));
						}
						colVal.Add(n);
						if (!string.IsNullOrEmpty(col.CalculatedColumnFormula))
						{
							int fromRow = tbl.ShowHeader ? tbl.Address._fromRow + 1 : tbl.Address._fromRow;
							int toRow = tbl.ShowTotal ? tbl.Address._toRow - 1 : tbl.Address._toRow;
							for (int row = fromRow; row <= toRow; row++)
							{
								this.SetFormula(row, colNum, col.CalculatedColumnFormula);
							}
						}
						colNum++;
					}
				}
				if (tbl.Part == null)
				{
					var id = tbl.Id;
					tbl.TableUri = GetNewUri(this.Package.Package, @"/xl/tables/table{0}.xml", ref id);
					tbl.Id = id;
					tbl.Part = this.Package.Package.CreatePart(tbl.TableUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", Workbook.Package.Compression);
					var stream = tbl.Part.GetStream(FileMode.Create);
					tbl.TableXml.Save(stream);
					var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, tbl.TableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");
					tbl.RelationshipID = rel.Id;

					this.CreateNode("d:tableParts");
					XmlNode tbls = TopNode.SelectSingleNode("d:tableParts", NameSpaceManager);
					var tblNode = tbls.OwnerDocument.CreateElement("tablePart", ExcelPackage.schemaMain);
					tbls.AppendChild(tblNode);
					tblNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
				}
				else
				{
					var stream = tbl.Part.GetStream(FileMode.Create);
					tbl.TableXml.Save(stream);
				}
			}
		}

		private void SavePivotTables()
		{
			foreach (var pt in PivotTables)
			{
				pt.SetXmlNodeString("d:location/@ref", pt.Address.FirstAddress);
				if (pt.DataFields.Count > 1)
				{
					XmlElement parentNode;
					if (pt.DataOnRows == true)
					{
						parentNode = pt.PivotTableXml.SelectSingleNode("//d:rowFields", pt.NameSpaceManager) as XmlElement;
						if (parentNode == null)
						{
							pt.CreateNode("d:rowFields");
							parentNode = pt.PivotTableXml.SelectSingleNode("//d:rowFields", pt.NameSpaceManager) as XmlElement;
						}
					}
					else
					{
						parentNode = pt.PivotTableXml.SelectSingleNode("//d:colFields", pt.NameSpaceManager) as XmlElement;
						if (parentNode == null)
						{
							pt.CreateNode("d:colFields");
							parentNode = pt.PivotTableXml.SelectSingleNode("//d:colFields", pt.NameSpaceManager) as XmlElement;
						}
					}

					if (parentNode.SelectSingleNode("d:field[@ x= \"-2\"]", pt.NameSpaceManager) == null)
					{
						XmlElement fieldNode = pt.PivotTableXml.CreateElement("field", ExcelPackage.schemaMain);
						fieldNode.SetAttribute("x", "-2");
						parentNode.AppendChild(fieldNode);
					}
				}

				//Rewrite the pivottable address again if any rows or columns have been inserted or deleted
				pt.SetXmlNodeString("d:location/@ref", pt.Address.Address);
				if (pt.CacheDefinition.GetSourceRangeAddress() != null && !pt.CacheDefinition.GetSourceRangeAddress().IsName)
				{
					string pivotTableReferenceSheet = pt.CacheDefinition.GetXmlNodeString(ExcelPivotCacheDefinition.SourceWorksheetPath);
					string address = string.IsNullOrEmpty(pivotTableReferenceSheet) ?
						pt.CacheDefinition.GetSourceRangeAddress().FullAddress :
						pt.CacheDefinition.GetSourceRangeAddress().Address;
					pt.CacheDefinition.SetXmlNodeString(ExcelPivotCacheDefinition.SourceAddressPath, address);
				}
				
				if (pt.CacheDefinition.GetSourceRangeAddress() != null)
				{
					foreach (var df in pt.DataFields)
					{
						if (string.IsNullOrEmpty(df.Name))
						{
							string name;
							if (df.Function == DataFieldFunctions.None)
							{
								name = df.Field.Name; //Name must be set or Excel will crash on rename.
							}
							else
							{
								name = df.Function.ToString() + " of " + df.Field.Name; //Name must be set or Excel will crash on rename.
							}
							//Make sure name is unique
							var newName = name;
							var i = 2;
							while (pt.DataFields.ExistsDfName(newName, df))
							{
								newName = name + (i++).ToString(CultureInfo.InvariantCulture);
							}
							df.Name = newName;
						}
					}
				}
				pt.PivotTableXml.Save(pt.Part.GetStream(FileMode.Create));
			}
		}

		private string GetNewName(HashSet<string> flds, string fldName)
		{
			int ix = 2;
			while (flds.Contains(fldName + ix.ToString(CultureInfo.InvariantCulture)))
			{
				ix++;
			}
			return fldName + ix.ToString(CultureInfo.InvariantCulture);
		}

		private static string GetTotalFunction(ExcelTableColumn col, string FunctionNum)
		{
			string escapedColumn = Regex.Replace(col.Name, @"[\[\]#']", new MatchEvaluator(m => "'" + m.Value));
			return string.Format("SUBTOTAL({0},{1}[{2}])", FunctionNum, col._tbl.Name, escapedColumn);
		}

		private void SaveXml(Stream stream)
		{
			//Create the nodes if they do not exist.
			StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.UTF8, 65536);
			if (this is ExcelChartsheet)
			{
				sw.Write(_worksheetXml.OuterXml);
			}
			else
			{
				CreateNode("d:cols");
				CreateNode("d:sheetData");
				CreateNode("d:mergeCells");
				CreateNode("d:hyperlinks");
				CreateNode("d:rowBreaks");
				CreateNode("d:colBreaks");

				//StreamWriter sw=new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
				var xml = _worksheetXml.OuterXml;
				int colStart = 0, colEnd = 0;
				GetBlockPos(xml, "cols", ref colStart, ref colEnd);

				sw.Write(xml.Substring(0, colStart));
				var colBreaks = new List<int>();
				//if (_columns.Count > 0)
				//{
				UpdateColumnData(sw);
				//}

				int cellStart = colEnd, cellEnd = colEnd;
				GetBlockPos(xml, "sheetData", ref cellStart, ref cellEnd);

				sw.Write(xml.Substring(colEnd, cellStart - colEnd));
				var rowBreaks = new List<int>();
				UpdateRowCellData(sw);

				int mergeStart = cellEnd, mergeEnd = cellEnd;

				GetBlockPos(xml, "mergeCells", ref mergeStart, ref mergeEnd);
				sw.Write(xml.Substring(cellEnd, mergeStart - cellEnd));

				CleanupMergedCells(_mergedCells);
				if (_mergedCells.Count > 0)
				{
					UpdateMergedCells(sw);
				}

				int hyperStart = mergeEnd, hyperEnd = mergeEnd;
				GetBlockPos(xml, "hyperlinks", ref hyperStart, ref hyperEnd);
				sw.Write(xml.Substring(mergeEnd, hyperStart - mergeEnd));
				//if (_hyperLinkCells.Count > 0)
				//{
				UpdateHyperLinks(sw);
				// }

				int rowBreakStart = hyperEnd, rowBreakEnd = hyperEnd;
				GetBlockPos(xml, "rowBreaks", ref rowBreakStart, ref rowBreakEnd);
				sw.Write(xml.Substring(hyperEnd, rowBreakStart - hyperEnd));
				//if (rowBreaks.Count > 0)
				//{
				UpdateRowBreaks(sw);
				//}

				int colBreakStart = rowBreakEnd, colBreakEnd = rowBreakEnd;
				GetBlockPos(xml, "colBreaks", ref colBreakStart, ref colBreakEnd);
				sw.Write(xml.Substring(rowBreakEnd, colBreakStart - rowBreakEnd));
				//if (colBreaks.Count > 0)
				//{
				UpdateColBreaks(sw);
				//}
				sw.Write(xml.Substring(colBreakEnd, xml.Length - colBreakEnd));
			}
			sw.Flush();
			//sw.Close();
		}

		private void CleanupMergedCells(MergeCellsCollection _mergedCells)
		{
			int i = 0;
			while (i < _mergedCells.List.Count)
			{
				if (_mergedCells[i] == null)
				{
					_mergedCells.List.RemoveAt(i);
				}
				else
				{
					i++;
				}
			}
		}

		private void UpdateColBreaks(StreamWriter sw)
		{
			StringBuilder breaks = new StringBuilder();
			int count = 0;
			var cse = _values.GetEnumerator(0, 0, 0, ExcelPackage.MaxColumns);
			//foreach (ExcelColumn col in _columns)
			while (cse.MoveNext())
			{
				var col = cse.Value._value as ExcelColumn;
				if (col != null && col.PageBreak)
				{
					breaks.AppendFormat("<brk id=\"{0}\" max=\"16383\" man=\"1\"/>", cse.Column);
					count++;
				}
			}
			if (count > 0)
			{
				sw.Write(string.Format("<colBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</colBreaks>", count, breaks.ToString()));
			}
		}

		private void UpdateRowBreaks(StreamWriter sw)
		{
			StringBuilder breaks = new StringBuilder();
			int count = 0;
			var cse = _values.GetEnumerator(0, 0, ExcelPackage.MaxRows, 0);
			//foreach(ExcelRow row in _rows)
			while (cse.MoveNext())
			{
				var row = cse.Value._value as RowInternal;
				if (row != null && row.PageBreak)
				{
					breaks.AppendFormat("<brk id=\"{0}\" max=\"1048575\" man=\"1\"/>", cse.Row);
					count++;
				}
			}
			if (count > 0)
			{
				sw.Write(string.Format("<rowBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</rowBreaks>", count, breaks.ToString()));
			}
		}

		/// <summary>
		/// Inserts the cols collection into the XML document
		/// </summary>
		private void UpdateColumnData(StreamWriter sw)
		{
			var cse = _values.GetEnumerator(0, 1, 0, ExcelPackage.MaxColumns);
			bool first = true;
			while (cse.MoveNext())
			{
				if (first)
				{
					sw.Write("<cols>");
					first = false;
				}
				var col = cse.Value._value as ExcelColumn;
				ExcelStyleCollection<ExcelXfs> cellXfs = this.Package.Workbook.Styles.CellXfs;

				sw.Write("<col min=\"{0}\" max=\"{1}\"", col.ColumnMin, col.ColumnMax);
				if (col.Hidden == true)
				{
					sw.Write(" hidden=\"1\"");
				}
				else if (col.BestFit)
				{
					sw.Write(" bestFit=\"1\"");
				}
				sw.Write(string.Format(CultureInfo.InvariantCulture, " width=\"{0}\" customWidth=\"1\"", col.Width));

				if (col.OutlineLevel > 0)
				{
					sw.Write(" outlineLevel=\"{0}\" ", col.OutlineLevel);
					if (col.Collapsed)
					{
						if (col.Hidden)
						{
							sw.Write(" collapsed=\"1\"");
						}
						else
						{
							sw.Write(" collapsed=\"1\" hidden=\"1\""); //Always hidden
						}
					}
				}
				if (col.Phonetic)
				{
					sw.Write(" phonetic=\"1\"");
				}

				var styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;
				if (styleID > 0)
				{
					sw.Write(" style=\"{0}\"", styleID);
				}
				sw.Write("/>");
			}
			if (!first)
			{
				sw.Write("</cols>");
			}
		}

		/// <summary>
		/// Insert row and cells into the XML document
		/// </summary>
		private void UpdateRowCellData(StreamWriter sw)
		{
			ExcelStyleCollection<ExcelXfs> cellXfs = this.Package.Workbook.Styles.CellXfs;

			int row = -1;

			StringBuilder sbXml = new StringBuilder();
			var ss = this.Package.Workbook.SharedStrings;
			var styles = this.Package.Workbook.Styles;
			var cache = new StringBuilder();
			cache.Append("<sheetData>");

			////Set a value for cells with style and no value set.
			//var cseStyle =CellStoreEnumeratorFactory<ExcelCoreValue>.GetNewEnumerator(_values, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
			//foreach (var s in cseStyle)
			//{
			//    if(!ExistsValueInner(cseStyle.Row, cseStyle.Column))
			//    {
			//        SetValueInner(cseStyle.Row, cseStyle.Column, null);
			//    }
			//}

			columnStyles = new Dictionary<int, int>();
			var cse = _values.GetEnumerator(1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
			//foreach (IRangeID r in _cells)
			while (cse.MoveNext())
			{
				if (cse.Column > 0)
				{
					var val = cse.Value;
					//int styleID = cellXfs[styles.GetStyleId(this, cse.Row, cse.Column)].newID;
					int styleID = cellXfs[(val._styleId == 0 ? GetStyleIdDefaultWithMemo(cse.Row, cse.Column) : val._styleId)].newID;
					//Add the row element if it's a new row
					if (cse.Row != row)
					{
						WriteRow(cache, cellXfs, row, cse.Row);
						row = cse.Row;
					}
					object v = val._value;
					object formula = _formulas.GetValue(cse.Row, cse.Column);
					if (formula is int)
					{
						int sfId = (int)formula;
						var f = _sharedFormulas[(int)sfId];
						if (f.Address.IndexOf(':') > 0)
						{
							if (f.StartCol == cse.Column && f.StartRow == cse.Row)
							{
								if (f.IsArray)
								{
									cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
								}
								else
								{
									cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{6}><f ref=\"{2}\" t=\"shared\" si=\"{3}\">{4}</f>{5}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, sfId, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
								}
							}
							else if (f.IsArray)
							{
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, styleID < 0 ? 0 : styleID);
							}
							else
							{
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{4}><f t=\"shared\" si=\"{2}\"/>{3}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, sfId, GetFormulaValue(v), GetCellType(v, true));
							}
						}
						else
						{
							// We can also have a single cell array formula
							if (f.IsArray)
							{
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, string.Format("{0}:{1}", f.Address, f.Address), ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
							}
							else
							{
								object resultValue = this.TryGetTypedValue(v, out resultValue) ? resultValue : v;
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", f.Address, styleID < 0 ? 0 : styleID, GetCellType(resultValue, true));
								cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(resultValue));
							}
						}
					}
					else if (!string.IsNullOrEmpty(formula?.ToString()))
					{
						object resultValue = this.TryGetTypedValue(v, out resultValue) ? resultValue : v;
						cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(resultValue, true));
						cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(formula.ToString()), GetFormulaValue(resultValue));
					}
					else
					{
						if (v == null && styleID > 0)
						{
							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, styleID < 0 ? 0 : styleID);
						}
						else if (v != null)
						{
							// Fix for issue 15460
							if (v is System.Collections.IEnumerable enumerableResult && !(v is string))
							{
								var enumerator = enumerableResult.GetEnumerator();
								if (enumerator.MoveNext() && enumerator.Current != null)
									v = enumerator.Current;
								else
									v = string.Empty;
							}
							if (this.TryGetTypedValue(v, out object resultValue))
							{
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(resultValue));
								cache.AppendFormat("{0}</c>", GetFormulaValue(resultValue));
							}
							else
							{
								var vString = Convert.ToString(v);
								int ix;
								if (!ss.ContainsKey(vString))
								{
									ix = ss.Count;
									ss.Add(vString, new ExcelWorkbook.SharedStringItem() { isRichText = _flags.GetFlagValue(cse.Row, cse.Column, CellFlags.RichText), pos = ix });
								}
								else
								{
									ix = ss[vString].pos;
								}
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cse.CellAddress, styleID < 0 ? 0 : styleID);
								cache.AppendFormat("<v>{0}</v></c>", ix);
							}
						}
					}
				}
				else  //ExcelRow
				{
					WriteRow(cache, cellXfs, row, cse.Row);
					row = cse.Row;
				}
				if (cache.Length > 0x600000)
				{
					sw.Write(cache.ToString());
					sw.Flush();
					cache.Length = 0;
				}
			}
			columnStyles = null;

			if (row != -1) cache.Append("</row>");
			cache.Append("</sheetData>");
			sw.Write(cache.ToString());
			sw.Flush();
		}

		private bool TryGetTypedValue(object v, out object resultValue)
		{
			resultValue = null;
			eErrorType valueErrorType = default(eErrorType);
			if (v?.GetType()?.IsPrimitive == true || v is double || v is decimal || v is DateTime || v is TimeSpan || v is ExcelErrorValue
				|| (v is string stringValue && ExcelErrorValue.Values.TryGetErrorType(stringValue, out valueErrorType)))
			{
				resultValue = valueErrorType == default(eErrorType) ? v : ExcelErrorValue.Create(valueErrorType);
				return true;
			}
			return false;
		}

		// get StyleID without cell style for UpdateRowCellData
		internal int GetStyleIdDefaultWithMemo(int row, int col)
		{
			int v = 0;
			if (ExistsStyleInner(row, 0, ref v)) //First Row
			{
				return v;
			}
			else // then column
			{
				if (!columnStyles.ContainsKey(col))
				{
					if (ExistsStyleInner(0, col, ref v))
					{
						columnStyles.Add(col, v);
					}
					else
					{
						int r = 0, c = col;
						if (_values.PrevCell(ref r, ref c))
						{
							//var column=ws.GetValueInner(0,c) as ExcelColumn;
							var val = _values.GetValue(0, c);
							var column = (ExcelColumn)(val._value);
							if (column != null && column.ColumnMax >= col) //Fixes issue 15174
							{
								//return ws.GetStyleInner(0, c);
								columnStyles.Add(col, val._styleId);
							}
							else
							{
								columnStyles.Add(col, 0);
							}
						}
						else
						{
							columnStyles.Add(col, 0);
						}
					}
				}
				return columnStyles[col];
			}
		}

		private object GetFormulaValue(object v)
		{
			//if (this.Package.Workbook._isCalculated)
			//{
			if (v != null && v.ToString() != "")
			{
				return "<v>" + ConvertUtil.ExcelEscapeString(GetValueForXml(v)) + "</v>"; //Fixes issue 15071
			}
			else
			{
				return "";
			}
		}

		private string GetCellType(object v, bool allowStr = false)
		{
			if (v is bool)
			{
				return " t=\"b\"";
			}
			else if ((v is double && double.IsInfinity((double)v)) || v is ExcelErrorValue)
			{
				return " t=\"e\"";
			}
			else if (allowStr && v != null && !(v.GetType().IsPrimitive || v is double || v is decimal || v is DateTime || v is TimeSpan))
			{
				return " t=\"str\"";
			}
			else
			{
				return "";
			}
		}

		private string GetValueForXml(object v)
		{
			string s;
			try
			{
				if (v is DateTime)
				{
					double sdv = ((DateTime)v).ToOADate();

					if (Workbook.Date1904)
					{
						sdv -= ExcelWorkbook.date1904Offset;
					}
					s = sdv.ToString(CultureInfo.InvariantCulture);
				}
				else if (v is TimeSpan)
				{
					s = DateTime.FromOADate(0).Add(((TimeSpan)v)).ToOADate().ToString(CultureInfo.InvariantCulture);
				}
				else if (v.GetType().IsPrimitive || v is double || v is decimal)
				{
					if (v is double && double.IsNaN((double)v))
					{
						s = "";
					}
					else if (v is double && double.IsInfinity((double)v))
					{
						s = "#NUM!";
					}
					else
					{
						s = Convert.ToDouble(v, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
					}
				}
				else
				{
					s = v.ToString();
				}
			}

			catch
			{
				s = "0";
			}
			return s;
		}

		private void WriteRow(StringBuilder cache, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
		{
			if (prevRow != -1) cache.Append("</row>");
			//ulong rowID = ExcelRow.GetRowID(SheetID, row);
			cache.AppendFormat("<row r=\"{0}\"", row);
			RowInternal currRow = GetValueInner(row, 0) as RowInternal;
			if (currRow != null)
			{

				// if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
				if (currRow.Hidden == true)
				{
					cache.Append(" hidden=\"1\"");
				}
				if (currRow.Height >= 0)
				{
					cache.AppendFormat(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));
					if (currRow.CustomHeight)
					{
						cache.Append(" customHeight=\"1\"");
					}
				}

				if (currRow.OutlineLevel > 0)
				{
					cache.AppendFormat(" outlineLevel =\"{0}\"", currRow.OutlineLevel);
					if (currRow.Collapsed)
					{
						if (currRow.Hidden)
						{
							cache.Append(" collapsed=\"1\"");
						}
						else
						{
							cache.Append(" collapsed=\"1\" hidden=\"1\""); //Always hidden
						}
					}
				}
				if (currRow.Phonetic)
				{
					cache.Append(" ph=\"1\"");
				}
			}
			var s = GetStyleInner(row, 0);
			if (s > 0)
			{
				cache.AppendFormat(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID);
			}
			cache.Append(">");
		}

		private void WriteRow(StreamWriter sw, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
		{
			if (prevRow != -1) sw.Write("</row>");
			//ulong rowID = ExcelRow.GetRowID(SheetID, row);
			sw.Write("<row r=\"{0}\"", row);
			RowInternal currRow = GetValueInner(row, 0) as RowInternal;
			if (currRow != null)
			{

				// if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
				if (currRow.Hidden == true)
				{
					sw.Write(" hidden=\"1\"");
				}
				if (currRow.Height >= 0)
				{
					sw.Write(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));
					if (currRow.CustomHeight)
					{
						sw.Write(" customHeight=\"1\"");
					}
				}

				if (currRow.OutlineLevel > 0)
				{
					sw.Write(" outlineLevel =\"{0}\"", currRow.OutlineLevel);
					if (currRow.Collapsed)
					{
						if (currRow.Hidden)
						{
							sw.Write(" collapsed=\"1\"");
						}
						else
						{
							sw.Write(" collapsed=\"1\" hidden=\"1\""); //Always hidden
						}
					}
				}
				if (currRow.Phonetic)
				{
					sw.Write(" ph=\"1\"");
				}
			}
			var s = GetStyleInner(row, 0);
			if (s > 0)
			{
				sw.Write(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID);
			}
			sw.Write(">");
		}

		/// <summary>
		/// Update xml with hyperlinks
		/// </summary>
		/// <param name="sw">The stream</param>
		private void UpdateHyperLinks(StreamWriter sw)
		{
			Dictionary<string, string> hyps = new Dictionary<string, string>();
			var cse = _hyperLinks.GetEnumerator();
			bool first = true;
			//foreach (ulong cell in _hyperLinks)
			while (cse.MoveNext())
			{
				if (first)
				{
					sw.Write("<hyperlinks>");
					first = false;
				}
				//int row, col;
				var uri = _hyperLinks.GetValue(cse.Row, cse.Column);
				//ExcelCell cell = _cells[cellId] as ExcelCell;
				if (uri is ExcelHyperLink && !string.IsNullOrEmpty((uri as ExcelHyperLink).ReferenceAddress))
				{
					ExcelHyperLink hl = uri as ExcelHyperLink;
					sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\"{2}{3}/>",
							Cells[cse.Row, cse.Column, cse.Row + hl.RowSpann, cse.Column + hl.ColSpann].Address,
							ExcelCellBase.GetFullAddress(SecurityElement.Escape(Name), SecurityElement.Escape(hl.ReferenceAddress)),
								string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
								string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
				}
				else if (uri != null)
				{
					string id;
					Uri hyp;
					if (uri is ExcelHyperLink)
					{
						hyp = ((ExcelHyperLink)uri).OriginalUri;
					}
					else
					{
						hyp = uri;
					}
					if (hyps.ContainsKey(hyp.OriginalString))
					{
						id = hyps[hyp.OriginalString];
					}
					else
					{
						var relationship = Part.CreateRelationship(hyp, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
						if (uri is ExcelHyperLink)
						{
							ExcelHyperLink hl = uri as ExcelHyperLink;
							sw.Write("<hyperlink ref=\"{0}\"{2}{3} r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id,
								string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
								string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
						}
						else
						{
							sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id);
						}
						id = relationship.Id;
					}
					//cell.HyperLinkRId = id;
				}
			}
			if (!first)
			{
				sw.Write("</hyperlinks>");
			}
		}

		/// <summary>
		/// Create the hyperlinks node in the XML
		/// </summary>
		/// <returns></returns>
		private XmlNode CreateHyperLinkCollection()
		{
			XmlElement hl = _worksheetXml.CreateElement("hyperlinks", ExcelPackage.schemaMain);
			XmlNode prevNode = _worksheetXml.SelectSingleNode("//d:conditionalFormatting", NameSpaceManager);
			if (prevNode == null)
			{
				prevNode = _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager);
				if (prevNode == null)
				{
					prevNode = _worksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
				}
			}
			return _worksheetXml.DocumentElement.InsertAfter(hl, prevNode);
		}

		private void UpdateConditionalFormatting(int rowFrom, int rows, int columnFrom, int columns)
		{
			var rulesToDelete = new List<IExcelConditionalFormattingRule>();
			foreach (var rule in this.ConditionalFormatting)
			{
				if (this.EntirelyInRemovedRows(rule.Address.ToString(), rowFrom, rows) || this.EntirelyInRemovedColumns(rule.Address.ToString(), columnFrom, columns))
					rulesToDelete.Add(rule);
				else
					rule.Address = new ExcelAddress(this.UpdateAddresses(rule.Address.ToString(), rowFrom, rows, columnFrom, columns));
			}
			this.ConditionalFormatting.TransformFormulaReferences(f => this.UpdateAddresses(f, rowFrom, rows, columnFrom, columns));
			foreach (var rule in rulesToDelete)
			{
				this.ConditionalFormatting.Remove(rule);
			}
		}

		private void UpdateX14ConditionalFormatting(int rowFrom, int rows, int columnFrom, int columns)
		{
			var rulesToDelete = new List<X14CondtionalFormattingRule>();
			foreach (var rule in this.X14ConditionalFormatting.X14Rules)
			{
				if (this.EntirelyInRemovedRows(rule.Address, rowFrom, rows) || this.EntirelyInRemovedColumns(rule.Address, columnFrom, columns))
					rulesToDelete.Add(rule);
				else
					rule.Address = this.UpdateAddresses(rule.Address, rowFrom, rows, columnFrom, columns);
			}
			this.X14ConditionalFormatting.TransformFormulaReferences(f => this.UpdateAddresses(f, rowFrom, rows, columnFrom, columns));
			foreach (var rule in rulesToDelete)
			{
				this.X14ConditionalFormatting.X14Rules.Remove(rule);
			}
		}

		private string UpdateAddresses(string originalAddress, int rowFrom, int rows, int columnFrom, int columns)
		{
			const char seperator = ',';
			List<string> movedAddresses = new List<string>();
			foreach (var stringAddress in originalAddress.ToString().Split(seperator))
			{
				var newAddress = this.Package.FormulaManager.UpdateFormulaReferences(stringAddress, -rows, -columns, rowFrom, columnFrom, this.Name, this.Name);
				if(newAddress != Values.Ref)
					movedAddresses.Add(newAddress);
			}
			return string.Join(seperator.ToString(), movedAddresses);
		}

		private bool EntirelyInRemovedRows(string originalAddress, int rowFrom, int rows)
		{
			return originalAddress.Split(',').All(addressString => 
			{
				var address = new ExcelAddress(addressString);
				return address.Start.Row >= rowFrom && address.End.Row <= rowFrom + rows - 1;
			});
		}

		private bool EntirelyInRemovedColumns(string originalAddress, int columnFrom, int columns)
		{
			return originalAddress.Split(',').All(addressString =>
			{
				var address = new ExcelAddress(addressString);
				return address.Start.Column >= columnFrom && address.End.Column <= columnFrom + columns - 1;
			});
		}

		private void UpdateDataValidationRanges(int rowFrom, int rows, int columnFrom, int columns)
		{
			if (rows < 0 || columns < 0)
				this.DataValidations.RemoveAll(validation => this.RangeIsBeingDeleted(validation.Address._fromRow, validation.Address._toRow, validation.Address._fromCol, validation.Address._toCol, rowFrom, rows, columnFrom, columns));
			foreach (var sheet in this.Workbook.Worksheets)
			{
				for (int i = sheet.DataValidations.Count - 1; i >= 0; i--)
				{
					var validation = sheet.DataValidations.ElementAt(i);
					var newAddress = this.Package.FormulaManager.UpdateFormulaReferences(validation.Address.Address, rows, columns, rowFrom, columnFrom, sheet.Name, this.Name);
					validation.Address = new ExcelAddress(newAddress);
					if (validation is ExcelDataValidationAny anyValidation)
					{
						// No formulas
					}
					else if (validation is ExcelDataValidationCustom customValidation)
					{
						if (!string.IsNullOrEmpty(customValidation.Formula?.ExcelFormula))
							customValidation.Formula.ExcelFormula = this.TranslateDataValidationFormula(sheet, customValidation.Address, customValidation.Formula.ExcelFormula, rowFrom, rows, columnFrom, columns);
					}
					else if (validation is ExcelDataValidationList listValidation)
					{
						if (!string.IsNullOrEmpty(listValidation.Formula?.ExcelFormula))
							listValidation.Formula.ExcelFormula = this.TranslateDataValidationFormula(sheet, listValidation.Address, listValidation.Formula.ExcelFormula, rowFrom, rows, columnFrom, columns);
					}
					else if (validation is ExcelDataValidationTime timeValidation)
					{
						if (!string.IsNullOrEmpty(timeValidation.Formula?.ExcelFormula))
							timeValidation.Formula.ExcelFormula = this.TranslateDataValidationFormula(sheet, timeValidation.Address, timeValidation.Formula.ExcelFormula, rowFrom, rows, columnFrom, columns);
						if (!string.IsNullOrEmpty(timeValidation.Formula2?.ExcelFormula))
							timeValidation.Formula2.ExcelFormula = this.TranslateDataValidationFormula(sheet, timeValidation.Address, timeValidation.Formula2.ExcelFormula, rowFrom, rows, columnFrom, columns);
					}
					else if (validation is ExcelDataValidationDateTime dateTimeValidation)
					{
						if (!string.IsNullOrEmpty(dateTimeValidation.Formula?.ExcelFormula))
							dateTimeValidation.Formula.ExcelFormula = this.TranslateDataValidationFormula(sheet, dateTimeValidation.Address, dateTimeValidation.Formula.ExcelFormula, rowFrom, rows, columnFrom, columns);
						if (!string.IsNullOrEmpty(dateTimeValidation.Formula2?.ExcelFormula))
							dateTimeValidation.Formula2.ExcelFormula = this.TranslateDataValidationFormula(sheet, dateTimeValidation.Address, dateTimeValidation.Formula2.ExcelFormula, rowFrom, rows, columnFrom, columns);
					}
					else if (validation is ExcelDataValidationInt intValidation)
					{
						if (!string.IsNullOrEmpty(intValidation.Formula?.ExcelFormula))
							intValidation.Formula.ExcelFormula = this.TranslateDataValidationFormula(sheet, intValidation.Address, intValidation.Formula.ExcelFormula, rowFrom, rows, columnFrom, columns);
						if (!string.IsNullOrEmpty(intValidation.Formula2?.ExcelFormula))
							intValidation.Formula2.ExcelFormula = this.TranslateDataValidationFormula(sheet, intValidation.Address, intValidation.Formula2.ExcelFormula, rowFrom, rows, columnFrom, columns);
					}
					else if (validation is ExcelDataValidationDecimal decimalValidation)
					{
						if (!string.IsNullOrEmpty(decimalValidation.Formula?.ExcelFormula))
							decimalValidation.Formula.ExcelFormula = this.TranslateDataValidationFormula(sheet, decimalValidation.Address, decimalValidation.Formula.ExcelFormula, rowFrom, rows, columnFrom, columns);
						if (!string.IsNullOrEmpty(decimalValidation.Formula2?.ExcelFormula))
							decimalValidation.Formula2.ExcelFormula = this.TranslateDataValidationFormula(sheet, decimalValidation.Address, decimalValidation.Formula2.ExcelFormula, rowFrom, rows, columnFrom, columns);
					}
				}
			}
		}

		private void UpdateX14DataValidationRanges(int rowFrom, int rows, int columnFrom, int columns)
		{
			if (rows < 0 || columns < 0)
				this.X14DataValidations.RemoveAll(validation => this.RangeIsBeingDeleted(validation.Address._fromRow, validation.Address._toRow, validation.Address._fromCol, validation.Address._toCol, rowFrom, rows, columnFrom, columns));
			foreach (var sheet in this.Workbook.Worksheets)
			{
				for (int i = sheet.X14DataValidations.Count - 1; i >= 0; i--)
				{
					var validation = sheet.X14DataValidations.ElementAt(i);
					var newAddress = this.Package.FormulaManager.UpdateFormulaReferences(validation.Address.Address, rows, columns, rowFrom, columnFrom, sheet.Name, this.Name);
					validation.Address = new ExcelAddress(newAddress);
					if (validation is ExcelX14DataValidation x14Validation)
					{
						if (!string.IsNullOrEmpty(x14Validation.Formula))
							x14Validation.Formula = this.TranslateDataValidationFormula(sheet, x14Validation.Address, x14Validation.Formula, rowFrom, rows, columnFrom, columns);
						if (!string.IsNullOrEmpty(x14Validation.Formula2))
							x14Validation.Formula2 = this.TranslateDataValidationFormula(sheet, x14Validation.Address, x14Validation.Formula2, rowFrom, rows, columnFrom, columns);
					}
				}
			}
		}

		private string TranslateDataValidationFormula(ExcelWorksheet validationSheet, ExcelAddress validationAddress, string validationFormula, int rowFrom, int rows, int columnFrom, int columns)
		{
			string worksheetName = "!";
			if (validationSheet.Name?.ToUpper() == this.Name.ToUpper()) // This formula references the sheet it is on.
				worksheetName = validationSheet.Name;
			else if (validationAddress.WorkSheet != null && validationAddress.WorkSheet.ToUpper() == this.Name.ToUpper()) // This formula references another sheet in the workbook.
				worksheetName = validationAddress.WorkSheet;
			else if (string.IsNullOrEmpty(validationAddress.WorkSheet))
				worksheetName = validationSheet.Name;
			if (!worksheetName.Equals("!"))// Only update the formula if we have a valid reference to a worksheet.
				return this.Package.FormulaManager.UpdateFormulaReferences(validationFormula, rows, columns, rowFrom, columnFrom, worksheetName, this.Name);
			return validationFormula;
		}

		private bool RangeIsBeingDeleted(int rangeRowFrom, int rangeRowTo, int rangeColumnFrom, int rangeColumnTo, int rowFrom, int rows, int columnFrom, int columns)
		{
			return (rangeRowFrom >= rowFrom && rangeRowTo < rowFrom + -rows ||
				rangeColumnFrom >= columnFrom && rangeColumnTo < columnFrom + -columns);
		}

		private void UpdateSparkLines(int rows, int rowFrom, int columns, int columnFrom)
		{
			this.RemoveDeletedSparklines(rows, rowFrom, columns, columnFrom);
			this.UpdateSparkLineReferences(rows, rowFrom, columns, columnFrom);
		}

		private void RemoveDeletedSparklines(int rows, int rowFrom, int columns, int columnFrom)
		{
			// Only delete sparklines if rows or columns are being deleted.
			if (rows >= 0 && columns >= 0)
				return;

			foreach (var group in this.SparklineGroups.SparklineGroups)
			{
				group.Sparklines.RemoveAll(sparkline => ExcelWorksheet.IsInRange(sparkline.HostCell, -rows, rowFrom, -columns, columnFrom));
			}
			var groupsToDelete = this.SparklineGroups.SparklineGroups.Where(sparklineGroup => sparklineGroup.Sparklines.Count == 0).ToArray();
			foreach (var sparklineGroup in groupsToDelete)
			{
				this.SparklineGroups.TopNode.RemoveChild(sparklineGroup.TopNode);
				this.SparklineGroups.SparklineGroups.Remove(sparklineGroup);
			}
		}

		private static bool IsInRange(ExcelAddress address, int rows, int rowFrom, int columns, int columnFrom)
		{
			if (rows != 0 && address.Start.Row >= rowFrom && address.Start.Row <= rows - 1 + rowFrom)
				return true;
			else if (columns != 0 && address.Start.Column >= columnFrom && address.Start.Column <= columns - 1 + columnFrom)
				return true;
			return false;
		}

		private void UpdateSparkLineReferences(int rows, int rowFrom, int columns, int columnFrom)
		{
			foreach (var sheet in this.Workbook.Worksheets)
			{
				foreach (var group in sheet.SparklineGroups.SparklineGroups)
				{
					foreach (var sparkline in group.Sparklines)
					{
						if (sparkline.Formula != null)
						{
							ExcelRangeBase.SplitAddress(sparkline.Formula.Address, out string workbook, out string worksheet, out string address);
							// Only update the formula if it references the modified sheet.
							if (worksheet == this.Name)
							{
								address = this.Package.FormulaManager.UpdateFormulaReferences(address, rows, columns, rowFrom, columnFrom, this.Name, this.Name);
								if (string.IsNullOrEmpty(worksheet))
									sparkline.Formula.Address = address;
								else
									sparkline.Formula.Address = ExcelRangeBase.GetFullAddress(worksheet, address);
							}
						}
						// Only update host cell if it is in the modified sheet. 
						if (sheet.Name == this.Name)
							sparkline.HostCell.Address = this.Package.FormulaManager.UpdateFormulaReferences(sparkline.HostCell.Address, rows, columns, rowFrom, columnFrom, sheet.Name, this.Name);
					}
				}
			}
		}

		private void ChangeSparklineSheetNames(string newName)
		{
			foreach (var sheet in this.Workbook.Worksheets)
			{
				foreach (var group in sheet.SparklineGroups.SparklineGroups)
				{
					foreach (var sparkline in group.Sparklines)
					{
						if (sparkline.Formula != null)
						{
							ExcelRangeBase.SplitAddress(sparkline.Formula.Address, out var workbook, out var worksheet, out var address);
							if (string.IsNullOrEmpty(worksheet))
								return;
							else if (worksheet.Equals(this.Name))
								sparkline.Formula.SetAddress(ExcelRangeBase.GetFullAddress(newName, address));
						}
						if (sheet.Name == this.Name)
							sparkline.HostCell._ws = newName;
					}
				}
			}
		}

		private static void SetValueInnerUpdate(List<ExcelCoreValue> list, int index, object value)
		{
			list[index] = new ExcelCoreValue { _value = value, _styleId = list[index]._styleId };
		}

		private void SetStyleInnerUpdate(List<ExcelCoreValue> list, int index, object styleId)
		{
			list[index] = new ExcelCoreValue { _value = list[index]._value, _styleId = (int)styleId };
		}

		private void SetRangeValueUpdate(List<ExcelCoreValue> list, int index, int row, int column, object values)
		{
			list[index] = new ExcelCoreValue { _value = ((object[,])values)[row, column], _styleId = list[index]._styleId };
		}
		#endregion
	}
}
