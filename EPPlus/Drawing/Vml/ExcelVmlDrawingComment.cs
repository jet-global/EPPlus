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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
	/// <summary>
	/// Drawing object used for comments
	/// </summary>
	public class ExcelVmlDrawingComment : ExcelVmlDrawingBase, IRangeID
	{
		internal ExcelVmlDrawingComment(XmlNode topNode, ExcelRangeBase range, XmlNamespaceManager ns) :
			 base(topNode, ns)
		{
			this.Range = range;
			this.SchemaNodeOrder = new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
		}
		internal ExcelRangeBase Range { get; set; }

		/// <summary>
		/// Gets or sets the address in the worksheet.
		/// </summary>
		public string Address
		{
			get
			{
				return this.Range.Address;
			}
			internal set
			{
				this.Range.Address = value;
			}
		}

		const string VERTICAL_ALIGNMENT_PATH = "x:ClientData/x:TextVAlign";
		/// <summary>
		/// Gets or sets the vertical alignment for text.
		/// </summary>
		public eTextAlignVerticalVml VerticalAlignment
		{
			get
			{
				switch (GetXmlNodeString(VERTICAL_ALIGNMENT_PATH))
				{
					case "Center":
						return eTextAlignVerticalVml.Center;
					case "Bottom":
						return eTextAlignVerticalVml.Bottom;
					default:
						return eTextAlignVerticalVml.Top;
				}
			}
			set
			{
				switch (value)
				{
					case eTextAlignVerticalVml.Center:
						this.SetXmlNodeString(VERTICAL_ALIGNMENT_PATH, "Center");
						break;
					case eTextAlignVerticalVml.Bottom:
						this.SetXmlNodeString(VERTICAL_ALIGNMENT_PATH, "Bottom");
						break;
					default:
						this.DeleteNode(VERTICAL_ALIGNMENT_PATH);
						break;
				}
			}
		}
		const string HORIZONTAL_ALIGNMENT_PATH = "x:ClientData/x:TextHAlign";
		/// <summary>
		/// Gets or sets the horizontal alignment for text.
		/// </summary>
		public eTextAlignHorizontalVml HorizontalAlignment
		{
			get
			{
				switch (GetXmlNodeString(HORIZONTAL_ALIGNMENT_PATH))
				{
					case "Center":
						return eTextAlignHorizontalVml.Center;
					case "Right":
						return eTextAlignHorizontalVml.Right;
					default:
						return eTextAlignHorizontalVml.Left;
				}
			}
			set
			{
				switch (value)
				{
					case eTextAlignHorizontalVml.Center:
						this.SetXmlNodeString(HORIZONTAL_ALIGNMENT_PATH, "Center");
						break;
					case eTextAlignHorizontalVml.Right:
						this.SetXmlNodeString(HORIZONTAL_ALIGNMENT_PATH, "Right");
						break;
					default:
						this.DeleteNode(HORIZONTAL_ALIGNMENT_PATH);
						break;
				}
			}
		}
		const string VISIBLE_PATH = "x:ClientData/x:Visible";
		/// <summary>
		/// Gets or sets whether the drawing object is visible.
		/// </summary>
		public bool Visible
		{
			get
			{
				return (this.TopNode.SelectSingleNode(VISIBLE_PATH, NameSpaceManager) != null);
			}
			set
			{
				if (value)
				{
					this.CreateNode(VISIBLE_PATH);
					this.Style = this.SetStyle(Style, "visibility", "visible");
				}
				else
				{
					this.DeleteNode(VISIBLE_PATH);
					this.Style = this.SetStyle(Style, "visibility", "hidden");
				}
			}
		}

		const string BACKGROUNDCOLOR_PATH = "@fillcolor";
		const string BACKGROUNDCOLOR2_PATH = "v:fill/@color2";
		/// <summary>
		/// Gets or sets the background color.
		/// </summary>
		public Color BackgroundColor
		{
			get
			{
				string col = GetXmlNodeString(BACKGROUNDCOLOR_PATH);
				if (col == "")
				{
					return Color.FromArgb(0xff, 0xff, 0xe1);
				}
				else
				{
					if (col.StartsWith("#")) col = col.Substring(1, col.Length - 1);
					int res;
					if (int.TryParse(col, System.Globalization.NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out res))
					{
						return Color.FromArgb(res);
					}
					else
					{
						return Color.Empty;
					}
				}
			}
			set
			{
				string color = "#" + value.ToArgb().ToString("X").Substring(2, 6);
				this.SetXmlNodeString(BACKGROUNDCOLOR_PATH, color);
			}
		}
		const string LINESTYLE_PATH = "v:stroke/@dashstyle";
		const string ENDCAP_PATH = "v:stroke/@endcap";
		/// <summary>
		/// Gets or sets the outline style for border.
		/// </summary>
		public eLineStyleVml OutlineStyle
		{
			get
			{
				string v = this.GetXmlNodeString(LINESTYLE_PATH);
				if (v == string.Empty)
				{
					return eLineStyleVml.Solid;
				}
				else if (v == "1 1")
				{
					v = this.GetXmlNodeString(ENDCAP_PATH);
					return (eLineStyleVml)Enum.Parse(typeof(eLineStyleVml), v, true);
				}
				else
				{
					return (eLineStyleVml)Enum.Parse(typeof(eLineStyleVml), v, true);
				}
			}
			set
			{
				if (value == eLineStyleVml.Round || value == eLineStyleVml.Square)
				{
					this.SetXmlNodeString(LINESTYLE_PATH, "1 1");
					if (value == eLineStyleVml.Round)
					{
						this.SetXmlNodeString(ENDCAP_PATH, "round");
					}
					else
					{
						this.DeleteNode(ENDCAP_PATH);
					}
				}
				else
				{
					string v = value.ToString();
					v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
					this.SetXmlNodeString(LINESTYLE_PATH, v);
					this.DeleteNode(ENDCAP_PATH);
				}
			}
		}
		const string LINECOLOR_PATH = "@strokecolor";
		/// <summary>
		/// Gets or sets the outline color for the border.
		/// </summary>
		public Color OutlineColor
		{
			get
			{
				string col = this.GetXmlNodeString(LINECOLOR_PATH);
				if (col == "")
				{
					return Color.Black;
				}
				else
				{
					if (col.StartsWith("#")) col = col.Substring(1, col.Length - 1);
					int res;
					if (int.TryParse(col, System.Globalization.NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out res))
					{
						return Color.FromArgb(res);
					}
					else
					{
						return Color.Empty;
					}
				}
			}
			set
			{
				string color = "#" + value.ToArgb().ToString("X").Substring(2, 6);
				this.SetXmlNodeString(LINECOLOR_PATH, color);
			}
		}
		const string LINEWIDTH_PATH = "@strokeweight";
		/// <summary>
		/// Gets or sets the width of the border in point format.
		/// </summary>
		public float OutlineWidth
		{
			get
			{
				string wt = this.GetXmlNodeString(LINEWIDTH_PATH);
				if (wt == "")
					return (float).75;
				if (wt.EndsWith("pt"))
					wt = wt.Substring(0, wt.Length - 2);
				float ret;
				if (float.TryParse(wt, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out ret))
					return ret;
				else
					return 0;
			}
			set
			{
				this.SetXmlNodeString(LINEWIDTH_PATH, value.ToString(CultureInfo.InvariantCulture) + "pt");
			}
		}
		
		/// <summary>
		/// Gets or sets the width of the comment in point format. 
		/// </summary>
		public float Width
		{
			get
			{
				string v;
				this.GetStyle(GetXmlNodeString(STYLE_PATH), "width", out v);
				if (v.EndsWith("pt"))
					v = v.Substring(0, v.Length - 2);
				float ret;
				if (float.TryParse(v, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out ret))
					return ret;
				else
					return 0;
			}
			set
			{
				this.SetXmlNodeString(STYLE_PATH, SetStyle(GetXmlNodeString(STYLE_PATH), "width", value.ToString("N2", CultureInfo.InvariantCulture) + "pt"));
			}
		}

		/// <summary>
		/// Gets or sets the height of the comment in point format.
		/// </summary>
		public float Height
		{
			get
			{
				string v;
				this.GetStyle(GetXmlNodeString(STYLE_PATH), "height", out v);
				if (v.EndsWith("pt"))
					v = v.Substring(0, v.Length - 2);
				float ret;
				if (float.TryParse(v, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out ret))
					return ret;
				else
					return 0;
			}
			set
			{
				this.SetXmlNodeString(STYLE_PATH, this.SetStyle(GetXmlNodeString(STYLE_PATH), "height", value.ToString("N2", CultureInfo.InvariantCulture) + "pt"));
			}
		}

		/// <summary>
		/// Gets or sets the top margin of the comment in point format.
		/// </summary>
		public float MarginTop
		{
			get
			{
				string v;
				this.GetStyle(GetXmlNodeString(STYLE_PATH), "margin-top", out v);
				if (v.EndsWith("pt"))
					v = v.Substring(0, v.Length - 2);
				float ret;
				if (float.TryParse(v, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out ret))
					return ret;
				else
					return 0;
			}
			set
			{
				this.SetXmlNodeString(STYLE_PATH, this.SetStyle(GetXmlNodeString(STYLE_PATH), "margin-top", value.ToString("N2", CultureInfo.InvariantCulture) + "pt"));
			}
		}

		/// <summary>
		/// Gets or sets the left margin of the comment in point format.
		/// </summary>
		public float MarginLeft
		{
			get
			{
				string v;
				this.GetStyle(GetXmlNodeString(STYLE_PATH), "margin-left", out v);
				if (v.EndsWith("pt"))
					v = v.Substring(0, v.Length - 2);
				float ret;
				if (float.TryParse(v, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out ret))
					return ret;
				else
					return 0;
			}
			set
			{
				this.SetXmlNodeString(STYLE_PATH, SetStyle(GetXmlNodeString(STYLE_PATH), "margin-left", value.ToString("N2", CultureInfo.InvariantCulture) + "pt"));
			}
		}

		/// <summary>
		/// Gets or sets whether to autofit the comment.
		/// </summary>
		const string TEXTBOX_STYLE_PATH = "v:textbox/@style";
		public bool AutoFit
		{
			get
			{
				string value;
				this.GetStyle(GetXmlNodeString(TEXTBOX_STYLE_PATH), "mso-fit-shape-to-text", out value);
				return value == "t";
			}
			set
			{
				this.SetXmlNodeString(TEXTBOX_STYLE_PATH, SetStyle(GetXmlNodeString(TEXTBOX_STYLE_PATH), "mso-fit-shape-to-text", value ? "t" : ""));
			}
		}

		/// <summary>
		/// Gets or sets whether the object is locked when the sheet is protected.
		/// </summary>
		const string LOCKED_PATH = "x:ClientData/x:Locked";
		public bool Locked
		{
			get
			{
				return this.GetXmlNodeBool(LOCKED_PATH, false);
			}
			set
			{
				this.SetXmlNodeBool(LOCKED_PATH, value, false);
			}
		}

		/// <summary>
		/// Gets or sets whether the object's text is locked.
		/// </summary>
		const string LOCK_TEXT_PATH = "x:ClientData/x:LockText";
		public bool LockText
		{
			get
			{
				return this.GetXmlNodeBool(LOCK_TEXT_PATH, false);
			}
			set
			{
				this.SetXmlNodeBool(LOCK_TEXT_PATH, value, false);
			}
		}

		/// <summary>
		/// Gets or sets the from-position. Only applies when Visible is set to true.
		/// </summary>
		ExcelVmlDrawingPosition _from = null;
		public ExcelVmlDrawingPosition From
		{
			get
			{
				if (_from == null)
				{
					_from = new ExcelVmlDrawingPosition(this.NameSpaceManager, this.TopNode.SelectSingleNode("x:ClientData", this.NameSpaceManager), 0);
				}
				return _from;
			}
		}

		/// <summary>
		/// Gets or sets the to-position. Only applies when Visible is set to true.
		/// </summary>
		ExcelVmlDrawingPosition _to = null;
		public ExcelVmlDrawingPosition To
		{
			get
			{
				if (_to == null)
				{
					_to = new ExcelVmlDrawingPosition(this.NameSpaceManager, this.TopNode.SelectSingleNode("x:ClientData", this.NameSpaceManager), 4);
				}
				return _to;
			}
		}

		/// <summary>
		/// Gets or sets the row position of the comment.
		/// </summary>
		const string ROW_PATH = "x:ClientData/x:Row";
		internal int Row
		{
			get
			{
				return this.GetXmlNodeInt(ROW_PATH);
			}
			set
			{
				this.SetXmlNodeString(ROW_PATH, value.ToString(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets or sets the column position of the comment.
		/// </summary>
		const string COLUMN_PATH = "x:ClientData/x:Column";
		internal int Column
		{
			get
			{
				return this.GetXmlNodeInt(COLUMN_PATH);
			}
			set
			{
				this.SetXmlNodeString(COLUMN_PATH, value.ToString(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets or sets the style of the comment.
		/// </summary>
		const string STYLE_PATH = "@style";
		internal string Style
		{
			get
			{
				return this.GetXmlNodeString(STYLE_PATH);
			}
			set
			{
				this.SetXmlNodeString(STYLE_PATH, value);
			}
		}

		#region IRangeID Members
		ulong IRangeID.RangeID
		{
			get
			{
				return ExcelCellBase.GetCellID(this.Range.Worksheet.SheetID, this.Range.Start.Row, this.Range.Start.Column);
			}
			set
			{

			}
		}
		#endregion
	}
}
