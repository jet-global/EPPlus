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
 * Eyal Seagull        Added       		  2012-04-03
 *******************************************************************************/
using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// Represents a conditional formatting rule in Excel.
	/// </summary>
	/// <remarks>
	/// This class does not align with the XML model of cfRule and conditionalFormatting elements.
	/// </remarks>
	public abstract class ExcelConditionalFormattingRule : XmlHelper, IExcelConditionalFormattingRule
	{
		#region Class Variables
		private ExcelDxfStyleConditionalFormatting myStyle = null;
		private eExcelConditionalFormattingRuleType? myType;

		/// <summary>
		/// Indicate that we are in a changing Priorities opeartion so that we won't enter
		/// a recursive loop.
		/// </summary>
		private static bool myChangingPriority = false;
		#endregion

		#region Public Properties
		/// <summary>
		/// Get the &lt;cfRule&gt; node.
		/// </summary>
		public XmlNode Node
		{
			get { return this.TopNode; }
		}

		/// <summary>
		/// Gets or sets the address of the conditional formatting rule.
		/// </summary>
		/// <remarks>
		/// The address is stored in a parent node called &lt;conditionalFormatting&gt; in the
		/// @sqref attribute. Excel (sometimes) groups rules that have the same address inside one node,
		/// but there are cases when it doesn't such as in pivot table conditional formattings.
		/// </remarks>
		public ExcelAddress Address
		{
			get
			{
				string attribute = this.Node.ParentNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value;
				return new ExcelAddress(attribute.Replace(' ',','));
			}
			set
			{
				if (this.Address.Address != value.Address)
				{
					XmlNode conditionalFormattingNode = this.Node.ParentNode;
					if (conditionalFormattingNode.ChildNodes.Count > 1)
					{
						XmlNode clonedConditionalFormattingNode = conditionalFormattingNode.CloneNode(false);
						conditionalFormattingNode.ParentNode.InsertBefore(clonedConditionalFormattingNode, conditionalFormattingNode);
						conditionalFormattingNode.RemoveChild(this.Node);
						this.TopNode = clonedConditionalFormattingNode.AppendChild(this.Node);
					}
					XmlHelper.SetAttribute(this.Node.ParentNode, ExcelConditionalFormattingConstants.Attributes.Sqref, value.AddressSpaceSeparated);
				}
			}
		}

		/// <summary>
		/// Gets or sets the type of conditional formatting rule. ST_CfType §18.18.12.
		/// </summary>
		public eExcelConditionalFormattingRuleType Type
		{
			get
			{
				// Transform the @type attribute to EPPlus Rule Type.
				if (this.myType == null)
				{
					this.myType = ExcelConditionalFormattingRuleType.GetTypeByAttrbiute(
					GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TypeAttribute),
					this.TopNode,
					this.Worksheet.NameSpaceManager);
				}
				return (eExcelConditionalFormattingRuleType)this.myType;
			}
			internal set
			{
				this.myType = value;
				// Transform the EPPlus Rule Type to @type attribute.
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.TypeAttribute,
				  ExcelConditionalFormattingRuleType.GetAttributeByType(value),
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the conditional formatting rule's priority.
		/// This value is used to determine which format should be evaluated and rendered. 
		/// Lower numeric values are higher priority than higher numeric values, where 1 is the highest priority.
		/// </summary>
		public int Priority
		{
			get
			{
				return GetXmlNodeInt(ExcelConditionalFormattingConstants.Paths.PriorityAttribute);
			}
			set
			{
				int priority = this.Priority;
				if (priority != value)
				{
					// Check if we are not already inside a "Change Priority" operation.
					if (!ExcelConditionalFormattingRule.myChangingPriority)
					{
						if (value < 1)
							throw new IndexOutOfRangeException(ExcelConditionalFormattingConstants.Errors.InvalidPriority);
						// Indicate that we are already changing cfRules priorities.
						ExcelConditionalFormattingRule.myChangingPriority = true;
						// Check if we lowered the priority.
						if (priority > value)
						{
							for (int i = priority - 1; i >= value; i--)
							{
								var cfRule = this.Worksheet.ConditionalFormatting.RulesByPriority(i);
								if (cfRule != null)
									cfRule.Priority++;
							}
						}
						else
						{
							for (int i = priority + 1; i <= value; i++)
							{
								var cfRule = this.Worksheet.ConditionalFormatting.RulesByPriority(i);
								if (cfRule != null)
									cfRule.Priority--;
							}
						}
						// Indicate that we are no longer changing cfRules priorities.
						ExcelConditionalFormattingRule.myChangingPriority = false;
					}
					// Change the priority in the XML.
					SetXmlNodeString(
						ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
						value.ToString(),
						true);
				}
			}
		}

		/// <summary>
		/// Gets or sets the StopIfTrue value.
		/// If this flag is true, no rules with lower priority shall be applied over this rule,
		/// when this rule evaluates to true.
		/// </summary>
		public bool StopIfTrue
		{
			get
			{
				return GetXmlNodeBool(ExcelConditionalFormattingConstants.Paths.StopIfTrueAttribute);
			}
			set
			{
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.StopIfTrueAttribute,
				  (value == true) ? "1" : string.Empty,
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the StdDev (zero is not allowed and will be converted to 1).
		/// </summary>
		public UInt16 StdDev
		{
			get
			{
				return Convert.ToUInt16(GetXmlNodeInt(ExcelConditionalFormattingConstants.Paths.StdDevAttribute));
			}
			set
			{
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.StdDevAttribute,
				  (value == 0) ? "1" : value.ToString(),
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the rank (zero is not allowed and will be converted to 1).
		/// </summary>
		public UInt16 Rank
		{
			get
			{
				return Convert.ToUInt16(GetXmlNodeInt(ExcelConditionalFormattingConstants.Paths.RankAttribute));
			}
			set
			{
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.RankAttribute,
				  (value == 0) ? "1" : value.ToString(),
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the DfxStyle of the conditional formatting.
		/// </summary>
		public ExcelDxfStyleConditionalFormatting Style
		{
			get
			{
				if (this.myStyle == null)
				{
					this.myStyle = new ExcelDxfStyleConditionalFormatting(this.NameSpaceManager, null, this.Worksheet.Workbook.Styles);
				}
				return this.myStyle;
			}
		}
		#endregion

		#region Internal Properties
		/// <summary>
		/// Gets or sets whether or not the conditional formatting rule is AboveAverage.
		/// </summary>
		internal protected bool? AboveAverage
		{
			get
			{
				bool? aboveAverage = GetXmlNodeBoolNullable( ExcelConditionalFormattingConstants.Paths.AboveAverageAttribute);
				// Above Avarege if TRUE or if attribute does not exist.
				return (aboveAverage == true) || (aboveAverage == null);
			}
			set
			{
				string aboveAverageValue = string.Empty;
				// Only the types that needs @AboveAverage.
				if ((this.myType == eExcelConditionalFormattingRuleType.BelowAverage)
					|| (this.myType == eExcelConditionalFormattingRuleType.BelowOrEqualAverage)
					|| (this.myType == eExcelConditionalFormattingRuleType.BelowStdDev))
				{
					aboveAverageValue = "0";
				}

				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.AboveAverageAttribute,
				  aboveAverageValue,
				  true);
			}
		}

		/// <summary>
		/// Gets or sets whether or not the conditional formatting rule is EqualAverage.
		/// </summary>
		internal protected bool? EqualAverage
		{
			get
			{
				bool? equalAverage = GetXmlNodeBoolNullable(ExcelConditionalFormattingConstants.Paths.EqualAverageAttribute);
				// Equal Avarege only if TRUE.
				return (equalAverage == true);
			}
			set
			{
				string equalAverageValue = string.Empty;
				// Only the types that needs @EqualAverage
				if ((this.myType == eExcelConditionalFormattingRuleType.AboveOrEqualAverage)
				  || (this.myType == eExcelConditionalFormattingRuleType.BelowOrEqualAverage))
				{
					equalAverageValue = "1";
				}
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.EqualAverageAttribute,
				  equalAverageValue,
				  true);
			}
		}

		/// <summary>
		/// Gets or sets whether or not the conditional formatting rule is Bottom.
		/// </summary>
		internal protected bool? Bottom
		{
			get
			{
				bool? bottom = GetXmlNodeBoolNullable(ExcelConditionalFormattingConstants.Paths.BottomAttribute);
				// Bottom only if TRUE.
				return (bottom == true);
			}
			set
			{
				string bottomValue = string.Empty;

				// Only the types that need @Bottom.
				if ((this.myType == eExcelConditionalFormattingRuleType.Bottom)
				  || (this.myType == eExcelConditionalFormattingRuleType.BottomPercent))
				{
					bottomValue = "1";
				}

				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.BottomAttribute,
				  bottomValue,
				  true);
			}
		}

		/// <summary>
		/// Gets or sets whether or not the conditional formatting rule is Percent.
		/// </summary>
		internal protected bool? Percent
		{
			get
			{
				bool? percent = GetXmlNodeBoolNullable(ExcelConditionalFormattingConstants.Paths.PercentAttribute);
				// Percent if TRUE.
				return (percent == true);
			}
			set
			{
				string percentValue = string.Empty;

				// Only the types that needs the @Percent.
				if ((this.myType == eExcelConditionalFormattingRuleType.BottomPercent)
				  || (this.myType == eExcelConditionalFormattingRuleType.TopPercent))
				{
					percentValue = "1";
				}
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.PercentAttribute,
				  percentValue,
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the conditional foramtting rule TimePeriod type.
		/// </summary>
		internal protected eExcelConditionalFormattingTimePeriodType TimePeriod
		{
			get
			{
				return ExcelConditionalFormattingTimePeriodType.GetTypeByAttribute(
				  GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TimePeriodAttribute));
			}
			set
			{
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.TimePeriodAttribute,
				  ExcelConditionalFormattingTimePeriodType.GetAttributeByType(value),
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the conditional formatting Operator type.
		/// </summary>
		internal protected eExcelConditionalFormattingOperatorType Operator
		{
			get
			{
				return ExcelConditionalFormattingOperatorType.GetTypeByAttribute(
				  GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.OperatorAttribute));
			}
			set
			{
				SetXmlNodeString(
				  ExcelConditionalFormattingConstants.Paths.OperatorAttribute,
				  ExcelConditionalFormattingOperatorType.GetAttributeByType(value),
				  true);
			}
		}

		/// <summary>
		/// Gets or sets the conditional formatting rule formula.
		/// </summary>
		public string Formula
		{
			get
			{
				return GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.Formula);
			}
			set
			{
				SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.Formula, value);
			}
		}

		/// <summary>
		/// Gets or sets the conditional formatting rule's second formula.
		/// </summary>
		public string Formula2
		{
			get
			{
				return GetXmlNodeString($"{ExcelConditionalFormattingConstants.Paths.Formula}[position()=2]");
			}
			set
			{
				// Create/Get the first <formula> node (ensure that it exists).
				var firstNode = CreateComplexNode(
				  this.TopNode,
				  string.Format(
					 "{0}[position()=1]",
					 // {0}
					 ExcelConditionalFormattingConstants.Paths.Formula));
				// Create/Get the seconde <formula> node (ensure that it exists).
				var secondNode = CreateComplexNode(this.TopNode, $"{ExcelConditionalFormattingConstants.Paths.Formula}[position()=2]");
				// Save the formula in the second <formula> node.
				secondNode.InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the DxfId Style Attribute.
		/// </summary>
		internal int DxfId
		{
			get
			{
				return GetXmlNodeInt(ExcelConditionalFormattingConstants.Paths.DxfIdAttribute);
			}
			set
			{
				SetXmlNodeString(
					ExcelConditionalFormattingConstants.Paths.DxfIdAttribute,
					(value == int.MinValue) ? string.Empty : value.ToString(),
					true);
			}
		}
		#endregion

		#region Private Properties
		private ExcelWorksheet Worksheet { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingRule"/>
		/// </summary>
		/// <param name="type">The enum type of the conditional formatting rule.</param>
		/// <param name="address">The address which this conditional formatting rule applies to.</param>
		/// <param name="priority">Used also as the cfRule unique key</param>
		/// <param name="worksheet">The worksheet that this conditional formatting rule applies to.</param>
		/// <param name="itemElementNode">The cfRule XML node.</param>
		/// <param name="namespaceManager">The <see cref="XmlNamespaceManager"/> for the cfRule.</param>
		internal ExcelConditionalFormattingRule(
		  eExcelConditionalFormattingRuleType type,
		  ExcelAddress address,
		  int priority,
		  ExcelWorksheet worksheet,
		  XmlNode itemElementNode,
		  XmlNamespaceManager namespaceManager)
		  : base(
			 namespaceManager,
			 itemElementNode)
		{
			Require.Argument(address).IsNotNull("address");
			// While MSDN states that 1 is the "highest priority," it also defines this
			// field as W3C XML Schema int, which would allow values less than 1. Excel
			// itself will, on occasion, use a value of 0, so this check will allow a 0.
			Require.Argument(priority).IsInRange(0, int.MaxValue, "priority");
			Require.Argument(worksheet).IsNotNull("worksheet");

			this.myType = type;
			this.Worksheet = worksheet;
			this.SchemaNodeOrder = this.Worksheet.SchemaNodeOrder;

			if (itemElementNode == null)
			{
				// Create/Get the <cfRule> inside <conditionalFormatting>
				itemElementNode = CreateComplexNode(
					this.Worksheet.WorksheetXml.DocumentElement,
				  string.Format(
						"{0}[{1}='{2}']/{1}='{2}'/{3}[{4}='{5}']/{4}='{5}'",
						//{0}
						ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
						// {1}
						ExcelConditionalFormattingConstants.Paths.SqrefAttribute,
						// {2}
						address.AddressSpaceSeparated,
						// {3}
						ExcelConditionalFormattingConstants.Paths.CfRule,
						//{4}
						ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
						//{5}
						priority));
			}

			// Point to <cfRule>
			this.TopNode = itemElementNode;
			this.Address = address;
			this.Priority = priority;
			this.Type = type;
			if (this.DxfId >= 0)
			{
				worksheet.Workbook.Styles.Dxfs[this.DxfId].AllowChange = true;  //This Id is referenced by CF, so we can use it when we save.
				this.myStyle = worksheet.Workbook.Styles.Dxfs[this.DxfId].Clone();    //Clone, so it can be altered without effecting other dxf styles
			}
		}

		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingRule"/>
		/// </summary>
		/// <param name="type">The enum type of the conditional formatting rule.</param>
		/// <param name="address">The address which this conditional formatting rule applies to.</param>
		/// <param name="priority">Used also as the cfRule unique key</param>
		/// <param name="worksheet">The worksheet that this conditional formatting rule applies to.</param>
		/// <param name="namespaceManager">The <see cref="XmlNamespaceManager"/> for the cfRule.</param>
		internal ExcelConditionalFormattingRule(
		  eExcelConditionalFormattingRuleType type,
		  ExcelAddress address,
		  int priority,
		  ExcelWorksheet worksheet,
		  XmlNamespaceManager namespaceManager)
		  : this(
			 type,
			 address,
			 priority,
			 worksheet,
			 null,
			 namespaceManager)
		{
		}
		#endregion Constructors
	}
}